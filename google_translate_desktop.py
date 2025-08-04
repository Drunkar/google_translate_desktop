import os
import platform  # OS判定用
import sys
import threading
import time
import urllib.parse  # URLエンコード用

import keyboard
import pyperclip
import pystray
import webview  # pywebviewライブラリ
from PIL import Image, ImageDraw
from pystray import MenuItem as item

# 言語検出ライブラリを追加
try:
    from langdetect import LangDetectException, detect

    langdetect_available = True
except ImportError:
    print("警告: 'langdetect' ライブラリが見つかりません。")
    print("pip install langdetect を実行してください。")
    print("言語検出機能は無効になります。")
    langdetect_available = False

import platform

# 条件付きインポート (Windows用を強化)
if platform.system() == "Windows":
    try:
        import win32api  # keybd_eventのため追加
        import win32con
        import win32gui

        print("pywin32 をインポートしました。")
    except ImportError:
        print("警告: pywin32 が見つかりません。'pip install pywin32' が必要です (Windows用強制前面表示)。")
        win32gui = None
        win32api = None
        win32con = None


# --- 設定 ---
DOUBLE_PRESS_TIMEOUT = 0.5  # ダブルプレスとみなす最大間隔（秒）
SOURCE_LANG = "auto"  # Google側の検出は'auto'のまま
TARGET_LANG = "ja"  # デフォルトの翻訳先言語 (日本語)
# TARGET_LANG = 'en'   # 英語に翻訳する場合
UI_LANG = "en"  # UI言語 (英語) - TARGET_LANGと同じ場合はこちらに切り替える
MIN_TEXT_LEN_FOR_DETECT = 10  # 言語検出を試みる最小文字数
RESTART_INTERVAL_HOURS = 12  # アプリケーションを再起動する間隔（時間）

# URLのベース部分 (tl は動的に決定)
BASE_TRANSLATE_URL_TEMPLATE = (
    "https://translate.google.com/?hl={ui_lang}&sl={source_lang}&tl={target_lang}&op=translate&text={encoded_text}"
)

# WebViewウィンドウのデフォルトサイズ
WINDOW_WIDTH = 1200
WINDOW_HEIGHT = 600

VISIBLE_WINDOW_TITLE = "Google Translate Helper"  # このタイトルでウィンドウを探す
# ----------------

# --- グローバル変数 ---
last_ctrl_c_time = 0
icon_instance = None
hidden_window = None  # 非表示のバックエンドウィンドウ
app_quit = False
_RESTART_SCHEDULED = False  # 定期再起動がスケジュールされたかどうかのフラグ
api_instance = None  # Apiインスタンスをグローバルで保持できるようにする


# --- タスクトレイアイコン関連 ---
# (変更なし: create_image, exit_action, run_pystray)
def create_image(width, height, color1, color2):
    image = Image.new("RGB", (width, height), color1)
    dc = ImageDraw.Draw(image)
    dc.rectangle((width // 2, 0, width, height // 2), fill=color2)
    dc.rectangle((0, height // 2, width // 2, height), fill=color2)
    return image


def exit_action(icon, item):
    global app_quit, hidden_window
    print("終了処理を開始...")
    app_quit = True
    if hidden_window:
        print("非表示WebViewウィンドウの破棄を試みます...")
        try:
            hidden_window.evaluate_js("window.pywebview.api.shutdown_app()")
        except Exception as e:
            print(f"WebView終了中のエラー: {e}")
            print("強制終了します。")
            os._exit(1)
    else:
        print("WebView未初期化。プロセスを終了します。")
        os._exit(0)


def run_pystray():
    global icon_instance
    print("pystray スレッド開始")
    image = create_image(64, 64, "#d2d2d3", "#5094fd")  # 色変更
    menu = (item("終了", exit_action),)
    # タイトルバーに合わせて変更 (例: Google Translate Helper)
    icon = pystray.Icon("google_translate_helper", image, "Google Translate Helper", menu)
    icon_instance = icon
    icon.run()
    print("pystray スレッド終了")


# --- キーボードリスナー (別スレッドで実行) ---
# (変更なし: listen_keyboard, on_ctrl_c, escape_js_string)
def listen_keyboard():
    global app_quit
    print("キーボードリスナースレッド開始")

    while not app_quit:
        # --- 1. 古いホットキーの解除を試みる ---
        try:
            # ループの最初に、前回のループで登録したホットキーを名前で解除する
            # これにより、コールバックの重複を防ぐ
            keyboard.remove_hotkey("ctrl+c")
            print("古い 'ctrl+c' ホットキーを正常に解除しました。")
        except KeyError:
            # 初回起動時など、まだホットキーが登録されていない場合はKeyErrorが発生するが、これは正常な動作
            print("解除対象のホットキーはありませんでした（初回実行）。")
        except Exception as e:
            # その他の予期せぬエラー
            print(f"ホットキーの解除中にエラーが発生: {e}")

        # --- 2. 新しいホットキーを登録する ---
        try:
            keyboard.add_hotkey("ctrl+c", on_ctrl_c, trigger_on_release=False)
            print("新しい 'ctrl+c' ホットキーをセットしました。")
        except Exception as e:
            print(f"ホットキーの登録に失敗: {e}")
            # 登録に失敗した場合、少し待ってからループを再試行
            time.sleep(60)
            continue  # ループの先頭に戻る

        # --- 3. 次の再登録まで待機 ---
        # 10分間(600秒)待機。スリープ中も時間は経過する。
        # 1秒ごとに終了フラグを確認し、速やかに終了できるようにする
        for _ in range(600):
            if app_quit:
                break
            time.sleep(1)

    # --- 最終的なクリーンアップ ---
    try:
        print("キーボードリスナーループ終了。全てのフックを解除します。")
        keyboard.unhook_all()
    except Exception as e:
        print(f"最終的なフック解除中にエラー: {e}")


def on_ctrl_c():
    global last_ctrl_c_time, hidden_window, app_quit
    if app_quit:
        return
    current_time = time.time()
    if current_time - last_ctrl_c_time < DOUBLE_PRESS_TIMEOUT:
        print("Ctrl+C ダブルプレス検出！")
        time.sleep(0.1)
        try:
            clipboard_content = pyperclip.paste()
            if clipboard_content and clipboard_content.strip():
                if hidden_window:
                    js_text = escape_js_string(clipboard_content)
                    print(f"メインスレッドに関数呼び出し依頼: {clipboard_content[:30]}...")
                    # API関数名を変更 create_or_focus_translate_window にする
                    hidden_window.evaluate_js(f'window.pywebview.api.create_or_focus_translate_window("{js_text}")')
                else:
                    print("エラー: 非表示ウィンドウ未初期化。")
            elif clipboard_content is None:
                print("クリップボードアクセス失敗(None)。")
            else:
                print("クリップボード空。")
        except Exception as e:
            print(f"クリップボード/JS評価エラー: {e}")
        last_ctrl_c_time = 0
    else:
        last_ctrl_c_time = current_time


def escape_js_string(text):
    return (
        text.replace("\\", "\\\\")
        .replace("'", "\\'")
        .replace('"', '\\"')
        .replace("\n", "\\n")
        .replace("\r", "")
        .replace("&", "\\u0026")
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
    )


# --- WebView API (メインスレッドで実行される) ---
class Api:
    def __init__(self):
        self.visible_window = None
        print("Api instance created.")

    def _handle_window_closed(self):
        current_ref = self.visible_window
        if current_ref:
            print(f"表示用ウィンドウ closed. Clearing ref.")
        self.visible_window = None

    def create_or_focus_translate_window(self, text):
        """表示用ウィンドウが存在すれば再利用し強制前面表示、なければ作成する"""
        global app_quit, hidden_window
        if app_quit:
            return

        print(f"API呼び出し受信 (メインスレッド): {text[:30]}...")
        effective_target_lang = TARGET_LANG

        # --- 言語検出ロジック (変更なし) ---
        if langdetect_available:
            try:
                if text and len(text.strip()) >= MIN_TEXT_LEN_FOR_DETECT:
                    detected_lang = detect(text)
                    print(f"Detected source language: {detected_lang}")
                    if detected_lang == TARGET_LANG and UI_LANG != TARGET_LANG:
                        print(f"ソース({detected_lang})==ターゲット({TARGET_LANG}) -> UI言語({UI_LANG})へ変更")
                        effective_target_lang = UI_LANG
            except LangDetectException:
                print(f"言語検出失敗 -> デフォルトターゲット")
            except Exception as e_detect:
                print(f"言語検出エラー: {e_detect} -> デフォルトターゲット")
        # (言語検出ロジックここまで)
        print(f"最終的なターゲット言語: {effective_target_lang}")

        try:
            encoded_text = urllib.parse.quote(text)
            final_url = BASE_TRANSLATE_URL_TEMPLATE.format(
                ui_lang=UI_LANG, source_lang=SOURCE_LANG, target_lang=effective_target_lang, encoded_text=encoded_text
            )
            print(f"目標URL (max 100): {final_url[:100]}...")
            if len(final_url) > 2000:
                print("警告: URL長 > 2000")

            # --- ウィンドウの再利用可能性をタイトルでチェック (Windowsのみ) ---
            window_to_reuse_hwnd = 0  # HWNDを初期化
            can_reuse = False  # 再利用フラグ

            if self.visible_window and platform.system() == "Windows" and win32gui:
                try:
                    # 固定タイトルでウィンドウを検索
                    window_to_reuse_hwnd = win32gui.FindWindow(None, VISIBLE_WINDOW_TITLE)
                    if window_to_reuse_hwnd != 0:
                        print(
                            f"既存ウィンドウをタイトルで発見 (HWND: {window_to_reuse_hwnd})。内部参照も有効なため再利用します。"
                        )
                        # オプション: 内部参照のHWNDと一致するか確認 (ただしhwnd属性の取得が不確実)
                        # ref_hwnd = getattr(self.visible_window, 'hwnd', 0)
                        # if ref_hwnd != 0 and ref_hwnd != window_to_reuse_hwnd:
                        #    print(f"警告: 内部参照HWND({ref_hwnd})とFindWindow({window_to_reuse_hwnd})が不一致。新規作成します。")
                        # else:
                        can_reuse = True  # FindWindowで見つかればOKとする
                    else:
                        print("既存ウィンドウはタイトルで見つかりませんでした。内部参照をクリアします。")
                        self.visible_window = None  # FindWindowで見つからないなら参照は無効
                except Exception as e_find:
                    print(f"FindWindow実行中にエラー: {e_find}。新規作成します。")
                    self.visible_window = None
                    window_to_reuse_hwnd = 0
            elif self.visible_window:
                # Windows以外、またはpywin32がない場合、内部参照だけで判断（不確実性あり）
                try:
                    _ = self.visible_window.uid  # 簡単な生存確認
                    print("既存ウィンドウの内部参照は有効です (Windows以外/pywin32なし)。show()で再利用試行。")
                    can_reuse = True
                except:
                    print("既存ウィンドウ参照は無効でした。")
                    self.visible_window = None

            # --- Reuse or Create ---
            if can_reuse and self.visible_window:
                # --- 既存ウィンドウを再利用 ---
                window_to_use = self.visible_window
                try:
                    print(f"既存ウィンドウ (UID: {window_to_use.uid}) を再利用します。")
                    # Load URL
                    print(f"既存ウィンドウにURLをロード: {final_url[:100]}...")
                    window_to_use.load_url(final_url)
                    # Set dynamic title
                    window_to_use.set_title(VISIBLE_WINDOW_TITLE)

                    # Force front (WindowsはHWND、その他はshow())
                    activated = False
                    print("既存ウィンドウを最前面表示します...")
                    if (
                        platform.system() == "Windows"
                        and win32gui
                        and win32api
                        and win32con
                        and window_to_reuse_hwnd != 0
                    ):
                        try:
                            print(f"HWND {window_to_reuse_hwnd} を使って ALT+SetForegroundWindowを実行します...")
                            win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
                            time.sleep(0.05)
                            win32gui.SetForegroundWindow(window_to_reuse_hwnd)
                            time.sleep(0.05)
                            win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
                            print("ALT+SetForegroundWindow を試行しました。")
                            activated = True
                        except Exception as e_win:
                            print(f"Windows API (ALT+SetForegroundWindow) 呼び出しエラー: {e_win}")
                            try:
                                win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
                            except:
                                pass
                    if not activated:
                        print("OS固有API失敗または対象外のため show() を試みます。")
                        window_to_use.show()

                except Exception as e_reuse:
                    print(f"既存ウィンドウの再利用処理中にエラー: {e_reuse}")
                    self.visible_window = None  # Clear reference if reuse fails
            else:
                # --- 新規ウィンドウを作成 ---
                print("新しい翻訳ウィンドウを作成します...")
                try:
                    # *** 固定タイトルで作成 ***
                    webview.create_window(
                        VISIBLE_WINDOW_TITLE,  # 固定タイトルを使用
                        final_url,
                        width=WINDOW_WIDTH,
                        height=WINDOW_HEIGHT,
                        resizable=True,
                        confirm_close=False,
                    )
                    print("新規ウィンドウ作成指示完了。参照取得試行...")
                    # 参照取得とイベントハンドラ設定 (以前と同様、不安定な可能性あり)
                    time.sleep(0.3)
                    all_windows = webview.windows
                    visible_windows = [w for w in all_windows if w != hidden_window]
                    if visible_windows:
                        newly_created_window = visible_windows[-1]
                        # オプション: 作成直後に動的タイトルに設定し直す
                        # dynamic_title = f"翻訳: {text[:20]}..."
                        # newly_created_window.set_title(dynamic_title)
                        self.visible_window = newly_created_window
                        print(f"新しいウィンドウへの参照を取得 (UID: {self.visible_window.uid})。")
                        self.visible_window.events.closed += self._handle_window_closed
                        print("閉じるイベントハンドラ設定。")
                    else:
                        print("エラー: 新規作成された表示用ウィンドウが見つかりません。")
                        self.visible_window = None
                except Exception as e_create:
                    print(f"新規ウィンドウ作成/参照取得エラー: {e_create}")
                    self.visible_window = None

        except Exception as e_outer:
            print(f"URL構築/ウィンドウ処理外部エラー: {e_outer}")

    def schedule_restart_flag(self):
        """定期再起動のフラグを立て、シャットダウン処理を開始する"""
        global _RESTART_SCHEDULED, app_quit
        print("API: 定期再起動フラグがセットされました。")
        _RESTART_SCHEDULED = True
        # app_quit = True # shutdown_appの中でセットされるのでここでは不要かも
        self.shutdown_app()  # 通常のシャットダウン処理を呼び出す

    def shutdown_app(self):
        global app_quit, hidden_window
        print("API経由でシャットダウン要求受信。")
        app_quit = True  # 全スレッドに終了を通知
        if hidden_window:
            try:
                print("非表示WebViewウィンドウの破棄を試みます (shutdown_app)...")
                hidden_window.destroy()
            except Exception as e:
                print(f"非表示ウィンドウ破棄エラー (shutdown_app): {e}")
        print("メインスレッド終了を試みます (shutdown_app)...")


def restart_scheduler_thread(api_ref):
    """指定時間後にAPI経由で再起動をスケジュールするスレッド"""
    global app_quit
    try:
        sleep_seconds = RESTART_INTERVAL_HOURS * 60 * 60
        print(f"再起動スケジューラースレッド開始。{RESTART_INTERVAL_HOURS}時間後に再起動を試みます。")

        # app_quit が True になるまで待機するか、指定時間が経過するまで待機
        # これにより、通常の終了操作が行われた場合は、このスレッドも速やかに終了する
        for _ in range(sleep_seconds):
            if app_quit:
                print("再起動スケジューラー: アプリケーション終了シグナル受信。再起動せずに終了します。")
                return
            time.sleep(1)

        if not app_quit:  # 通常終了ではなく、時間が来た場合
            print(f"再起動スケジューラー: {RESTART_INTERVAL_HOURS}時間が経過しました。再起動処理を開始します。")
            if api_ref:
                api_ref.schedule_restart_flag()
            else:
                print("再起動スケジューラー: APIインスタンスが無効なため、再起動できません。")

    except Exception as e:
        print(f"再起動スケジューラースレッドでエラー: {e}")


# --- メイン処理 と import, 設定 など ---
# (Apiクラス以外の部分は前回のコードと同じなので省略)
# ... (import keyboard, pyperclip, etc.) ...
# ... (設定) ...
# ... (グローバル変数) ...
# ... (タスクトレイ関連: create_image, exit_action, run_pystray) ...
# ... (キーボードリスナー関連: listen_keyboard, on_ctrl_c, escape_js_string) ...

# --- メイン処理 (メインスレッド) ---
if __name__ == "__main__":
    print("WebView Google翻訳ヘルパー (ウィンドウ再利用/一時的最前面版) を開始します。")
    if not langdetect_available:
        print("!!! 言語検出機能は無効です。'langdetect'をインストールしてください。 !!!")

    tray_thread = threading.Thread(target=run_pystray, daemon=True)
    tray_thread.start()

    keyboard_thread = threading.Thread(target=listen_keyboard, daemon=True)
    keyboard_thread.start()

    print("メインスレッドでWebViewを初期化します...")
    api_instance = Api()  # Apiインスタンスを作成
    try:
        hidden_window = webview.create_window(
            "WebView Helper Backend", html="<html><body>Helper Running</body></html>", hidden=True, js_api=api_instance
        )
        print("WebViewイベントループを開始します。")

        # 再起動スケジューラースレッドを開始 (Apiインスタンスを渡す)
        # hidden_windowが作成された後に開始する
        scheduler_thread = threading.Thread(
            target=restart_scheduler_thread, args=(api_instance,), name="RestartSchedulerThread", daemon=True
        )
        scheduler_thread.start()

        webview.start(debug=False)
        print("WebViewイベントループが終了しました。")

    except Exception as e:
        print(f"\n*** WebView初期化/実行エラー: {e} ***")
        app_quit = True  # エラー発生時もフラグを立てる
        if icon_instance:  # アイコンがあれば停止を試みる
            icon_instance.stop()
        os._exit(1)  # エラー終了

    # --- ここから WebView イベントループ終了後の処理 ---
    print("アプリケーション終了処理の最終段階に入ります...")
    app_quit = True  # 念のため再度フラグを立て、全スレッドに終了を促す

    # キーボードスレッドとトレイアイコンの明示的な終了待ちや停止処理
    # listen_keyboardスレッドはapp_quitを見てループを抜ける
    # run_pystrayスレッドのicon.run()はicon_instance.stop()で抜ける
    print("バックグラウンドスレッドの終了を待ちます...")
    if keyboard_thread.is_alive():
        print("キーボードスレッドがまだ動作中。終了を待ちます。")
        keyboard_thread.join(timeout=5)  # 最大5秒待つ
    if tray_thread.is_alive():
        print("トレイアイコンスレッドがまだ動作中。アイコンを停止し、終了を待ちます。")
        if icon_instance:
            icon_instance.stop()
        tray_thread.join(timeout=5)  # 最大5秒待つ

    # キーボードフックの最終的な解除
    try:
        print("最終的なキーボードフック解除を試みます。")
        keyboard.unhook_all()
    except Exception as e_unhook:
        print(f"最終的なキーボードフック解除エラー: {e_unhook}")

    if _RESTART_SCHEDULED:
        print(f"定期再起動を実行します。次のコマンドで再起動: {sys.executable} {sys.argv}")
        time.sleep(1)  # 念のため少し待つ
        try:
            # 自分自身を新しいプロセスで実行し、現在のプロセスを置き換える
            os.execv(sys.executable, [sys.executable] + sys.argv)
        except Exception as e_exec:
            print(f"再起動 (os.execv) に失敗しました: {e_exec}")
            os._exit(1)  # 再起動に失敗したらエラー終了
    else:
        print("通常のアプリケーション終了処理が完了しました。")
        os._exit(0)  # 通常終了
