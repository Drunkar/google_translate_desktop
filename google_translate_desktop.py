import atexit
import os
import platform
import subprocess
import sys
import tempfile
import threading
import time
import urllib.parse

import keyboard
import pyperclip
import pystray
import webview
from PIL import Image, ImageDraw
from pystray import MenuItem as item

# --- Optional language detection (soft dependency) ---
try:
    from langdetect import LangDetectException, detect

    LANGDETECT_OK = True
except Exception:
    LANGDETECT_OK = False

# --- Optional Win32 frontmost support ---
if platform.system() == "Windows":
    try:
        import win32api
        import win32con
        import win32gui
    except Exception:
        win32api = win32con = win32gui = None

# =========================
# Settings
# =========================
DOUBLE_PRESS_TIMEOUT = 0.5
SOURCE_LANG = "auto"
TARGET_LANG = "ja"
UI_LANG = "en"  # Google Translate UI language
MIN_TEXT_LEN_FOR_DETECT = 10
RESTART_INTERVAL_HOURS = 12  # periodic restart
WINDOW_WIDTH = 1200
WINDOW_HEIGHT = 600
VISIBLE_WINDOW_TITLE = "Google Translate Helper"

# =========================
# Globals
# =========================
last_ctrl_c_time = 0.0
icon_instance = None
hidden_window = None
app_quit = False
_RESTART_SCHEDULED = False
api_instance = None
_single_instance_lock = None


# =========================
# Helpers
# =========================
def create_image(width, height, color1, color2):
    """Tray icon image."""
    image = Image.new("RGB", (width, height), color1)
    dc = ImageDraw.Draw(image)
    dc.rectangle((width // 2, 0, width, height // 2), fill=color2)
    dc.rectangle((0, height // 2, width // 2, height), fill=color2)
    return image


def escape_js_string(text: str) -> str:
    """Basic JS-string escaping."""
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


def _quote_args(args):
    return " ".join(['"{}"'.format(a) for a in args])


# =========================
# Tray
# =========================
def exit_action(icon, item):
    """Tray [Exit]. Ask app to shutdown cleanly."""
    global app_quit, hidden_window
    app_quit = True
    try:
        if hidden_window:
            hidden_window.evaluate_js("window.pywebview.api.shutdown_app()")
    except Exception:
        pass


def run_pystray():
    global icon_instance
    img = create_image(64, 64, "#d2d2d3", "#5094fd")
    menu = (item("終了", exit_action),)
    icon = pystray.Icon("google_translate_helper", img, "Google Translate Helper", menu)
    icon_instance = icon
    icon.run()


# =========================
# Keyboard
# =========================
def on_ctrl_c():
    """Double-press Ctrl+C → translate clipboard text."""
    global last_ctrl_c_time, app_quit, hidden_window
    if app_quit:
        return
    now = time.time()
    if now - last_ctrl_c_time < DOUBLE_PRESS_TIMEOUT:
        try:
            time.sleep(0.08)  # let OS copy finish
            cb = pyperclip.paste()
            if cb and cb.strip() and hidden_window:
                hidden_window.evaluate_js(
                    f'window.pywebview.api.create_or_focus_translate_window("{escape_js_string(cb)}")'
                )
        except Exception as e:
            print(f"clipboard/js error: {e}")
        finally:
            last_ctrl_c_time = 0.0
    else:
        last_ctrl_c_time = now


def listen_keyboard():
    """Keep a living hotkey; re-register periodically just in case."""
    global app_quit
    while not app_quit:
        try:
            keyboard.remove_hotkey("ctrl+c")
        except Exception:
            pass
        try:
            keyboard.add_hotkey("ctrl+c", on_ctrl_c, trigger_on_release=False)
        except Exception as e:
            print(f"hotkey add failed: {e}")
            time.sleep(60)
            continue
        for _ in range(600):
            if app_quit:
                break
            time.sleep(1)
    try:
        keyboard.unhook_all()
    except Exception:
        pass


# =========================
# WebView API
# =========================
class Api:
    def __init__(self):
        self.visible_window = None

    def _handle_window_closed(self):
        self.visible_window = None

    def create_or_focus_translate_window(self, text):
        """Create or reuse a visible window and navigate."""
        if app_quit:
            return
        effective_target_lang = TARGET_LANG

        # auto flip target for JP text
        if LANGDETECT_OK:
            try:
                if text and len(text.strip()) >= MIN_TEXT_LEN_FOR_DETECT:
                    detected_lang = detect(text)
                    if detected_lang == TARGET_LANG and UI_LANG != TARGET_LANG:
                        effective_target_lang = UI_LANG
            except LangDetectException:
                pass
            except Exception:
                pass

        encoded_text = urllib.parse.quote(text)
        url = (
            "https://translate.google.com/"
            f"?hl={UI_LANG}&sl={SOURCE_LANG}&tl={effective_target_lang}&op=translate&text={encoded_text}"
        )

        # try reuse
        reuse_hwnd = 0
        can_reuse = False
        if self.visible_window and platform.system() == "Windows" and win32gui:
            try:
                reuse_hwnd = win32gui.FindWindow(None, VISIBLE_WINDOW_TITLE)
                if reuse_hwnd:
                    can_reuse = True
                else:
                    self.visible_window = None
            except Exception:
                self.visible_window = None
        elif self.visible_window:
            try:
                _ = self.visible_window.uid
                can_reuse = True
            except Exception:
                self.visible_window = None

        if can_reuse and self.visible_window:
            w = self.visible_window
            try:
                w.load_url(url)
                w.set_title(VISIBLE_WINDOW_TITLE)
                activated = False
                if platform.system() == "Windows" and win32gui and win32api and win32con and reuse_hwnd:
                    try:
                        win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
                        time.sleep(0.05)
                        win32gui.SetForegroundWindow(reuse_hwnd)
                        time.sleep(0.05)
                        win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
                        activated = True
                    except Exception:
                        try:
                            win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
                        except Exception:
                            pass
                if not activated:
                    w.show()
            except Exception as e:
                print(f"reuse failed: {e}")
                self.visible_window = None
        else:
            try:
                webview.create_window(
                    VISIBLE_WINDOW_TITLE,
                    url,
                    width=WINDOW_WIDTH,
                    height=WINDOW_HEIGHT,
                    resizable=True,
                    confirm_close=False,
                )
                time.sleep(0.3)
                # pick last visible (not the hidden backend)
                visible = [w for w in webview.windows if w is not hidden_window]
                if visible:
                    self.visible_window = visible[-1]
                    self.visible_window.events.closed += self._handle_window_closed
            except Exception as e:
                print(f"create window failed: {e}")
                self.visible_window = None

    def schedule_restart_flag(self):
        """Flag periodic restart and trigger shutdown."""
        global _RESTART_SCHEDULED, app_quit
        _RESTART_SCHEDULED = True
        self.shutdown_app()

    def shutdown_app(self):
        """Request app shutdown (called from tray or scheduler)."""
        global app_quit, hidden_window
        app_quit = True
        try:
            if hidden_window:
                hidden_window.destroy()
        except Exception:
            pass


# =========================
# Clean exit and relaunch
# =========================
def relaunch_after_exit():
    """
    Relaunch after this process *fully* exits to avoid _MEI race.
    On Windows onefile: write a self-deleting .cmd that waits for our PID to vanish,
    then starts a new instance. On others: execv.
    """
    try:
        if platform.system() == "Windows" and getattr(sys, "frozen", False):
            exe = sys.executable
            args = sys.argv[1:]
            parent_pid = os.getpid()
            fd, bat_path = tempfile.mkstemp(prefix="relaunch_", suffix=".cmd")
            os.close(fd)
            script = f"""@echo off
setlocal ENABLEDELAYEDEXPANSION
set PID={parent_pid}
:waitloop
tasklist /FI "PID eq %PID%" | find "%PID%" >nul
if not errorlevel 1 (
  timeout /t 1 /nobreak >nul
  goto waitloop
)
start "" "{exe}" {_quote_args(args)}
del "%~f0"
"""
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(script)

            flags = 0
            if hasattr(subprocess, "DETACHED_PROCESS"):
                flags |= subprocess.DETACHED_PROCESS
            if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP"):
                flags |= subprocess.CREATE_NEW_PROCESS_GROUP
            subprocess.Popen(["cmd.exe", "/c", bat_path], creationflags=flags, close_fds=True)
        else:
            executable = sys.executable
            os.execv(executable, [executable] + sys.argv[1:])
    except Exception as e:
        print(f"relaunch failed: {e}")


def cleanup_on_exit():
    """Release hooks and windows so PyInstaller can remove _MEI cleanly."""
    global icon_instance, hidden_window
    try:
        keyboard.unhook_all()
    except Exception:
        pass
    try:
        if icon_instance:
            icon_instance.stop()
    except Exception:
        pass
    try:
        if hidden_window:
            hidden_window.destroy()
    except Exception:
        pass


# =========================
# Single-instance guard
# =========================
def _acquire_single_instance_lock():
    """
    Create a simple file lock under %LOCALAPPDATA% (Windows) / ~/.cache (others).
    Prevents double-launch during restart/cleanup which can break _MEI.
    """
    import errno
    import pathlib

    base = (
        os.environ.get("LOCALAPPDATA")
        if platform.system() == "Windows"
        else os.path.join(pathlib.Path.home(), ".cache")
    )
    lock_dir = os.path.join(base, "GoogleTranslateHelper")
    os.makedirs(lock_dir, exist_ok=True)
    lock_path = os.path.join(lock_dir, "app.lock")
    if platform.system() == "Windows":
        import msvcrt

        f = open(lock_path, "a+")
        try:
            msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            f.close()
            return None
        return (f, "msvcrt")
    else:
        import fcntl

        f = open(lock_path, "a+")
        try:
            fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            f.close()
            return None
        return (f, "fcntl")


def _release_single_instance_lock(token):
    if not token:
        return
    f, kind = token
    try:
        if platform.system() == "Windows":
            import msvcrt

            try:
                msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
            except Exception:
                pass
        else:
            import fcntl

            try:
                fcntl.flock(f, fcntl.LOCK_UN)
            except Exception:
                pass
    finally:
        try:
            f.close()
        except Exception:
            pass


# =========================
# Restart scheduler thread
# =========================
def restart_scheduler_thread(api_ref):
    secs = RESTART_INTERVAL_HOURS * 3600
    for _ in range(secs):
        if app_quit:
            return
        time.sleep(1)
    if not app_quit and api_ref:
        api_ref.schedule_restart_flag()


# =========================
# Main
# =========================
if __name__ == "__main__":
    # Paranoia: avoid CWD under _MEIPASS
    try:
        if getattr(sys, "_MEIPASS", None):
            os.chdir(os.path.expanduser("~"))
    except Exception:
        pass

    # Single-instance
    _single_instance_lock = _acquire_single_instance_lock()
    if _single_instance_lock is None:
        print("Already running. Exiting.")
        sys.exit(0)

    atexit.register(cleanup_on_exit)

    # Tray + keyboard threads
    t_tray = threading.Thread(target=run_pystray, daemon=True, name="Tray")
    t_tray.start()
    t_key = threading.Thread(target=listen_keyboard, daemon=True, name="Hotkey")
    t_key.start()

    # Hidden backend window with JS API
    api_instance = Api()
    try:
        hidden_window = webview.create_window(
            "WebView Helper Backend",
            html="<html><body>Helper Running</body></html>",
            hidden=True,
            js_api=api_instance,
        )

        # Restart timer
        t_restart = threading.Thread(
            target=restart_scheduler_thread, args=(api_instance,), daemon=True, name="RestartScheduler"
        )
        t_restart.start()

        # GUI loop (must be in main thread on Windows)
        webview.start(debug=False)
    except Exception as e:
        print(f"webview error: {e}")
        app_quit = True
        try:
            if icon_instance:
                icon_instance.stop()
        except Exception:
            pass
        _release_single_instance_lock(_single_instance_lock)
        sys.exit(1)

    # After GUI loop ends
    app_quit = True
    try:
        if t_key.is_alive():
            t_key.join(timeout=5)
        if t_tray.is_alive():
            try:
                if icon_instance:
                    icon_instance.stop()
            except Exception:
                pass
            t_tray.join(timeout=5)
    except Exception:
        pass

    # Relaunch safely if scheduled
    if _RESTART_SCHEDULED:
        relaunch_after_exit()
        _release_single_instance_lock(_single_instance_lock)
        sys.exit(0)
    else:
        _release_single_instance_lock(_single_instance_lock)
        sys.exit(0)
