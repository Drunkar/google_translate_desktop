# Google Translate Desktop Helper

![top image](top_image.png)

[![Python Version](https://img.shields.io/badge/python-3.x-blue.svg)](https://www.python.org/)

A desktop helper tool for Windows/macOS/Linux that allows you to quickly translate selected text from any application using Google Translate by double-pressing the `Ctrl+C` hotkey.

Translation is performed by displaying the Google Translate webpage in a WebView window using `pywebview`.

## Features

- **System Tray Icon:** Runs in the background with an icon in the system tray (notification area). You can exit the application via the icon's right-click menu.
- **Global Hotkey:** Monitors the `Ctrl+C` key combination system-wide. Triggers when pressed twice within approximately 0.5 seconds.
- **Google Translate Integration:** Opens the selected text (copied to the clipboard) in a WebView window displaying `translate.google.com`.
- **Automatic Language Switching:** Attempts to detect the source language of the input text using the `langdetect` library. If the detected language matches the default target language, it automatically switches the target language to the configured UI language for the translation URL.
- **Window Reuse:** If a translation window is already open, it reuses the existing window and loads the new translation content instead of opening a new one.
- **Force Foreground (Windows):** Attempts to forcefully bring the existing translation window to the foreground when reused, using Windows-specific APIs.

## Requirements

- Python 3.x
- Required Python libraries (see `requirements.txt` for details):
  - `keyboard`
  - `pyperclip`
  - `pystray`
  - `Pillow`
  - `pywebview`
  - `langdetect`
  - `pypiwin32` (Windows only)
- **WebView2 Runtime** (Windows): The default browser engine used by `pywebview` on Windows. Usually included in modern Windows 10/11, but may need to be installed otherwise. ([Download Microsoft Edge WebView2 Runtime](https://developer.microsoft.com/en-us/microsoft-edge/webview2/))
- **Administrator Privileges (Especially on Windows):**
  - Required for the `keyboard` library to monitor key presses in other applications.
  - May be required for the `SetForegroundWindow` API (used to force the window to the front on Windows) to work reliably.

## Installation

1.  **Clone or Download the Repository/Script:**

    ```bash
    git clone <repository_url>
    cd <repository_name>
    ```

    Or download the `google_translate_desktop.py` file directly.

2.  **Install Dependencies:**
    Open a command prompt or terminal, navigate to the directory containing `requirements.txt`, and run:
    ```bash
    pip install -r requirements.txt
    ```
    (This will automatically install `pypiwin32` only on Windows environments).

## Usage

1.  **Run the Script:**
    Open a command prompt or terminal and run the following. **Running with administrator privileges is highly recommended (especially on Windows).**

    ```bash
    python google_translate_desktop.py
    ```

    Logs will be printed to the console, and an icon will appear in the system tray.

2.  **Translate:**

    - Select the text you want to translate in any application.
    - Quickly press `Ctrl+C` twice on your keyboard.
    - A WebView window will open displaying the translation (or the existing window will be updated and brought to the front).

3.  **Exit:**
    Right-click the system tray icon and select "Exit".

## Configuration

You can adjust settings in the "Settings" section at the top of the `google_translate_desktop.py` script:

- `DOUBLE_PRESS_TIMEOUT`: Time window (in seconds) to detect a double-press of the hotkey.
- `TARGET_LANG`: Default target language code (e.g., 'ja', 'en').
- `UI_LANG`: UI language for the Google Translate page, and the alternate target language if the source matches `TARGET_LANG`.
- `MIN_TEXT_LEN_FOR_DETECT`: Minimum text length required to attempt language detection.
- `WINDOW_WIDTH`, `WINDOW_HEIGHT`: Default size of the translation window.
- `VISIBLE_WINDOW_TITLE`: The fixed title used to find the window on Windows for reuse.

## Building the Executable (.exe using PyInstaller)

You can use `PyInstaller` to create a single `.exe` file that can run on Windows environments without requiring Python to be installed.

1.  **Install PyInstaller:**

    ```bash
    pip install pyinstaller
    ```

2.  **Build Command:**
    Navigate to the script's directory in the command prompt or PowerShell and run:

    ```bash
    pyinstaller --noconsole --onefile .\google_translate_desktop.py
    ```

    - `--noconsole`: Prevents the console window from appearing when the `.exe` runs.
    - `--onefile`: Bundles everything into a single executable file.
    - (Optional) `-i <path_to_icon>.ico`: Sets a custom icon for the `.exe` file.

3.  **Important Notes on Packaging:**
    Packaging applications with complex dependencies like `pywebview`, `pystray`, `langdetect`, and `pypiwin32` can be challenging. The basic command above **might not be sufficient** and could result in runtime errors (e.g., "ModuleNotFound") or missing functionality (language detection, icons, WebView not loading).
    You will likely need to generate and edit a `.spec` file for more complex configurations:

    1.  Generate spec: `pyi-makespec --noconsole --onefile .\google_translate_desktop.py`
    2.  Edit `google_translate_desktop.spec`: Manually add entries to `hiddenimports`, `datas`, and `binaries` within the `Analysis` section. Examples you might need:
        - **Hidden Imports:** `hiddenimports=['pystray.backends.win32', 'win32api', 'win32gui', 'win32con', 'pkg_resources.py2_warn']` (actual list may vary)
        - **Data Files:** `datas=[('C:/path/to/venv/Lib/site-packages/langdetect/profiles', 'langdetect/profiles')]` (Find the correct path to `langdetect`'s `profiles` directory in your environment)
        - **Binaries:** `binaries=[('C:/path/to/WebView2Loader.dll', '.')]` (Find and include the correct WebView2 loader DLL if needed)
    3.  Build from spec: `pyinstaller .\google_translate_desktop.spec`
        Finding the correct paths and required files often involves some trial and error based on runtime issues.

4.  **Running the Executable:**
    The final `.exe` file will be located in the `dist` folder after a successful build. Remember that running the `.exe` with **administrator privileges** is recommended for the hotkey and window activation features to work correctly.

## License

(Add your license information here if applicable, e.g., MIT License)
