import os

APP_MAPPINGS = {
    "win32": {
        "chrome": "chrome.exe",
        "firefox": "firefox.exe",
        "vscode": "Code.exe",
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
        "camera": "WindowsCamera.exe",
        "photos": "Microsoft.Photos.exe",
        "settings": "SystemSettings.exe",
        "store": "WinStore.App.exe",
    },
    "darwin": {
        "chrome": "Google Chrome",
        "firefox": "Firefox",
        "vscode": "Visual Studio Code",
        "safari": "Safari",
        "textedit": "TextEdit",
    },
    "linux": {
        "chrome": "google-chrome",
        "firefox": "firefox",
        "vscode": "code",
        "gedit": "gedit",
        "terminal": "gnome-terminal",
    },
}

# Common installation directories
SEARCH_PATHS = {
    "win32": [
        "C:\\Program Files",
        "C:\\Program Files (x86)",
        os.path.join(os.getenv('LOCALAPPDATA'), 'Microsoft\\WindowsApps')
    ],
    "darwin": [
        "/Applications"
    ],
    "linux": [
        "/usr/bin",
        "/usr/local/bin",
        "/snap/bin",
    ]
}
