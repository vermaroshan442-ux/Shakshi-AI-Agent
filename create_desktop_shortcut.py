import os
import sys
import shutil
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
APP_EXE = shutil.which("python") or sys.executable
TARGET_SCRIPT = str(APP_DIR / "app.py")
DESKTOP = Path.home() / "Desktop"
ICON_PATH = APP_DIR / "icon.ico"


def create_shortcut():
    if os.name != "nt":
        print("This shortcut helper is for Windows only.")
        return

    try:
        import win32com.client
    except Exception:
        print("Install pywin32 to create a Windows shortcut: pip install pywin32")
        return

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(str(DESKTOP / "Shakshi AI Agent.lnk"))
    shortcut.Targetpath = APP_EXE
    shortcut.Arguments = f'"{TARGET_SCRIPT}"'
    shortcut.WorkingDirectory = str(APP_DIR)
    if ICON_PATH.exists():
        shortcut.IconLocation = str(ICON_PATH)
    shortcut.save()
    print(f"Shortcut created on Desktop: {DESKTOP / 'Shakshi AI Agent.lnk'}")


if __name__ == "__main__":
    create_shortcut()
