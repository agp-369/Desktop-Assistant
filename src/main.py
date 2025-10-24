import subprocess
import sys
import shlex
import os
import psutil
import webbrowser
from app_mappings import APP_MAPPINGS, SEARCH_PATHS

if sys.platform == "win32":
    import win32gui
    import win32con

def find_executable(app_name):
    """
    Finds the full path of an executable.
    """
    # Check for alias
    platform = sys.platform
    if platform in APP_MAPPINGS and app_name in APP_MAPPINGS[platform]:
        app_name = APP_MAPPINGS[platform][app_name]

    # Check if it's already a full path
    if os.path.isfile(app_name):
        return app_name

    # Check in system PATH
    path = os.environ.get("PATH", "")
    for p in path.split(os.pathsep):
        full_path = os.path.join(p, app_name)
        if os.path.isfile(full_path) and os.access(full_path, os.X_OK):
            return full_path

    # Search in common directories
    if platform in SEARCH_PATHS:
        for p in SEARCH_PATHS[platform]:
            for root, _, files in os.walk(p):
                if app_name in files:
                    return os.path.join(root, app_name)
    return None

def get_process_pid(process_name):
    """
    Gets the PID of a running process.
    """
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'].lower() == process_name.lower():
            return proc.info['pid']
    return None

def bring_window_to_front(pid):
    """
    Brings a window to the front on Windows.
    """
    if sys.platform != "win32":
        print("Bringing window to front is only supported on Windows for now.")
        return

    def enum_windows_callback(hwnd, pids):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
            _, found_pid = win32gui.GetWindowThreadProcessId(hwnd)
            if found_pid in pids:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)

    win32gui.EnumWindows(enum_windows_callback, [pid])

def open_application(app_name):
    """
    Opens an application, bringing it to the front if already running.
    """
    executable_path = find_executable(app_name)

    if not executable_path:
        response = input(f"Application '{app_name}' not found. Search online? (yes/no) ").lower().strip()
        if response == 'yes':
            search_url = f"https://www.google.com/search?q={app_name}"
            webbrowser.open(search_url)
        return

    process_name = os.path.basename(executable_path)
    pid = get_process_pid(process_name)
    if pid:
        print(f"Application '{app_name}' is already running. Bringing to front.")
        bring_window_to_front(pid)
        return

    try:
        if sys.platform == "win32":
            subprocess.Popen([executable_path])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-a", executable_path])
        else:
            subprocess.Popen([executable_path])
        print(f"Opening {app_name}...")
    except Exception as e:
        print(f"An error occurred: {e}")

def close_application(app_name):
    """
    Closes a running application by its name.
    """
    executable_path = find_executable(app_name)
    if not executable_path:
        print(f"Application '{app_name}' not found.")
        return

    process_name = os.path.basename(executable_path)
    process_found = False
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'].lower() == process_name.lower():
            try:
                proc.terminate()
                proc.wait(timeout=3)
                print(f"Successfully closed {app_name}.")
                process_found = True
            except psutil.NoSuchProcess:
                print(f"{app_name} was already closed.")
            except psutil.TimeoutExpired:
                print(f"Could not close {app_name} gracefully. Forcing termination.")
                proc.kill()
                process_found = True
            except (psutil.AccessDenied, psutil.ZombieProcess):
                print(f"Could not close {app_name} due to system restrictions.")
                process_found = True

    if not process_found:
        print(f"Application '{app_name}' is not running.")

def open_file(filepath):
    """
    Opens a file with its default application.
    """
    try:
        if sys.platform == "win32":
            os.startfile(filepath)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", filepath])
        else:
            subprocess.Popen(["xdg-open", filepath])
        print(f"Opening {filepath}...")
    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

def play_on_youtube(query):
    """
    Searches for and plays a video on YouTube.
    """
    search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
    webbrowser.open(search_url)

def main():
    """
    Main function for the desktop assistant.
    """
    print("Welcome to Nora, your desktop assistant!")
    print("You can say 'open <app>', 'close <app>', 'open file <path>', 'play on youtube <query>', or 'exit'.")

    while True:
        try:
            command_str = input(">> ").strip()
            if not command_str:
                continue

            parts = shlex.split(command_str)
            command = parts[0].lower()
            args = parts[1:]

            if command == "exit":
                print("Goodbye!")
                break
            elif command == "open" and args:
                if args[0] == "file" and len(args) > 1:
                    open_file(" ".join(args[1:]))
                else:
                    open_application(" ".join(args))
            elif command == "close" and args:
                close_application(" ".join(args))
            elif command == "play" and len(args) > 2 and " ".join(args[:2]) == "on youtube":
                play_on_youtube(" ".join(args[2:]))
            else:
                print(f"Unknown command: {command_str}")
        except ValueError:
            print("Error: Unmatched quotes in command.")

if __name__ == "__main__":
    main()
