import subprocess
import sys
import shlex
import os
import psutil

def is_process_running(process_name):
    """
    Checks if a process with the given name is currently running.
    """
    for proc in psutil.process_iter(['name']):
        if proc.info['name'].lower() == process_name.lower():
            return True
    return False

def open_application(app_name):
    """
    Opens an application based on the operating system, checking if it's already running.
    """
    if is_process_running(app_name):
        print(f"Application '{app_name}' is already running.")
        return

    try:
        if sys.platform == "win32":
            subprocess.Popen([app_name])
        elif sys.platform == "darwin":  # macOS
            subprocess.Popen(["open", "-a", app_name])
        else:  # Linux and other Unix-like OS
            subprocess.Popen(shlex.split(app_name))
        print(f"Opening {app_name}...")
    except FileNotFoundError:
        print(f"Error: Application '{app_name}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

def open_file(filepath):
    """
    Opens a file with its default application.
    """
    try:
        if sys.platform == "win32":
            os.startfile(filepath)
        elif sys.platform == "darwin":  # macOS
            subprocess.Popen(["open", filepath])
        else:  # Linux and other Unix-like OS
            subprocess.Popen(["xdg-open", filepath])
        print(f"Opening {filepath}...")
    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    """
    Main function for the desktop assistant.
    """
    print("Welcome to Nora, your desktop assistant!")
    print("You can say 'open <application>', 'open file <filepath>', or 'exit' to quit.")

    while True:
        command = input(">> ").lower().strip()

        if command == "exit":
            print("Goodbye!")
            break
        elif command.startswith("open file "):
            parts = command.split(" ", 2)
            if len(parts) > 2:
                filepath = parts[2]
                open_file(filepath)
            else:
                print("Please specify a file to open.")
        elif command.startswith("open "):
            parts = command.split(" ", 1)
            if len(parts) > 1:
                app_name = parts[1]
                open_application(app_name)
            else:
                print("Please specify an application to open.")
        else:
            print(f"Unknown command: {command}")

if __name__ == "__main__":
    main()
