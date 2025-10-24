import subprocess
import sys
import shlex
import os
import psutil
import webbrowser
import pyttsx3
import speech_recognition as sr
import pvporcupine
import pyaudio
import struct
from dotenv import load_dotenv
from app_mappings import APP_MAPPINGS, SEARCH_PATHS

load_dotenv()

PICOVOICE_ACCESS_KEY = os.getenv("PICOVOICE_ACCESS_KEY")

if sys.platform == "win32":
    import win32gui
    import win32con

engine = pyttsx3.init()
recognizer = sr.Recognizer()

def speak(text):
    """
    Speaks the given text.
    """
    print(f"Nora: {text}")
    engine.say(text)
    engine.runAndWait()

def listen_for_wake_word():
    """
    Listens for the wake word "porcupine".
    """
    if not PICOVOICE_ACCESS_KEY:
        speak("Picovoice access key not found. Please set the PICOVOICE_ACCESS_KEY environment variable.")
        return False

    porcupine = None
    pa = None
    audio_stream = None
    try:
        # NOTE: "Nora" is a custom wake word. Using "porcupine" as a placeholder.
        porcupine = pvporcupine.create(access_key=PICOVOICE_ACCESS_KEY, keywords=['porcupine'])
        pa = pyaudio.PyAudio()
        audio_stream = pa.open(
            rate=porcupine.sample_rate,
            channels=1,
            format=pyaudio.paInt16,
            input=True,
            frames_per_buffer=porcupine.frame_length
        )
        print("Listening for wake word ('porcupine')...")
        while True:
            pcm = audio_stream.read(porcupine.frame_length)
            pcm = struct.unpack_from("h" * porcupine.frame_length, pcm)
            keyword_index = porcupine.process(pcm)
            if keyword_index >= 0:
                print("Wake word detected!")
                return True
    except Exception as e:
        speak(f"An error occurred with the wake word listener: {e}")
        return False
    finally:
        if audio_stream is not None:
            audio_stream.close()
        if pa is not None:
            pa.terminate()
        if porcupine is not None:
            porcupine.delete()


def listen_for_command():
    """
    Listens for a command from the user after the wake word is detected.
    """
    with sr.Microphone() as source:
        print("Listening for command...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

    try:
        command = recognizer.recognize_google(audio)
        print(f"You said: {command}")
        return command.lower()
    except sr.UnknownValueError:
        speak("Sorry, I didn't catch that.")
        return None
    except sr.RequestError:
        speak("Sorry, my speech service is down.")
        return None

# --- (The rest of the functions like find_executable, open_application, etc. remain unchanged) ---

def find_executable(app_name):
    platform = sys.platform
    if platform in APP_MAPPINGS and app_name in APP_MAPPINGS[platform]:
        app_name = APP_MAPPINGS[platform][app_name]
    if os.path.isfile(app_name):
        return app_name
    path = os.environ.get("PATH", "")
    for p in path.split(os.pathsep):
        full_path = os.path.join(p, app_name)
        if os.path.isfile(full_path) and os.access(full_path, os.X_OK):
            return full_path
    if platform in SEARCH_PATHS:
        for p in SEARCH_PATHS[platform]:
            for root, _, files in os.walk(p):
                if app_name in files:
                    return os.path.join(root, app_name)
    return None

def get_process_pid(process_name):
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'].lower() == process_name.lower():
            return proc.info['pid']
    return None

def bring_window_to_front(pid):
    # This is a Windows-only implementation
    if sys.platform != "win32":
        speak("Bringing window to front is only supported on Windows for now.")
        return
    def enum_windows_callback(hwnd, pids):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
            _, found_pid = win32gui.GetWindowThreadProcessId(hwnd)
            if found_pid in pids:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
    win32gui.EnumWindows(enum_windows_callback, [pid])

def open_application(app_name):
    executable_path = find_executable(app_name)
    if not executable_path:
        speak(f"Application '{app_name}' not found. Would you like to search online?")
        response = listen_for_command()
        if response and 'yes' in response:
            search_url = f"https://www.google.com/search?q={app_name}"
            webbrowser.open(search_url)
        return
    process_name = os.path.basename(executable_path)
    pid = get_process_pid(process_name)
    if pid:
        speak(f"Application '{app_name}' is already running. Bringing to front.")
        bring_window_to_front(pid)
        return
    try:
        if sys.platform == "win32":
            subprocess.Popen([executable_path])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-a", executable_path])
        else:
            subprocess.Popen([executable_path])
        speak(f"Opening {app_name}...")
    except Exception as e:
        speak(f"An error occurred: {e}")

def close_application(app_name):
    executable_path = find_executable(app_name)
    if not executable_path:
        speak(f"Application '{app_name}' not found.")
        return
    process_name = os.path.basename(executable_path)
    process_found = False
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'].lower() == process_name.lower():
            try:
                proc.terminate()
                proc.wait(timeout=3)
                speak(f"Successfully closed {app_name}.")
                process_found = True
            except psutil.NoSuchProcess:
                speak(f"{app_name} was already closed.")
            except psutil.TimeoutExpired:
                speak(f"Could not close {app_name} gracefully. Forcing termination.")
                proc.kill()
            except (psutil.AccessDenied, psutil.ZombieProcess):
                speak(f"Could not close {app_name} due to system restrictions.")
            finally:
                process_found = True
    if not process_found:
        speak(f"Application '{app_name}' is not running.")

def open_file(filepath):
    try:
        if sys.platform == "win32":
            os.startfile(filepath)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", filepath])
        else:
            subprocess.Popen(["xdg-open", filepath])
        speak(f"Opening {filepath}...")
    except FileNotFoundError:
        speak(f"Error: File '{filepath}' not found.")
    except Exception as e:
        speak(f"An error occurred: {e}")

def play_on_youtube(query):
    speak(f"Playing {query} on YouTube.")
    search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
    webbrowser.open(search_url)


def main():
    """
    Main function for the desktop assistant.
    """
    speak("Assistant is running.")
    while True:
        if listen_for_wake_word():
            speak("How can I help you?")
            command_str = listen_for_command()
            if command_str:
                try:
                    parts = shlex.split(command_str)
                    command = parts[0].lower()
                    args = parts[1:]

                    if command == "goodbye" or command == "exit":
                        speak("Goodbye!")
                        break
                    elif command == "open" and args:
                        if args[0] == "file" and len(args) > 1:
                            open_file(" ".join(args[1:]))
                        else:
                            open_application(" ".join(args))
                    elif command == "close" and args:
                        close_application(" ".join(args))
                    elif command == "play" and "on" in args and "youtube" in args:
                        query_index = args.index("youtube") + 1
                        if query_index < len(args):
                             query = " ".join(args[query_index:])
                             play_on_youtube(query)
                        else:
                             speak("Please specify what you want to play on YouTube.")
                    else:
                        speak(f"Sorry, I don't know the command: {command_str}")
                except ValueError:
                    speak("I had trouble understanding the command due to quotes.")

if __name__ == "__main__":
    main()
