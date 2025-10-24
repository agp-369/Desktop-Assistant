import subprocess
import sys
import os
import psutil
import webbrowser
import pyttsx3
import speech_recognition as sr
import pvporcupine
import pyaudio
import struct
import pyautogui
import json
from dotenv import load_dotenv
from app_mappings import APP_MAPPINGS, SEARCH_PATHS
from command_parser import parse_command

load_dotenv()

PICOVOICE_ACCESS_KEY = os.getenv("PICOVOICE_ACCESS_KEY")

if sys.platform == "win32":
    import win32gui
    import win32con

class Assistant:
    def __init__(self, output_callback=None):
        self.config = self.load_config()
        self.assistant_name = self.config.get("assistant_name", "Nora")
        self.wake_word = self.config.get("wake_word", "porcupine")
        self.output_callback = output_callback

        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', self.config.get("voice_options", {}).get("rate", 150))
        self.engine.setProperty('volume', self.config.get("voice_options", {}).get("volume", 0.9))

        self.recognizer = sr.Recognizer()

    def load_config(self):
        try:
            with open("config.json", "r") as f:
                return json.load(f)
        except FileNotFoundError:
            return {}

    def speak(self, text):
        if self.output_callback:
            self.output_callback(text)
        else:
            print(f"{self.assistant_name}: {text}")
            self.engine.say(text)
            self.engine.runAndWait()

    def listen_for_wake_word(self):
        if not PICOVOICE_ACCESS_KEY:
            self.speak("Picovoice access key not found. Please set the PICOVOICE_ACCESS_KEY environment variable.")
            return

        porcupine = None
        pa = None
        audio_stream = None
        try:
            porcupine = pvporcupine.create(access_key=PICOVOICE_ACCESS_KEY, keywords=[self.wake_word])
            pa = pyaudio.PyAudio()
            audio_stream = pa.open(
                rate=porcupine.sample_rate,
                channels=1,
                format=pyaudio.paInt16,
                input=True,
                frames_per_buffer=porcupine.frame_length
            )
            print(f"Listening for wake word ('{self.wake_word}')...")
            while True:
                pcm = audio_stream.read(porcupine.frame_length)
                pcm = struct.unpack_from("h" * porcupine.frame_length, pcm)
                keyword_index = porcupine.process(pcm)
                if keyword_index >= 0:
                    print("Wake word detected!")
                    command = self.listen_for_command()
                    if command:
                        if not self.process_command(command):
                            break # Exit loop if process_command signals to
        except Exception as e:
            self.speak(f"An error occurred with the wake word listener: {e}")
        finally:
            if audio_stream is not None:
                audio_stream.close()
            if pa is not None:
                pa.terminate()
            if porcupine is not None:
                porcupine.delete()

    def listen_for_command(self):
        with sr.Microphone() as source:
            print("Listening for command...")
            self.recognizer.adjust_for_ambient_noise(source)
            audio = self.recognizer.listen(source)
        try:
            command = self.recognizer.recognize_google(audio)
            print(f"You said: {command}")
            return command.lower()
        except sr.UnknownValueError:
            self.speak("Sorry, I didn't catch that.")
            return None
        except sr.RequestError:
            self.speak("Sorry, my speech service is down.")
            return None

    def process_command(self, command_str):
        command, args = parse_command(command_str)
        if command == "exit":
            self.speak("Goodbye!")
            return False # Signal to exit
        elif command == "open_app":
            self.open_application(args)
        elif command == "close_app":
            self.close_application(args)
        elif command == "open_file":
            self.open_file(args)
        elif command == "play_youtube":
            self.play_on_youtube(args)
        elif command == "type_text":
            self.type_text(args)
        else:
            self.speak(f"Sorry, I don't know the command: {command_str}")
        return True

    def type_text(self, text):
        try:
            self.speak(f"Typing: {text}")
            pyautogui.write(text, interval=0.05)
        except Exception as e:
            self.speak(f"An error occurred while typing: {e}")

    def find_executable(self, app_name):
        platform = sys.platform
        app_name = APP_MAPPINGS.get(platform, {}).get(app_name, app_name)
        if os.path.isfile(app_name) and os.access(app_name, os.X_OK):
            return app_name
        search_paths = os.environ.get("PATH", "").split(os.pathsep) + SEARCH_PATHS.get(platform, [])
        for path in search_paths:
            full_path = os.path.join(path, app_name)
            if os.path.isfile(full_path) and os.access(full_path, os.X_OK):
                return full_path
        return None

    def find_file(self, filename):
        home_dir = os.path.expanduser("~")
        for root, _, files in os.walk(home_dir):
            if filename in files:
                return os.path.join(root, filename)
        return None

    def open_uwp_app(self, app_name):
        try:
            subprocess.run(f'start shell:appsfolder\\{app_name}', shell=True, check=True)
            self.speak(f"Opening {app_name}...")
        except subprocess.CalledProcessError:
            self.speak(f"Could not open the UWP app: {app_name}")

    def get_process_pid(self, process_name):
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'].lower() == process_name.lower():
                return proc.info['pid']
        return None

    def bring_window_to_front(self, pid):
        if sys.platform != "win32":
            self.speak("Bringing window to front is only supported on Windows for now.")
            return
        def enum_windows_callback(hwnd, pids):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
                _, found_pid = win32gui.GetWindowThreadProcessId(hwnd)
                if found_pid in pids:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
        win32gui.EnumWindows(enum_windows_callback, [pid])

    def open_application(self, app_name):
        executable_path = self.find_executable(app_name)
        if sys.platform == "win32" and not executable_path and "." in app_name:
            self.open_uwp_app(app_name)
            return
        if not executable_path:
            self.speak(f"Application '{app_name}' not found. Would you like to search online?")
            response = self.listen_for_command()
            if response and 'yes' in response:
                search_url = f"https://www.google.com/search?q={app_name}"
                webbrowser.open(search_url)
            return
        process_name = os.path.basename(executable_path)
        pid = self.get_process_pid(process_name)
        if pid:
            self.speak(f"Application '{app_name}' is already running. Bringing to front.")
            self.bring_window_to_front(pid)
            return
        try:
            if sys.platform == "win32":
                subprocess.Popen([executable_path])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-a", executable_path])
            else:
                subprocess.Popen([executable_path])
            self.speak(f"Opening {app_name}...")
        except Exception as e:
            self.speak(f"An error occurred: {e}")

    def close_application(self, app_name):
        executable_path = self.find_executable(app_name)
        if not executable_path:
            self.speak(f"Application '{app_name}' not found.")
            return
        process_name = os.path.basename(executable_path)
        process_found = False
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'].lower() == process_name.lower():
                try:
                    proc.terminate()
                    proc.wait(timeout=3)
                    self.speak(f"Successfully closed {app_name}.")
                    process_found = True
                except psutil.NoSuchProcess:
                    self.speak(f"{app_name} was already closed.")
                except psutil.TimeoutExpired:
                    self.speak(f"Could not close {app_name} gracefully. Forcing termination.")
                    proc.kill()
                except (psutil.AccessDenied, psutil.ZombieProcess):
                    self.speak(f"Could not close {app_name} due to system restrictions.")
                finally:
                    process_found = True
        if not process_found:
            self.speak(f"Application '{app_name}' is not running.")

    def open_file(self, filename):
        filepath = self.find_file(filename)
        if not filepath:
            self.speak(f"File '{filename}' not found in your home directory.")
            return
        try:
            if sys.platform == "win32":
                os.startfile(filepath)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", filepath])
            else:
                subprocess.Popen(["xdg-open", filepath])
            self.speak(f"Opening {filename}...")
        except Exception as e:
            self.speak(f"An error occurred: {e}")

    def play_on_youtube(self, query):
        self.speak(f"Playing {query} on YouTube.")
        search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
        webbrowser.open(search_url)
