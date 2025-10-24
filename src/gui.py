import customtkinter as ctk
import threading

class AssistantGUI(ctk.CTk):
    def __init__(self, assistant):
        super().__init__()
        self.title(assistant.assistant_name)
        self.geometry("500x600")

        self.assistant = assistant
        self.assistant.output_callback = self.update_conversation

        self.conversation_area = ctk.CTkTextbox(self, wrap='word', state='disabled')
        self.conversation_area.pack(pady=10, padx=10, expand=True, fill='both')

        self.input_frame = ctk.CTkFrame(self)
        self.input_frame.pack(pady=10, padx=10, fill='x')

        self.input_box = ctk.CTkEntry(self.input_frame, placeholder_text="Type your command...")
        self.input_box.pack(side='left', expand=True, fill='x')
        self.input_box.bind("<Return>", self.send_command)

        self.send_button = ctk.CTkButton(self.input_frame, text="Send", command=self.send_command)
        self.send_button.pack(side='right', padx=5)

        self.voice_button = ctk.CTkButton(self.input_frame, text="Activate Voice", command=self.activate_voice)
        self.voice_button.pack(side='right')

    def send_command(self, event=None):
        command = self.input_box.get()
        if command:
            self.update_conversation(f"You: {command}", is_user=True)
            self.input_box.delete(0, 'end')
            threading.Thread(target=self.process_and_handle_exit, args=(command,), daemon=True).start()

    def activate_voice(self):
        self.update_conversation("You: (Listening for voice command...)", is_user=True)
        threading.Thread(target=self.listen_and_process, daemon=True).start()

    def listen_and_process(self):
        command = self.assistant.listen_for_command()
        if command:
            self.process_and_handle_exit(command)

    def process_and_handle_exit(self, command):
        if not self.assistant.process_command(command):
            self.destroy()

    def update_conversation(self, text, is_user=False):
        self.conversation_area.configure(state='normal')
        if is_user:
            self.conversation_area.insert('end', text + '\n\n')
        else:
            self.conversation_area.insert('end', f"{self.assistant.assistant_name}: {text}\n\n")
            threading.Thread(target=self.speak_text, args=(text,), daemon=True).start()
        self.conversation_area.configure(state='disabled')
        self.conversation_area.see('end')

    def speak_text(self, text):
        self.assistant.engine.say(text)
        self.assistant.engine.runAndWait()

    def start_wake_word_listener(self):
        wake_word_thread = threading.Thread(target=self.assistant.listen_for_wake_word, daemon=True)
        wake_word_thread.start()

    def start(self):
        self.update_conversation(f"Hello! I'm {self.assistant.assistant_name}. How can I help you?")
        self.start_wake_word_listener()
        self.mainloop()
