import asyncio
import customtkinter as ctk
import threading
import keyboard
import speech_recognition as sr
import logging
import openai
from config import *
import functions
import json
import openpyxl
import cryptography

cryptography.decrypt_file("encrypted_assistant_creds.xlsx")

# Load the workbook
workbook = openpyxl.load_workbook('decrypted_assistant_creds.xlsx')

# Select the active worksheet (or you can select by name using workbook['SheetName'])
sheet = workbook.active

# Read the value from cell A2
value_a2 = sheet['B1'].value

# Initialize OpenAI API
client = openai.OpenAI(api_key=value_a2)

# Configure logging
logging.basicConfig(filename=LOG_FILE, level=LOG_LEVEL,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Load GPT functions
with open(FUNCTIONS_FILE, "r") as f:
    gpt_functions = json.load(f)

# Initialize speech recognizer and microphone
recognizer = sr.Recognizer()
mic = sr.Microphone()

class VoiceCommandApp:
    """Main application class for the Voice Command App."""

    def __init__(self):
        self.root = ctk.CTk()
        self.label = None
        self.command_label = None

        self.setup_ui()

    def setup_ui(self):
        """Set up the user interface."""
        self.root.title("Voice Command App")
        self.root.geometry("400x200")

        self.label = ctk.CTkLabel(self.root, text="Загрузка...", font=("Helvetica", 18))
        self.label.pack(pady=20)

        self.command_label = ctk.CTkLabel(self.root, text="", font=("Helvetica", 14))
        self.command_label.pack(pady=20)

        self.root.after(2000, self.on_loaded)

    def on_loaded(self):
        """Callback when the app has finished loading."""
        self.label.configure(text="Загружено!\n Нажмите ALT + I чтобы ввести команду")
        keyboard.add_hotkey('alt+i', self.on_activate)

    def on_activate(self):
        """Callback when the activation hotkey is pressed."""
        self.label.configure(text="Слушаю...")
        threading.Thread(target=self.listen_and_process_command).start()

    def listen_and_process_command(self):
        """Handles listening to the voice command and processing it."""
        asyncio.run(self.process_command())

    async def process_command(self):
        """Processes the voice command."""
        try:
            voice_command = await self.listen_to_voice()
            if voice_command:
                self.label.configure(text="Ожидайте...")
                self.command_label.configure(text=f"Команда: {voice_command}")
                await self.run_conversation(voice_command)
            else:
                self.command_label.configure(text="Не удалось распознать команду")
        except Exception as e:
            logger.error(f"Error processing command: {e}")

    async def listen_to_voice(self):
        """Listens for and recognizes voice input."""
        with mic as source:
            logger.info("Listening for voice command...")
            audio = await asyncio.to_thread(recognizer.listen, source)
        try:
            return await asyncio.to_thread(recognizer.recognize_google, audio, language=LANGUAGE_CODE)
        except sr.UnknownValueError:
            logger.warning("Could not understand the audio")
            return ""
        except sr.RequestError as e:
            logger.error(f"Could not request results; {e}")
            return ""

    async def run_conversation(self, voice_command):
        """Processes the voice command using OpenAI's API."""
        try:
            self.label.configure(text="Выполняю...")
            response = await asyncio.to_thread(
                client.chat.completions.create,
                model="gpt-4o",
                messages=[{"role": "user", "content": voice_command}],
                functions=gpt_functions,
                function_call="auto",
            )
            message = response.choices[0].message

            if message.function_call:
                await self.handle_function_call(message.function_call)

            if message.content:
                await asyncio.to_thread(functions.text_to_speech, message.content)
        except Exception as e:
            logger.error(f"Error in run_conversation: {e}")

    async def handle_function_call(self, function_call):
        """Handles the function call returned by the OpenAI API."""
        function_name = function_call.name
        function_params = json.loads(function_call.arguments)

        if function_name == "start_multiple_functions":
            await self.handle_multiple_functions(function_params)
        else:
            await self.handle_single_function(function_name, function_params)

    async def handle_multiple_functions(self, params_dict):
        """Handles multiple function calls."""
        for func in params_dict.get("functions", []):
            await self.handle_single_function(func, params_dict)

    async def handle_single_function(self, function_name, params_dict):
        """Handles a single function call."""
        if function_name == "open_website":
            website_url = params_dict["website"]
            await asyncio.to_thread(functions.function_choose, function_name, website_url=website_url)
        elif function_name == "text":
            text = params_dict["text"]
            await asyncio.to_thread(functions.function_choose, function_name, text=text)
        else:
            await asyncio.to_thread(functions.function_choose, function_name)

    def run(self):
        """Runs the application."""
        self.root.mainloop()


if __name__ == "__main__":
    app = VoiceCommandApp()
    app.run()
