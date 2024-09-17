import subprocess
from typing import Optional, Callable, Dict
import logging
from pygame import mixer
from google.cloud import texttospeech

import config
from security import safe_command

# Configure logging
logging.basicConfig(filename=config.LOG_FILE, level=config.LOG_LEVEL,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Initialize the TTS client
tts_client = texttospeech.TextToSpeechClient()

def text_to_speech(text: str) -> None:
    """
    Converts text to speech and plays the audio.

    Args:
        text (str): The text to be converted to speech.
    """
    try:
        synthesis_input = texttospeech.SynthesisInput(text=text)
        voice = texttospeech.VoiceSelectionParams(
            language_code=config.LANGUAGE_CODE,
            name=f"{config.LANGUAGE_CODE}-Standard-B"
        )
        audio_config = texttospeech.AudioConfig(
            audio_encoding=texttospeech.AudioEncoding.MP3
        )

        response = tts_client.synthesize_speech(
            input=synthesis_input, voice=voice, audio_config=audio_config
        )

        with open("output.mp3", "wb") as out:
            out.write(response.audio_content)
            logger.info('Audio content written to file "output.mp3"')

        mixer.init()
        mixer.music.load('output.mp3')
        mixer.music.play()
    except Exception as e:
        logger.error(f"Error in text_to_speech: {e}")

def run_application(path: str, args: Optional[list] = None) -> None:
    """
    Runs an application with optional arguments.

    Args:
        path (str): The path to the application executable.
        args (Optional[list]): Optional list of arguments for the application.
    """
    try:
        if args:
            safe_command.run(subprocess.Popen, [path] + args)
        else:
            safe_command.run(subprocess.Popen, [path])
        logger.info(f"Application started: {path}")
    except Exception as e:
        logger.error(f"Error starting application {path}: {e}")

def run_dota() -> None:
    """Runs Dota 2."""
    run_application(config.DOTA_PATH)

def open_website(website_url: str) -> None:
    """
    Opens a website in the default browser.

    Args:
        website_url (str): The URL of the website to open.
    """
    run_application(config.CHROME_PATH, [website_url])

def dota_stats() -> None:
    """Opens Dota 2 stats website."""
    open_website('https://www.dotabuff.com/players/436096055')

def work_hours() -> None:
    """Opens work hours tracking website."""
    open_website('https://anthill.workbook.dk/')

def bot_settings() -> None:
    """Opens bot settings website."""
    open_website('https://dashboard.heroku.com/apps/tg-pisunchik-bot')

def work_mode() -> None:
    """Starts applications for work mode."""
    start_teams()
    start_outlook()
    start_spotify()
    start_webstorm()

def home_mode() -> None:
    """Starts applications for home mode."""
    start_spotify()
    start_telegram()
    start_discord()

def start_factorio() -> None:
    """Starts Factorio game."""
    run_application(config.FACTORIO_PATH)

def start_discord() -> None:
    """Starts Discord application."""
    run_application(config.DISCORD_PATH, ['--processStart', 'Discord.exe'])

def start_spotify() -> None:
    """Starts Spotify application."""
    run_application('Spotify.exe')

def start_steam() -> None:
    """Starts Steam application."""
    run_application(config.STEAM_PATH)

def start_teams() -> None:
    """Starts Microsoft Teams application."""
    run_application('ms-teams.exe')

def start_telegram() -> None:
    """Starts Telegram application."""
    run_application(config.TELEGRAM_PATH)

def start_outlook() -> None:
    """Starts Microsoft Outlook application."""
    run_application('olk.exe')

def start_webstorm() -> None:
    """Starts WebStorm IDE."""
    run_application(config.WEBSTORM_PATH)

def start_pycharm() -> None:
    """Starts PyCharm IDE."""
    run_application(config.PYCHARM_PATH)

def start_epic_games() -> None:
    """Starts Epic Games Launcher."""
    run_application(config.EPIC_GAMES_PATH)

def function_choose(function_name: str, website_url: Optional[str] = None, text: Optional[str] = None) -> None:
    """
    Chooses and executes a function based on the provided function name.

    Args:
        function_name (str): Name of the function to execute.
        website_url (Optional[str]): URL for the open_website function.
        text (Optional[str]): Text for the text_to_speech function.
    """
    function_dict: Dict[str, Callable] = {
        "run_dota": run_dota,
        "open_website": lambda: open_website(website_url) if website_url else None,
        "dota_stats": dota_stats,
        "start_factorio": start_factorio,
        "start_discord": start_discord,
        "start_steam": start_steam,
        "start_epic_games": start_epic_games,
        "start_spotify": start_spotify,
        "start_teams": start_teams,
        "start_telegram": start_telegram,
        "start_outlook": start_outlook,
        "start_webstorm": start_webstorm,
        "start_pycharm": start_pycharm,
        "text_to_speech": lambda: text_to_speech(text) if text else None,
        "work_mode": work_mode,
        "home_mode": home_mode,
        "work_hours": work_hours,
        "bot_settings": bot_settings
    }

    function = function_dict.get(function_name)
    if function:
        function()
        logger.info(f"{function_name} executed")
    else:
        logger.warning(f"Function {function_name} not found")

# Error handling decorator
def error_handler(func):
    """
    A decorator for handling errors in functions.

    Args:
        func (Callable): The function to be decorated.

    Returns:
        Callable: The decorated function.
    """
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logger.error(f"Error in {func.__name__}: {str(e)}")
            return None
    return wrapper

# Apply error_handler to all functions
for name, func in list(globals().items()):
    if callable(func) and not name.startswith("__"):
        globals()[name] = error_handler(func)
