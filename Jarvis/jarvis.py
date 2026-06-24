import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser as wb
import os
import random
import json
import pyautogui
import pyjokes
from win32com.client import Dispatch

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

CHAT_MODEL = os.getenv("JARVIS_OPENAI_MODEL", "gpt-4.1-mini")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONVERSATION_FILE = os.path.join(BASE_DIR, "conversation_history.json")
conversation_history = []
openai_client = None

speaker = None
engine = None

try:
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Rate = 0
    speaker.Volume = 100
except Exception:
    speaker = None

if speaker is None:
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[1].id)
    engine.setProperty('rate', 150)
    engine.setProperty('volume', 1)


def speak(audio) -> None:
    print(audio)
    if speaker is not None:
        try:
            speaker.Speak(audio)
            return
        except Exception:
            pass

    if engine is not None:
        engine.say(audio)
        engine.runAndWait()


def time() -> None:
    """Tells the current time."""
    current_time = get_time_text()
    speak("The current time is")
    speak(current_time)
    print("The current time is", current_time)


def date() -> None:
    """Tells the current date."""
    current_date = get_date_text()
    speak("The current date is")
    speak(current_date)
    print(f"The current date is {current_date}")


def wishme() -> None:
    """Greets the user based on the time of day."""
    speak("Welcome back, sir!")
    print("Welcome back, sir!")

    hour = datetime.datetime.now().hour
    if 4 <= hour < 12:
        speak("Good morning!")
        print("Good morning!")
    elif 12 <= hour < 16:
        speak("Good afternoon!")
        print("Good afternoon!")
    elif 16 <= hour < 24:
        speak("Good evening!")
        print("Good evening!")
    else:
        speak("Good night, see you tomorrow.")

    assistant_name = load_name()
    speak(f"{assistant_name} at your service. Please tell me how may I assist you.")
    print(f"{assistant_name} at your service. Please tell me how may I assist you.")


def screenshot() -> None:
    """Takes a screenshot and saves it."""
    img = pyautogui.screenshot()
    img_path = os.path.expanduser("~\\Pictures\\screenshot.png")
    img.save(img_path)
    speak(f"Screenshot saved as {img_path}.")
    print(f"Screenshot saved as {img_path}.")

def takecommand() -> str:
    """Takes microphone input from the user and returns it as text."""
    r = sr.Recognizer()
    r.dynamic_energy_threshold = True
    r.energy_threshold = 250
    r.pause_threshold = 0.8
    r.non_speaking_duration = 0.5

    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source, duration=0.8)

        try:
            audio = r.listen(source, timeout=8, phrase_time_limit=8)
        except sr.WaitTimeoutError:
            speak("Timeout occurred. Please try again.")
            return None

    print("Recognizing...")
    for language in ("en-US", "en-IN"):
        try:
            query = r.recognize_google(audio, language=language)
            print(query)
            return query.lower()
        except sr.UnknownValueError:
            continue
        except sr.RequestError:
            speak("Speech recognition service is unavailable.")
            return None

    try:
        speak("Sorry, I did not understand that.")
        return None
    except Exception as e:
        speak(f"An error occurred: {e}")
        print(f"Error: {e}")
        return None

def play_music(song_name=None) -> None:
    """Plays music from the user's Music directory."""
    song_dir = os.path.expanduser("~\\Music")
    songs = os.listdir(song_dir)

    if song_name:
        songs = [song for song in songs if song_name.lower() in song.lower()]

    if songs:
        song = random.choice(songs)
        os.startfile(os.path.join(song_dir, song))
        speak(f"Playing {song}.")
        print(f"Playing {song}.")
    else:
        speak("No song found.")
        print("No song found.")


def get_time_text() -> str:
    """Returns the current time in speech-friendly format."""
    return datetime.datetime.now().strftime("%I:%M:%S %p")


def get_date_text() -> str:
    """Returns the current date in speech-friendly format."""
    now = datetime.datetime.now()
    return f"{now.day} {now.strftime('%B')} {now.year}"

def set_name() -> None:
    """Sets a new name for the assistant."""
    speak("What would you like to name me?")
    name = takecommand()
    if name:
        with open("assistant_name.txt", "w") as file:
            file.write(name)
        speak(f"Alright, I will be called {name} from now on.")
    else:
        speak("Sorry, I couldn't catch that.")

def load_name() -> str:
    """Loads the assistant's name from a file, or uses a default name."""
    try:
        with open("assistant_name.txt", "r") as file:
            return file.read().strip()
    except FileNotFoundError:
        return "Jarvis"  # Default name


def search_wikipedia(query):
    """Searches Wikipedia and returns a summary."""
    try:
        speak("Searching Wikipedia...")
        result = wikipedia.summary(query, sentences=2)
        speak(result)
        print(result)
    except wikipedia.exceptions.DisambiguationError:
        speak("Multiple results found. Please be more specific.")
    except Exception:
        speak("I couldn't find anything on Wikipedia.")


def search_wikipedia_text(query: str) -> str:
    """Searches Wikipedia and returns a summary string."""
    try:
        result = wikipedia.summary(query, sentences=2)
        return result
    except wikipedia.exceptions.DisambiguationError:
        return "Multiple Wikipedia results were found. Please be more specific."
    except Exception:
        return "I couldn't find anything on Wikipedia."


def describe_capabilities() -> None:
    """Tells the user what commands the assistant supports."""
    message = (
        "I can tell the time and date, search Wikipedia, open Google or YouTube, "
        "play music, take a screenshot, tell a joke, change my name, or go offline. "
        "If OpenAI is configured, I can also answer general questions like ChatGPT, "
        "remember recent conversations across restarts, and decide when to use my tools."
    )
    speak(message)


def is_greeting(query: str) -> bool:
    """Returns True when a query is a direct greeting or presence check."""
    words = set(query.split())
    short_greetings = {"hello", "hi", "hey"}
    long_greetings = {
        "are you there",
        "are you still there",
        "can you respond",
        "can you hear me",
    }

    return bool(words & short_greetings) or any(phrase in query for phrase in long_greetings)


def reset_conversation() -> None:
    """Clears the AI conversation memory for the current run."""
    conversation_history.clear()
    save_conversation_history()
    speak("Conversation memory cleared.")


def load_conversation_history() -> list:
    """Loads persistent conversation history from disk."""
    try:
        with open(CONVERSATION_FILE, "r", encoding="utf-8") as history_file:
            history = json.load(history_file)
    except FileNotFoundError:
        return []
    except (json.JSONDecodeError, OSError):
        return []

    if not isinstance(history, list):
        return []

    valid_messages = []
    for item in history:
        if not isinstance(item, dict):
            continue
        role = item.get("role")
        content = item.get("content")
        if role in {"user", "assistant"} and isinstance(content, str):
            valid_messages.append({"role": role, "content": content})

    return valid_messages[-12:]


def save_conversation_history() -> None:
    """Persists recent conversation history to disk."""
    try:
        with open(CONVERSATION_FILE, "w", encoding="utf-8") as history_file:
            json.dump(conversation_history[-12:], history_file, indent=2)
    except OSError:
        pass


def get_openai_client():
    """Returns a lazily-initialized OpenAI client."""
    global openai_client

    if OpenAI is None:
        return None

    if not OPENAI_API_KEY:
        return None

    if openai_client is None:
        openai_client = OpenAI(api_key=OPENAI_API_KEY)

    return openai_client


def get_tool_definitions() -> list:
    """Returns tool definitions exposed to OpenAI."""
    return [
        {
            "type": "function",
            "name": "get_time",
            "description": "Get the current local time.",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        {
            "type": "function",
            "name": "get_date",
            "description": "Get the current local date.",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        {
            "type": "function",
            "name": "open_google",
            "description": "Open Google in the default browser.",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        {
            "type": "function",
            "name": "open_youtube",
            "description": "Open YouTube in the default browser.",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        {
            "type": "function",
            "name": "search_wikipedia",
            "description": "Search Wikipedia and return a short summary.",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "The topic to search on Wikipedia."}
                },
                "required": ["query"],
                "additionalProperties": False,
            },
        },
        {
            "type": "function",
            "name": "play_music",
            "description": "Play a song from the user's Music folder, optionally matching a title.",
            "parameters": {
                "type": "object",
                "properties": {
                    "song_name": {"type": "string", "description": "Optional song title or keyword."}
                },
                "additionalProperties": False,
            },
        },
        {
            "type": "function",
            "name": "take_screenshot",
            "description": "Take a screenshot and save it to the user's Pictures folder.",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        {
            "type": "function",
            "name": "tell_joke",
            "description": "Return a short joke.",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        {
            "type": "function",
            "name": "describe_capabilities",
            "description": "Describe what Jarvis can do.",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        {
            "type": "function",
            "name": "reset_chat_memory",
            "description": "Clear saved conversation memory for Jarvis.",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
    ]


def call_tool(tool_name: str, arguments: dict) -> str:
    """Executes a local tool requested by OpenAI and returns text output."""
    if tool_name == "get_time":
        return f"The current time is {get_time_text()}."

    if tool_name == "get_date":
        return f"Today's date is {get_date_text()}."

    if tool_name == "open_google":
        wb.open("https://google.com")
        return "Opened Google."

    if tool_name == "open_youtube":
        wb.open("https://youtube.com")
        return "Opened YouTube."

    if tool_name == "search_wikipedia":
        query = arguments.get("query", "").strip()
        if not query:
            return "No Wikipedia query was provided."
        return search_wikipedia_text(query)

    if tool_name == "play_music":
        song_name = arguments.get("song_name", "").strip() or None
        play_music(song_name)
        if song_name:
            return f"Tried to play music matching {song_name}."
        return "Tried to play music from the Music folder."

    if tool_name == "take_screenshot":
        screenshot()
        return "Screenshot taken and saved to the Pictures folder."

    if tool_name == "tell_joke":
        return pyjokes.get_joke()

    if tool_name == "describe_capabilities":
        return (
            "Jarvis can tell the time and date, search Wikipedia, open Google or YouTube, "
            "play music, take a screenshot, tell jokes, change its name, and chat with OpenAI."
        )

    if tool_name == "reset_chat_memory":
        conversation_history.clear()
        save_conversation_history()
        return "Conversation memory cleared."

    return f"Unknown tool requested: {tool_name}"


def run_ai_turn(client, response) -> str:
    """Processes OpenAI tool calls until a final text response is produced."""
    max_round_trips = 3

    for _ in range(max_round_trips):
        function_calls = [item for item in response.output if item.type == "function_call"]
        if not function_calls:
            answer = response.output_text.strip()
            if answer:
                return answer
            return "I could not generate a response just now. Please try again."

        tool_outputs = []
        for tool_call in function_calls:
            try:
                arguments = json.loads(tool_call.arguments or "{}")
            except json.JSONDecodeError:
                arguments = {}

            output = call_tool(tool_call.name, arguments)
            tool_outputs.append(
                {
                    "type": "function_call_output",
                    "call_id": tool_call.call_id,
                    "output": output,
                }
            )

        response = client.responses.create(
            model=CHAT_MODEL,
            previous_response_id=response.id,
            input=tool_outputs,
        )

    return "I used my tools, but I could not finish the response cleanly. Please try again."


def ask_ai(query: str) -> str | None:
    """Uses OpenAI as a conversational fallback for unknown commands."""
    client = get_openai_client()
    if client is None:
        if not OPENAI_API_KEY:
            return (
                "I can be more intelligent with ChatGPT, but OPENAI_API_KEY is not set yet. "
                "Set that environment variable, then restart Jarvis."
            )

        return (
            "The OpenAI package is not installed yet. Run the launcher install option, "
            "then restart Jarvis."
        )

    messages = [
        {
            "role": "system",
            "content": (
                "You are Jarvis, a concise desktop voice assistant. "
                "Answer naturally, keep responses short enough to be spoken aloud, "
                "and avoid markdown or bullet lists unless the user explicitly asks for them. "
                "Use tools when a user is asking for an action like opening websites, telling the time, "
                "taking screenshots, searching Wikipedia, or playing music."
            ),
        }
    ]

    messages.extend(conversation_history[-6:])
    messages.append({"role": "user", "content": query})

    try:
        response = client.responses.create(
            model=CHAT_MODEL,
            input=messages,
            tools=get_tool_definitions(),
        )
        answer = run_ai_turn(client, response)
    except Exception as error:
        return f"OpenAI request failed: {error}"

    if not answer:
        return "I could not generate a response just now. Please try again."

    conversation_history.append({"role": "user", "content": query})
    conversation_history.append({"role": "assistant", "content": answer})

    if len(conversation_history) > 12:
        del conversation_history[:-12]

    save_conversation_history()

    return answer


conversation_history = load_conversation_history()


def handle_query(query: str) -> bool:
    """Handles a recognized voice command. Returns False to stop the assistant."""
    if "time" in query:
        time()

    elif "date" in query:
        date()

    elif "wikipedia" in query:
        query = query.replace("wikipedia", "").strip()
        search_wikipedia(query)

    elif "play music" in query:
        song_name = query.replace("play music", "").strip()
        play_music(song_name)

    elif "open youtube" in query:
        wb.open("youtube.com")

    elif "open google" in query:
        wb.open("google.com")

    elif "change your name" in query:
        set_name()

    elif "screenshot" in query:
        screenshot()
        speak("I've taken screenshot, please check it")

    elif "tell me a joke" in query:
        joke = pyjokes.get_joke()
        speak(joke)
        print(joke)

    elif is_greeting(query):
        speak("Yes, I'm here. How can I help you?")

    elif "what can you do" in query or "help" in query:
        describe_capabilities()

    elif "how are you" in query:
        speak("I'm working properly and ready to help you.")

    elif "who are you" in query:
        speak("I am Jarvis, your desktop assistant.")

    elif "reset chat" in query or "clear memory" in query:
        reset_conversation()

    elif "shutdown" in query:
        speak("Shutting down the system, goodbye!")
        os.system("shutdown /s /f /t 1")
        return False

    elif "restart" in query:
        speak("Restarting the system, please wait!")
        os.system("shutdown /r /f /t 1")
        return False

    elif "offline" in query or "exit" in query:
        speak("Going offline. Have a good day!")
        return False

    else:
        message = ask_ai(query)
        speak(message)

    return True


if __name__ == "__main__":
    wishme()

    while True:
        query = takecommand()
        if not query:
            continue

        if not handle_query(query):
            break
