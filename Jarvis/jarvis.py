import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser as wb
import os
import random
import pyautogui
import pyjokes
from win32com.client import Dispatch

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
    current_time = datetime.datetime.now().strftime("%I:%M:%S %p")
    speak("The current time is")
    speak(current_time)
    print("The current time is", current_time)


def date() -> None:
    """Tells the current date."""
    now = datetime.datetime.now()
    speak("The current date is")
    speak(f"{now.day} {now.strftime('%B')} {now.year}")
    print(f"The current date is {now.day}/{now.month}/{now.year}")


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


def describe_capabilities() -> None:
    """Tells the user what commands the assistant supports."""
    message = (
        "I can tell the time and date, search Wikipedia, open Google or YouTube, "
        "play music, take a screenshot, tell a joke, change my name, or go offline."
    )
    speak(message)


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

    elif any(
        phrase in query
        for phrase in (
            "hello",
            "hi",
            "hey",
            "are you there",
            "are you still there",
            "can you respond",
            "can you hear me",
        )
    ):
        speak("Yes, I'm here. How can I help you?")

    elif "what can you do" in query or "help" in query:
        describe_capabilities()

    elif "how are you" in query:
        speak("I'm working properly and ready to help you.")

    elif "who are you" in query:
        speak("I am Jarvis, your desktop assistant.")

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
        message = (
            "I heard you, but I do not know that command yet. "
            "Try saying time, date, open google, open youtube, wikipedia, "
            "tell me a joke, screenshot, play music, or exit."
        )
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
