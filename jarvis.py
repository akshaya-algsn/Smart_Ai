import threading
import speech_recognition as sr
import pyttsx3
import pywhatkit
import datetime
import wikipedia
import pyjokes
import create
import random
import time
from read import start_reading_tasks 
import webbrowser

# Initialize the voice engine
listener = sr.Recognizer()
engine = pyttsx3.init()
engine.setProperty('rate', 175)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

talk_lock = threading.Lock()

def talk(text):
    engine.say(text)
    engine.runAndWait()

def take_command():
    with sr.Microphone() as source:
        print('Listening....')
        try:
            voice = listener.listen(source, timeout=5, phrase_time_limit=5)
            command = listener.recognize_google(voice, language='en-US').lower()
            if 'alexa' in command:
                command = command.replace('alexa', '').strip()
            print(f"You said: {command}")
        except sr.WaitTimeoutError:
            return None
        except sr.UnknownValueError:
            return None
        except sr.RequestError:
            print("API error")
            return None
    return command

def open_gmail():
    talk("Opening Gmail")
    webbrowser.open("https://mail.google.com/")

def play(query):
    song = query.replace('play', '').strip()
    talk(f'Sure, playing {song} for you')
    pywhatkit.playonyt(song)

def time():
    time_str = datetime.datetime.now().strftime('%I:%M %p').lstrip('0')
    talk(f'The current time is {time_str}')

def wiki(command):
    try:
        info = wikipedia.summary(command, sentences=1)
        print(info)
        talk(info)
    except wikipedia.exceptions.DisambiguationError:
        talk("There are multiple results. Can you be more specific?")
    except wikipedia.exceptions.PageError:
        talk("I couldn't find anything on that topic.")



def run_alexa():
    command = take_command()
    if command is None:
        return True

    if 'play' in command:
        play(command)
    elif "open mail" in command or "gmail" in command:
        open_gmail()
    elif 'time' in command:
        time()
    elif 'are you' in command  or 'there' in command:
        talk("Yeah, I am here only")
    elif any(keyword in command for keyword in ['wikipedia', 'about', 'what', 'who']):
        wiki(command)
    elif 'joke' in command:
        joke = pyjokes.get_joke()
        print(joke)
        talk(joke)
    elif 'exit' in command or 'stop' in command:
        talk("I miss you! Goodbye")
        return False
    elif 'remind' in command: #make changes here 
        create.extract_time_and_task(command)
    elif 'schedule' in command or 'task' in command:
        talk("Sure! Please provide the task details.")
        create.main_program()
        talk("Your task has been scheduled! I will notify you at the right time.")

    return True

# Start reading tasks in a background thread
start_reading_tasks()

talk("Hello! How can I assist you today?")

while True:
    if not run_alexa():
        break
