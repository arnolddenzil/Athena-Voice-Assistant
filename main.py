import speech_recognition as sr
from wit import Wit
import json
import wikipedia
import win32api
import win32gui
from pynput import keyboard
from threading import Thread
import time
import pyttsx3
import datetime
import pywhatkit
import os

WM_APPCOMMAND = 0x319
APPCOMMAND_MICROPHONE_VOLUME_MUTE = 0x180000

hwnd_active = win32gui.GetForegroundWindow()

mic_muted = False

def mic_toggler():
    def on_press(key):
        global mic_muted
        if key == keyboard.Key.esc:
            return False  # stop listener
        try:
            k = key.char  # single-char keys
        except:
            k = key.name  # other keys

        if k in ['1', '2', 'left', 'right', 'space']:  # keys of interest

            # self.keys.append(k)  # store it in global-like variable
            print('Key pressed: ' + k)
            win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)
            if not mic_muted:
                print("Mic Muted")
                mic_muted = True
            else:
                print("Mic Unmuted")
                mic_muted = False
            time.sleep(2)
            # win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)
            # print("Mic Unmuted")
            # return False  # stop listener; remove this if want more keys

    keyboard_listener = keyboard.Listener(on_press=on_press)
    keyboard_listener.start()  # start to listen on a separate thread
    print("Started listening for keyboard interrupt for mic toggling")
    keyboard_listener.join()  # remove if main thread is polling self.keys


Thread(target=mic_toggler).start()

# Initialize listener instance
listener = sr.Recognizer()

# Initialize text to speach engine
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)   # Female Voice is given
engine.setProperty('rate', 165)     # setting up speaking rate


# Speaks the text passed to it
def talk(text):
    engine.say(text)
    engine.runAndWait()


talk("Athena is initializing")

# Creates an nlp_client instance to get curated data from text data
nlp_client = Wit("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")


def get_wit_response(command):
    wit_response = nlp_client.message(command)
    json_string = str(wit_response).replace("'", '"').replace("True", "true")
    json_data = json.loads(json_string)
    print("####################################")
    beautified_json_string = json.dumps(json_data, indent=4)
    print(beautified_json_string)
    print("####################################")
    return json_data


c = 0

# Returns the voice command as text
def take_command():
    global c
    try:
        # Listens for speach and when it recognizes, it creates an audio file of the speach
        with sr.Microphone() as source:
            listener.adjust_for_ambient_noise(source, duration=1)
            if c == 0:
                talk("Initialization Complete. listening.")
                c += 1
            print("Listening ...")
            audio_file = listener.listen(source=source)
            print("Voice Heard")

            # Gets text from the audio file using google recognizer
            try:
                command = listener.recognize_google(audio_file)
                print("Google Speech Recognition results: " + command)
            except sr.UnknownValueError:
                print("Google Speech Recognition could not understand audio")
                command = "notaspeach"
            except sr.RequestError as e:
                print("Could not request results from Google Speech Recognition service; {0}".format(e))
                command = "notaspeach"

    # Handles exception that could arise from the listener instance
    except Exception as e:
        print("execption occured")
        print(e)
        command = "notaspeach"

    # Returns text command
    return command


def play(data):
    song = data["entities"]["item:item"][0]["value"]
    talk('playing ' + song)
    pywhatkit.playonyt(song)


def hello():
    talk("hello")


def time_tell():
    time = datetime.datetime.now().strftime('%I:%M %p')
    talk('Current time is ' + time)


def tell_joke():
    talk(pyjokes.get_joke())


def create_note():
    talk("What is the content to be written ?")
    note_to_save = take_command()
    talk("What should be name of the file ?")
    name = take_command()
    with open(f"{name}.txt", "w") as f:
        f.write(note_to_save)
    talk("Note Created")
    talk("Do you want me to read the note just created ?")
    yes_or_no = take_command()
    while yes_or_no == "notaspeach":
        yes_or_no = take_command()

    if yes_or_no == "yes" or yes_or_no == "ok":
        with open(f"{name}.txt", "r") as f:
            note = f.read()
            talk(f"the content in {name} is {note}")


def delete_note(data):
    note_to_delete = data["entities"]["item:item"][0]["value"]
    print(f"Data returned to note_to_delete : {note_to_delete}")
    try:
        os.remove(f"{note_to_delete}.txt")
    except FileNotFoundError:
        try:
            os.remove(note_to_delete)
        except FileNotFoundError:
            talk("The request file was not found")
        else:
            talk(f"Deleted {note_to_delete}")
    else:
        talk(f"Deleted {note_to_delete}")




mapping = {
    # "greeting": lambda: hello(),
    "greeting": hello,
    "wit$get_time": time_tell,
    "joke": tell_joke,
    "play_youtube": play,
    "create_note": create_note,
    "delete_note": delete_note

}


def run_bot():
    command = take_command()
    if command == "notaspeach":
        pass
    else:
        json_data = get_wit_response(command)
        intent = json_data["intents"][0]["name"]
        print(intent, type(intent))
        try:
            print(mapping[intent])
            mapping[intent]()
        except:
            try:
                # data = json_data["entities"]["item:item"][0]["value"]
                mapping[intent](json_data)
            except:
                print("the intend funtion does not exist")

while True:
    run_bot()
