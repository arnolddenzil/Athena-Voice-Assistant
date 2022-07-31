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
import shutil
import time

WM_APPCOMMAND = 0x319
APPCOMMAND_MICROPHONE_VOLUME_MUTE = 0x180000

hwnd_active = win32gui.GetForegroundWindow()

mic_muted = False

validate = ['yes', 'y', 'ok', 'okay']
source_path = ""
source_item_name = ""
destination_path = ""

def mic_toggler():
    def on_press(key):
        global mic_muted
        if key == keyboard.Key.esc:
            return False  # stop listener
        try:
            k = key.char  # single-char keys
        except:
            k = key.name  # other keys

        # if k in ['1', '2', 'left', 'right', 'space']:  # keys of interest
        if k in ['1']:
            # self.keys.append(k)  # store it in global-like variable
            print('Key pressed: ' + k)
            win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)
            time.sleep(1)
            win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)
            # if not mic_muted:
            #     print("Mic Muted")
            #     mic_muted = True
            # else:
            #     print("Mic Unmuted")
            #     mic_muted = False
            # time.sleep(2)
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
    print(text)
    engine.say(text)
    engine.runAndWait()


talk("Athena is initializing")

# Creates an nlp_client instance to get curated data from text data

nlp_client = Wit("4XBJUBXVNZKW7ADEVHSR3EPCJSF6Z52C")


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
    command = "notaspeach"
    while command == "notaspeach":
        try:
            # Listens for speach and when it recognizes, it creates an audio file of the speach
            with sr.Microphone() as source:
                listener.adjust_for_ambient_noise(source, duration=0.5)
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


def create_file():
    global validate
    talk("What should be the file name? : ")
    file_name = take_command()
    try:
        f = open(file_name, 'x')
        f.close()
        talk(f"File '{file_name}' created.")
    except FileExistsError:
        talk("The file already exist, do you want to overwrite it ? : ")
        ans = take_command()
        if ans in validate:
            f = open(file_name, 'w')
            f.close()
            talk("File Created")
        else:
            talk("Operation terminated")
            return
    talk(f"Would you like to add some text to {file_name}? : ")
    ans = take_command()
    if ans in validate:
        talk("What would you like to add ? : ")
        text = take_command()
        with open(file_name, 'a') as f:
            f.write(text)
            f.write('\n')
        talk("Do you want to see the entered text ? : ")
        ans = take_command()
        if ans in validate:
            with open(file_name, 'r') as f:
                print("File content : ", f.read())
            talk("Operation complete")
        else:
            talk("Operation terminated")
            return
    else:
        talk("Operation terminated")
        return


def delete_item(data):
    global validate
    item_name = data["entities"]["item:item"][0]["value"]
    if os.path.exists(item_name):
        if os.path.isfile(item_name):
            talk(f"Are you sure you want to delete the file '{item_name}'? : ")
            ans = take_command()
        else:
            talk(f"Are you sure you want to delete the folder '{item_name}'? : ")
            ans = take_command()
        if ans in validate:
            try:
                os.remove(item_name)
            except PermissionError:
                shutil.rmtree(item_name)
                talk(f"Folder '{item_name}' deleted ")
            else:
                talk(f"File '{item_name}' deleted")
    else:
        talk("The item to delete was not found. Operation terminated")


def read_file(data):
    global validate
    file_name = data["entities"]["item:item"][0]["value"]
    if os.path.exists(file_name):
        if os.path.isfile(file_name):
            talk(f"The contents of {file_name} is as follows")
            time.sleep(0.5)
            with open(file_name, "r") as f:
                talk(f.read())
        else:
            talk(f"{file_name} is a directory. Try to list the items")
    else:
        talk(f"The file {file_name} is not found")


def rename_item(data):
    item_name = data["entities"]["item:item"][0]["value"]
    if os.path.exists(item_name):
        talk(f"What should be the new name of {item_name}")
        new_name = take_command()
        talk("Okay")
        os.rename(item_name, new_name)
        talk(f"'{item_name}' has been renamed to '{new_name}'")
    else:
        talk(f"{item_name} is not found. Operation terminated")


def add_text_to_file(data):
    file_name = data["entities"]["item:item"][0]["value"]
    if os.path.exists(file_name):
        if os.path.isfile(file_name):
            talk("What would you like to add ? : ")
            text = take_command()
            with open(file_name, 'a') as f:
                f.write(text)
                f.write('\n')
            talk("The contents are successfully added.")
        else:
            talk(f"{file_name} is a directory. Therefore i cannot add any text to it. Operation terminated")
    else:
        talk(f"The file {file_name} is not found. Operation terminated.")


def list_file_names():
    no_of_files = 0
    for path in os.listdir():
        # check if current path is a file
        if os.path.isfile(path):
            no_of_files += 1
    if no_of_files > 0:
        talk("The files are ")
        for path in os.listdir():
            # check if current path is a file
            if os.path.isfile(path):
                print(path)
                talk(path)
    else:
        talk("No files are present here")
            

def tell_no_of_files():
    no_of_files = 0
    for path in os.listdir():
        # check if current path is a file
        if os.path.isfile(path):
            no_of_files += 1
    if no_of_files == 0:
        talk("No files are present here")
    else:
        talk(f"{no_of_files} number of files are present in the current working directory")


def create_dir():
    global validate
    talk("What should be the folder name? : ")
    folder_name = take_command()
    try:
        os.mkdir(folder_name)
    except FileExistsError:
        talk("The folder already exist, do you want to overwrite it ? : ")
        ans = take_command()
        if ans in validate:
            shutil.rmtree(folder_name)
            os.mkdir(folder_name)


def list_dir_names():
    no_of_folders = 0
    for path in os.listdir():
        # check if current path is a file
        if not os.path.isfile(path):
            no_of_folders += 1
    if no_of_folders > 0:
        talk("The folders are :-")
        for path in os.listdir():
            # check if current path is a file
            if not os.path.isfile(path):
                print(path)
                talk(path)
    else:
        talk("No folders are present here")


def tell_no_of_dir():
    no_of_folders = 0
    for path in os.listdir():
        # check if current path is a file
        if not os.path.isfile(path):
            no_of_folders += 1
    if no_of_folders == 0:
        talk("No folders are present here")
    else:
        talk(f"{no_of_folders} number of folders are present in the current working directory")


def list_drive_names():
    drives = win32api.GetLogicalDriveStrings()
    drives = drives.split('\000')[:-1]
    output = ''
    if len(drives) > 1:
        for i in range(len(drives)-1):
            output += f"{drives[i][0]} "
        output += f"and {drives[i+1][0]}"
        talk("The Drives present are :- " + output)
    else:
        output = drives[0][0]
        talk(f"The only drive present in the system is {output} drive")


def tell_no_of_drives():
    drives = win32api.GetLogicalDriveStrings()
    drives = drives.split('\000')[:-1]
    no_of_drives = len(drives)
    talk(f"There is {no_of_drives} drives present in this computer")


def tell_no_of_items():
    no_of_items = len(os.listdir())
    if no_of_items == 1:
        talk(f"There is 1 item present in the current working directory")
    elif no_of_items == 0:
        talk("There are no items present in the current working directory")
    else:
        talk(f"There are {no_of_items} items present in the current working directory")


def set_working_location(data):
    name_major = data["entities"]["item:item"][0]["value"]
    print(f"Current working directory before hopping inside {name_major}")
    print(os.getcwd())
    locations = {
        'c drive': r'C:',
        'd drive': r'D:',
        'e drive': r'E:',
        'f drive': r'F:',
        'downloads': r'C:\Users\Arnold\Downloads',
        'documents': r'C:\Users\Arnold\Documents',
        'pictures': r'C:\Users\Arnold\Pictures',
        'music': r'C:\Users\Arnold\Music',
        'movies': r'E:\Movies',
        'games': r'D:\TodeleteGames'
    }
    try:
        path_major = locations[name_major]
    except KeyError:
        print("Unfortunately i cannot set the working location there automatically")
    else:
        try:
            os.chdir(path_major)
        except NotADirectoryError:
            print(f"'{name_major}' is not a directory.")
        else:
            print(f"Current working directory after hopping inside {name_major}")
            print(os.getcwd())


def set_source_path(data):
    global source_path
    global source_item_name
    file_name = data["entities"]["item:item"][0]["value"]
    working_dir = os.getcwd()
    tmp_path = os.path.join(working_dir, file_name)
    if not os.path.exists(tmp_path):
        talk("The path does not exist")
    else:
        source_path = tmp_path
        source_item_name = file_name
        talk(f"{file_name} is set as source path")


def set_destination_path(data):
    global destination_path
    path = data["entities"]["item:item"][0]["value"]
    working_dir = os.getcwd()
    temp_path = os.path.join(working_dir, path)
    if not os.path.exists(temp_path):
        talk("The path does not exist")
    else:
        destination_path = temp_path
        talk("Destination path is set")


def copy_operation():
    global source_item_name
    global source_path
    global destination_path
    if not os.path.exists(source_path) and not os.path.exists(destination_path):
        talk("The source and destination path does not exist or is not set")
    elif not os.path.exists(source_path):
        talk("The source path does not exist")
    elif not os.path.exists(destination_path):
        talk("The destination path does not exist")
    else:
        if os.path.isfile(source_path):
            if os.path.isfile(destination_path):
                len_of_dest_path = len(destination_path)
                len_of_source_item_name = len(source_item_name)
                if destination_path[len_of_dest_path - len_of_source_item_name:] == source_item_name:
                    talk("The file already exist. Do you wish to overwrite it ? : ")
                    ans = take_command()
                    if ans in validate:
                        shutil.copy2(source_path, destination_path)
                        talk("Copy operation completed")
                    else:
                        talk("Copy Operation terminated")
                else:
                    if os.stat(destination_path).st_size == 0:
                        shutil.copy2(source_path, destination_path)
                        talk("Copy operation completed")
                    else:
                        talk(f"The file is not empty. Do you wish to replace the contents of the file ? : ")
                        ans = take_command()
                        if ans in validate:
                            shutil.copy2(source_path, destination_path)
                            talk("Copy operation completed")
                        else:
                            talk("Copy Operation terminated")
            else:
                item_exist_in_folder = False
                print(os.listdir(destination_path))
                for item in os.listdir(destination_path):
                    if item == source_item_name:
                        item_exist_in_folder = True
                if item_exist_in_folder:
                    talk("The file already exist. Do you wish to overwrite it ? : ")
                    ans = take_command()
                    if ans in ['yes', 'y', 'ok', 'okay']:
                        shutil.copy2(source_path, destination_path)
                        talk("Copy operation completed")
                    else:
                        talk("Copy Operation terminated")
                else:
                    shutil.copy2(source_path, destination_path)
                    talk("Copy operation completed")

        else:
            if os.path.isfile(destination_path):
                talk("Cannot copy folder to file")
            else:
                item_exist_in_folder = False
                for item in os.listdir(destination_path):
                    if item == source_item_name:
                        item_exist_in_folder = True
                if item_exist_in_folder:
                    talk("A folder with the same name already exist. Do you wish to overwrite it ? : ")
                    ans = take_command()
                    if ans in ['yes', 'y', 'ok', 'okay']:
                        destination_path = rf'{destination_path}\{source_item_name}'
                        shutil.rmtree(destination_path)
                        shutil.copytree(source_path, destination_path)
                        talk("Copy operation completed")
                    else:
                        talk("Copy Operation terminated")
                else:
                    destination_path = rf'{destination_path}\{source_item_name}'
                    shutil.copytree(source_path, destination_path)
                    talk("Copy operation completed")


def move_operation():
    global source_item_name
    global source_path
    global destination_path
    if not os.path.exists(source_path) and not os.path.exists(destination_path):
        talk("The source and destination path does not exist or is not set")
    elif not os.path.exists(source_path):
        talk("Source path does not exist")
    elif not os.path.exists(destination_path):
        talk("Destination path does not exist")
    else:
        if os.path.isfile(source_path):
            if os.path.isfile(destination_path):
                talk("Cannot move a file into a file. Operation terminated")
            else:
                item_exist_in_folder = False
                for item in os.listdir(destination_path):
                    if item == source_item_name:
                        item_exist_in_folder = True
                if item_exist_in_folder:
                    talk("An item with the same name already exist. Do you wish to overwrite it ? : ")
                    ans = take_command()
                    if ans in validate:
                        destination_path = rf'{destination_path}\{source_item_name}'
                        try:
                            os.remove(destination_path)
                        except PermissionError:
                            shutil.rmtree(destination_path)
                        finally:
                            shutil.move(source_path, destination_path)
                            talk("Move operation completed")
                    else:
                        talk("Move Operation terminated")
                else:
                    shutil.move(source_path, destination_path)
                    talk("Move operation completed")
        else:
            if os.path.isfile(destination_path):
                talk("Cannot move folder to file")
            else:
                item_exist_in_folder = False
                for item in os.listdir(destination_path):
                    if item == source_item_name:
                        item_exist_in_folder = True
                if item_exist_in_folder:
                    talk("A folder with the same name already exist. Do you wish to overwrite it ? : ")
                    ans = take_command()
                    if ans in validate:
                        destination_path = rf'{destination_path}\{source_item_name}'
                        shutil.rmtree(destination_path)
                        shutil.move(source_path, destination_path)
                        talk("Move operation completed")
                    else:
                        talk("Move Operation terminated")
                else:
                    destination_path = rf'{destination_path}\{source_item_name}'
                    shutil.move(source_path, destination_path)
                    talk("Move operation completed")


def hop_back():
    print("Current working directory before hopping back")
    print(os.getcwd())
    os.chdir('../')
    talk("okay")
    print("Current working directory after hopping back")
    print(os.getcwd())


def hop_forward(data):
    folder = data["entities"]["item:item"][0]["value"]
    print(f"Current working directory before hopping inside {folder}")
    print(os.getcwd())
    try:
        os.chdir(folder)
    except NotADirectoryError:
        talk(f"'{folder}' is not a directory.")
    else:
        talk(f"The working directory is now set to {folder}")
        print(f"The Current working directory after hopping inside {folder}")
        print(os.getcwd())


def introduction():
    talk("Hi my name is Athena. I am bot developed to do basic file operation by using bot command.")


def end_program():
    talk("goodbye")
    exit(0)


mapping = {
    "greeting": hello,
    "wit$get_time": time_tell,
    "play_youtube": play,
    "create_file": create_file,
    "delete_item": delete_item,
    "end_program": end_program,
    "read_file": read_file,
    "rename_item": rename_item,
    "add_text_to_file": add_text_to_file,
    "list_file_names": list_file_names,
    "tell_no_of_files": tell_no_of_files,
    "create_dir": create_dir,
    "tell_no_of_dir": tell_no_of_dir,
    "list_drive_names": list_drive_names,
    "tell_no_of_drives": tell_no_of_drives,
    "tell_no_of_items": tell_no_of_items,
    "set_working_location": set_working_location,
    "set_source_path": set_source_path,
    "set_destination_path": set_destination_path,
    "copy_operation": copy_operation,
    "move_operation": move_operation,
    "hop_back": hop_back,
    "hop_forward": hop_forward,
    "intoduction": introduction
}


def run_bot():
    command = take_command()
    if command == "notaspeach":
        pass
    else:
        try:
            json_data = get_wit_response(command)
            intent = json_data["intents"][0]["name"]
            print(f"The intend of command is : {intent}")
            print(f"The intented funtion to run is : {mapping[intent]}")
            try:
                mapping[intent]()
            except:
                try:
                    # data = json_data["entities"]["item:item"][0]["value"]
                    mapping[intent](json_data)
                except:
                    print("the intend funtion does not exist")
        except Exception as e:
            print("Handled intend exception")
            print(e)


while True:
    run_bot()
