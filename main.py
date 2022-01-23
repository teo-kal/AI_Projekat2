from email.mime import audio
from tracemalloc import stop
from gtts import gTTS
from datetime import datetime 
import speech_recognition as sr
import win32com.client as wincl
import os 
import subprocess
from pathlib import Path
#import time

def openProgram(speak, programs):
    speak.Speak("Here's a list of programs. Which one would you like to open?")

    for key in programs.keys():
        print("  > " + key)

    audio = r.listen(source)
    recognizedAudio = str.lower(r.recognize_google(audio))

    if(recognizedAudio in programs.keys()):
        if isinstance(programs[recognizedAudio], tuple):
            speak.Speak("Do you want to open a specific page?")
            print("Do you want to open a specific page?")
            
            audio = r.listen(source)
            ans = str.lower(r.recognize_google(audio))

            if(ans == "yes"):
                speak.Speak("Which one? Example: google.com")
                print("Which one? (Example: google.com)")                            

                audio = r.listen(source)
                ans = str.lower(r.recognize_google(audio))

                subprocess.Popen([programs[recognizedAudio][0], ans])
            else:
                #treba Popen da ne bi prekinuo ovaj proces
                subprocess.Popen([programs[recognizedAudio][0], programs[recognizedAudio][1]])
        else:
            subprocess.Popen([programs[recognizedAudio]])
    else:
        speak.Speak("I have no idea what you just said")


def dateAndTime(speak):
    now = datetime.now()
    day = now.strftime("%d")

    if(int(day) % 10 == 1):
        day += "st"
    elif(int(day) % 10 == 2):
        day += "nd"
    elif(int(day) % 10 == 3):
        day += "rd"

    monthAndYear = now.strftime("%B %Y")

    ampm = "PM" if int(now.strftime("%H")) > 12 else "AM"
    hour = str(int(now.strftime("%H")) % 12)
    mins = now.strftime("%M")
    
    print("The date is: " + day + " of " + monthAndYear)
    print("The current time is: " + hour + ":" + mins + " " + ampm)

    speak.Speak("The date is: " + day + " of " + monthAndYear)
    speak.Speak("The current time is: " + hour + mins + ampm)


def closeProgram(speak):
    speak.Speak("Which program would you like to close?")
    print("Which program would you like to close?")
    print("NOTE: You have to specify the process name (Ex: 'chrome')")

    audio = r.listen(source)
    ans = str.lower(r.recognize_google(audio))

    os.system("taskkill /f /im " + ans + ".exe")


def createMemo(speak):
    print("I'll write a .txt file for you. What would you like to name it?")
    speak.Speak("I'll write a .t x t file for you. What would you like to name it?")

    audio = r.listen(source)
    ans = str.lower(r.recognize_google(audio))
   
    print("You said: " + ans)
    #speak("You said: " + str(ans))     #Puca -_-

    if os.path.exists("./memos/" + ans + ".txt"):
        print("A file with this name already exists. Do you want to repeat your input? (yes/no)")
        speak.Speak("A file with this name already exists. Do you want to repeat your input?")
    else:
        print("Do you want to repeat your input? (yes/no)")
        speak.Speak("Do you want to repeat your input?")

    audio = r.listen(source)
    confirmation = str.lower(r.recognize_google(audio))

    while "no" not in confirmation:
        print("What would you like to name the file?")
        speak.Speak("What would you like to name the file?")

        audio = r.listen(source)
        ans = str.lower(r.recognize_google(audio))

        print("You said: " + ans)

        if os.path.exists("./memos/" + ans + ".txt"):
            print("A file with this name already exists. Do you want to repeat your input? (yes/no)")
            speak.Speak("A file with this name already exists. Do you want to repeat your input?")
        else:
            print("Do you want to repeat your input? (yes/no)")
            speak.Speak("Do you want to repeat your input?")

        audio = r.listen(source)
        confirmation = str.lower(r.recognize_google(audio))

    speak.Speak("BEGIN MEMO:")

    audio = r.listen(source)
    text = str.lower(r.recognize_google(audio))

    try:
        Path("./memos").mkdir(parents=True, exist_ok=True)  #NOTE: stavila sam ./memos/ u .gitignore
        
        with open("./memos/" + ans + ".txt", 'w') as f:
            f.write(text)

    except Exception as e:
        print("Error: " + e)

    print("Done. The file is in the memos folder.")
    speak.Speak("Done. The file is in the memos folder.")

######MAIN###########
r = sr.Recognizer()
speak = wincl.Dispatch("SAPI.SpVoice")

programs = {
    "chrome": ('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe', '-new-tab'),
    "firefox": ('C:\\Program Files\\Mozilla Firefox\\firefox.exe', '-new-tab'),
    "notepad": 'C:\\Windows\\System32\\notepad.exe'
}

commandList = ["Show Commands", "Open Program", "Create Memo", "Google", "Date and Time"]

commands = {
    "Open Program": openProgram,
    "Date and Time": dateAndTime,
    "Close Program": closeProgram,
    "Create Memo": createMemo
}

with sr.Microphone() as source:
    r.adjust_for_ambient_noise(source)

    print("Say something")

    audio = r.listen(source)

    recognizedAudio = str.lower(r.recognize_google(audio))

    while("stop" not in recognizedAudio):

        try:
            print("You said: " + recognizedAudio)

            if("hello" in recognizedAudio):
                speak.Speak("Hello!")

            if("how are you" in recognizedAudio):
                speak.Speak("I'm fine, thank you.")

            if("list" in recognizedAudio or "commands" in recognizedAudio):
                for command in commandList:
                    print("  > " + command)

            if("open program" in recognizedAudio):
                commands["Open Program"](speak, programs)

            if("date" in recognizedAudio or "time" in recognizedAudio):
                commands["Date and Time"](speak)

            if("close program" in recognizedAudio):
                commands["Close Program"](speak)

            if("memo" in recognizedAudio):
                commands["Create Memo"](speak)
                

        except Exception as e:
            print("Error: " + e)

        print("Say something")

        audio = r.listen(source)

        recognizedAudio = r.recognize_google(audio)

