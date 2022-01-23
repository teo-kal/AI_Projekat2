from email.mime import audio
from tracemalloc import stop
from gtts import gTTS
from datetime import datetime 
import speech_recognition as sr
import win32com.client as wincl
#import os 
import subprocess
import time

def openProgram(speak, programs):
    speak.Speak("Here's a list of programs. Which one would you like to open?")

    for key in programs.keys():
        print("> " + key)

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

######MAIN###########
r = sr.Recognizer()
speak = wincl.Dispatch("SAPI.SpVoice")

programs = {
    "chrome": ('C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe', '-new-tab'),
    "firefox": ('C:\\Program Files\\Mozilla Firefox\\firefox.exe', '-new-tab'),
    "notepad": 'C:\\Windows\\System32\\notepad.exe'
}

commands = {
    "commands": ["Show Commands", "Open Program", "Memo", "Reminder", "Google", "Date and Time"],
    "Open Program": openProgram,
    "Date and Time": dateAndTime#,
    #"Memo": memo
}

with sr.Microphone() as source:
    r.adjust_for_ambient_noise(source)

    print("Say something")

    audio = r.listen(source)

    recognizedAudio = r.recognize_google(audio)

    while("stop" not in recognizedAudio):

        try:
            print("You said: " + recognizedAudio)

            if("hello" in recognizedAudio):
                #print("HI BACK!")
                #output = gTTS(text="Hello to you too!", lang = "en", slow = False)
                speak.Speak("Hello World")

            if("list" in recognizedAudio):
                for command in commands["commands"]:
                    print("  > " + command)

            if("open program" in recognizedAudio):
                commands["Open Program"](speak, programs)

            if("date" in recognizedAudio or "time" in recognizedAudio):
                commands["Date and Time"](speak)
                

        except Exception as e:
            print("Error: " + e)

        print("Say something")

        audio = r.listen(source)

        recognizedAudio = r.recognize_google(audio)

