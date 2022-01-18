from email.mime import audio
import speech_recognition as sr
import webbrowser as web

path = "C:/Program Files/Google/Chrome/Application/chrome.exe %s"

'''
pazi samo path mora da se poklapa sa tvoju putanju do chrome
'''
r = sr.Recognizer()

with sr.Microphone() as source:
    r.adjust_for_ambient_noise(source)
    print('Say anyting:')
    audio = r.listen(source)

    try:
        dest = r.recognize_google(audio)
        print("You have said : " + dest)
        web.get(path).open(dest)
        
    except Exception as e:
            print("Error : " + str(e))