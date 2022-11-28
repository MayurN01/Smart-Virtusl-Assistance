import pyttsx3
import speech_recognition as sr
import os

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
engine.setProperty('rate',170)

def Speak(Audio):
    print("   ")
    print(f": {Audio}")
    engine.say(Audio)
    print("    ")
    engine.runAndWait()

def takecommand(): 
    command = sr.Recognizer()
    with sr.Microphone(device_index=0) as source:
        print("Listening.......")
        command.pause_threshold = 0.8
        audio = command.listen(source, None, 5)

    try:
        print("Recognizing...")    
        query = command.recognize_google(audio, language='en-in')
        print(f"Your Command :  {query}\n")

    except:   
        return "None"
        
    return query.lower()

def TaskExe():
    
    def OpenApps():
        
        if 'wake up robert' in query:
            
            os.startfile("D:\\PBL Project\\Assistance Source Code\\Robert.py")

        elif 'hey Robert' in query:

            os.startfile("D:\\PBL Project\\Assistance Source Code\\Robert.py")
            
    while True:
    
        query = takecommand()

        if 'hey robert' in query:
            OpenApps()

        elif 'wake up robert' in query:
            OpenApps()

        elif 'hey Robert' in query:
            OpenApps()
TaskExe()