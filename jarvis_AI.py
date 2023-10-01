import speech_recognition as sr
import win32com.client
import webbrowser
import os
import datetime


speaker = win32com.client.Dispatch("SAPI.SpVoice")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        
        audio = r.listen(source)
        try:

            query = r.recognize_google(audio, language="en-in")
            return query
        except Exception as e:
            return "Some error occurred"


while True:
    print("Listening..")
    text = takeCommand()

    sites = [["youtube", "https://www.youtube.com"], ["google", "https://www.google.com/"], ["hotstar","https://www.hotstar.com/in/home?ref=%2Fin"], ["github", "https://github.com/"]]
    for site in sites:
        if f"Open {site[0]}".lower() in text.lower():
            speaker.Speak(f"Opening {site[0]} for you Tanisha!")
            webbrowser.open(site[1])

    if "open music".lower() in text.lower():
        musicPath = "song.mp3"
        os.startfile(musicPath)

    if "what is the time".lower() in text.lower():
        hour = datetime.datetime.now().strftime("%H")
        min = datetime.datetime.now().strftime("%M")
        speaker.Speak(f"Tanisha the time is {hour} {min}")    

    if "exit".lower() in text.lower():
        break

speaker.Speak("Thank You, This is Jarvis signing off!!")