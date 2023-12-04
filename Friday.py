import pyttsx3
import os
import pywhatkit as kit
import subprocess as sp
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import pyautogui
from googlesearch import search
from time import sleep
from pathlib import Path
import subprocess
import random
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from word2number import w2n
import re
import threading
import time
import psutil

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)
 
number =0

def speak(audio):
    engine.say(audio)
    engine.runAndWait()
    print(f"Friday : {audio}")

def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak('Good morning')

    elif hour>=12 and hour<15:
        speak("Good afternoon")

    else:
        speak("Good evening")


def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print('Listening.....')
        r.pause_threshold=1
        r.adjust_for_ambient_noise(source,duration=1)
        audio = r.listen(source)
        os.system('cls' if os.name == 'nt' else 'clear')
    try:
        print('Wait for few minutes')    
        query = r.recognize_google(audio,language='en-in')
        os.system('cls' if os.name == 'nt' else 'clear')
        print('user said : ',query)
    except Exception as e:
        print(e)
        query = "nothing"
    return query

def on_or_off():
    pyautogui.hotkey('win','a')
    sleep(2)
    pyautogui.moveTo(1200,400)
    sleep(1)
    pyautogui.click()
    sleep(1)
    pyautogui.hotkey('win','a')
def open_camera():
    sp.run('start microsoft.windows.camera:', shell=True)
    
    sleep(5)
    pyautogui.moveTo(1318,370)
    sleep(1)
    num_clicks = 10  # You can change this number to the desired amount of clicks

    for _ in range(num_clicks):
        sleep(1)
        pyautogui.click()
    path= Path.home()/ 'Pictures//Camera Roll'
    os.startfile(path)
    pyautogui.moveTo(1337,10)
    pyautogui.click()

def close(ans):

      os.system(F"TASKKILL /F /IM {ans}.exe")

def set_volume(percent):
    percent = max(0.0, min(100.0, percent))  # Ensure it's within the valid range (0-100)
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(percent / 100.0, None)

def play_on_youtube(video):
    kit.playonyt(video)

def website(querys,num_results):
    search_results = search(querys, num_results=num_results, lang='en')
    for result in search_results:
        webbrowser.open(result)

def alarm_function():
    speak("Time's up! Your alarm is ringing.")
    # You can add any additional actions or sounds here.

def find_path(path):

    for process in psutil.process_iter(['pid', 'name', 'exe']):
        if process.info['name'] == path +'.exe':
            return process.info['exe']


if __name__ == "__main__":
    
    # subprocess.call(["python", r"D:\Python project\Face Detection\my face recognition.py"])
    wishme()
    while True:
        query = takecommand().lower()
        if "friday" in query:
            speak("I'm on Sir, What I shall help you?")
            while True:
                query = takecommand().lower()
                if "wikipedia" in query:
                    speak("searching in wikipedia")
                    query = query.replace("wikipedia","")
                    results =wikipedia.summary(query,sentences=2)
                    speak("According to wikipedia")
                    print(results)
                    speak(results)

                elif "turn on" in query:
                    speak("turning on bluetooth")
                    on_or_off()

                elif "turn off" in query:
                    speak("turning off bluetooth")
                    on_or_off()

                elif "volume" in query:
                    speak("How would you like to set the volume (e.g., 80 percent or eighty percent)?")
                    i = takecommand().lower()

        # Use regular expressions to extract numeric values
                    match = re.search(r'\b\d+\b', i)
        
                    if match:
                        numbertow = int(match.group())
                        set_volume(numbertow)
                        speak(f"Volume is set to {numbertow}%")
                    else:
                        speak("Sorry, I couldn't understand the volume setting. Please use a clear numeric value.")

                elif "open youtube" in query:
                                speak("opening youtube")
                                webbrowser.open("https://www.youtube.com/")
                elif 'youtube' in query:
                            speak('What do you want to play on Youtube, sir?')
                            video = takecommand().lower()
                            play_on_youtube(video)
                elif 'open qr code' in query:
                     speak("opening QR-code software")
                     subprocess.call(['python',r"D:\Python project\QR-Code\full QR-code.py"])
                elif 'favorite' in query:
                     listm = ['tamil love song', 'tamil motivation song','Ordinary Person','Naa Ready','Badass tamil']
                     video = random.choice(listm)
                     play_on_youtube(video)
                     speak(f"Okay sir ! I'm Playing {video} video on youtube sir!")
                
                elif "good job" in query:
                     li = ["Thank you Sir", "Thanks Sir","You're welcome!","No worries","Don't mention it!","My pleasure!","Anytime","It was the least I could do","Glad to help","Sure!","Thank you"]
                     words= random.choice(li)
                     speak(words)

                elif "thank you" in query:
                     li = ["Thank you Sir", "Thanks Sir","You're welcome!","No worries","Don't mention it!","My pleasure!","Anytime","It was the least I could do","Glad to help","Sure!","Thank you"]
                     words= random.choice(li)
                     speak(words)
                elif "open camera" in query:
                      speak('opening camera sir')
                      open_camera()
                      
                elif 'website' in query:
                            speak('What do you want to open websites on Google, sir?')
                            querys = takecommand().lower()
                            num_results = 5
                            website(querys,num_results)
                elif "open google" in query:
                                speak("opening google")
                                webbrowser.open("https://www.google.com/")

                elif "about myself" in query:
                     with open('About_user.txt', 'r') as file:
                        for line in file:
                            speak(line)
                            sleep(1)

                elif "open chat" in query:
                     speak("opening ChatGPT")
                     webbrowser.open("https://chat.openai.com/")

                elif "open drive" in query:
                    speak("opening drive")   
                    webbrowser.open("https://drive.google.com/drive/my-drive")
                elif "show me data" in query:
                     speak("OK Sir ! I will show you QR-code Data.")
                     folder_path=Path.home() / r"Pictures\qrcode_folder"
                     os.startfile(folder_path)
                elif 'open chrome' in query:
                    filename= "chrome"
                    chromepath = find_path(filename)
                    os.startfile(chromepath)

                elif 'open viber' in query:
                     filename = "Viber"
                     viberpath = find_path(filename)
                     os.startfile(viberpath)

                elif 'close chrome' in query:
                    os.system(F"TASKKILL /F /IM chrome.exe")

                elif "the time" in query:
                    time = datetime.datetime.now().strftime("%H:%M")
                    speak(time)

                elif "shut down" in query:
                        speak("ok sir ,I am shut down pc . Just wait for a few minutes!")
                        os.system("shutdown /s /t 1")
                        exit()

                elif "show me photos" in query :
                      speak("OK Sir ! I will show you your Pictures.")
                      folder_path=Path.home() / "Pictures"
                      os.startfile(folder_path)

                elif "stop listen" in query:
                    break

                elif "open profile" in query:
                     speak('What do you want to open profile, sir?')
                     speak('For Example Github and Linkedin, sir!')
                     webbrowser.open("https://github.com/MinZawHtay")
                     webbrowser.open("https://www.linkedin.com/in/min-zaw-htay-0a2148220/")
                          
                elif 'goodbye' in query:
                    speak('Okay Sir , goodbye . Have a nice day. ')
                    exit()