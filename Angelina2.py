import sys
#importing modules
import speech_recognition as sr #install speechrecognition
import pyttsx3 #pip install pyttsx3
import datetime
import wikipedia #pip install wikipedia
import webbrowser
import os
import time
import subprocess
from ecapture import ecapture as ec #pip install ecapture
import wolframalpha  #pip install wolframalpha
import json
import requests

import math
import speedtest #pip install speedtest-cli
import psutil
from random import randint
from pynput.keyboard import Key, Controller
import wmi #pip install wmi

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QTimer, QTime, QDate, Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import  *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from angel_edit import Ui_Dialog
from splashscreen1 import Ui_SplashScreen


class TabOpt:
    def __init__(self):
        self.keyboard = Controller()

    def switchTab(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press(Key.tab)
        self.keyboard.release(Key.tab)
        self.keyboard.release(Key.ctrl)

    def closeTab(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('w')
        self.keyboard.release('w')
        self.keyboard.release(Key.ctrl)
        
    def openClosedTab(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press(Key.shift)
        self.keyboard.press('t')
        self.keyboard.release('t')
        self.keyboard.release(Key.shift)
        self.keyboard.release(Key.ctrl)

    def newTab(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('t')
        self.keyboard.release('t')
        self.keyboard.release(Key.ctrl)

class WindowOpt:
    def __init__(self):
        self.keyboard = Controller()

    def closeWindow(self):
        self.keyboard.press(Key.alt_l)
        self.keyboard.press(Key.f4)
        self.keyboard.release(Key.f4)
        self.keyboard.release(Key.alt_l)

    def minimizeWindow(self):
        for i in range(2):
            self.keyboard.press(Key.cmd)
            self.keyboard.press(Key.down)
            self.keyboard.release(Key.down)
            self.keyboard.release(Key.cmd)
            time.sleep(0.05)

    def maximizeWindow(self):
        self.keyboard.press(Key.cmd)
        self.keyboard.press(Key.up)
        self.keyboard.release(Key.up)
        self.keyboard.release(Key.cmd)
        
    def history(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('h')
        self.keyboard.release('h')
        self.keyboard.release(Key.ctrl)

    def switchWindow(self):
        self.keyboard.press(Key.alt_l)
        self.keyboard.press(Key.tab)
        self.keyboard.release(Key.tab)
        self.keyboard.release(Key.alt_l)

def systemInfo():
    c = wmi.WMI()
    my_system_1 = c.Win32_LogicalDisk()[0]
    my_system_2 = c.Win32_ComputerSystem()[0]
    info = f"Total Disk Space: {round(int(my_system_1.Size)/(1024**3),2)} GB\n" \
           f"Free Disk Space: {round(int(my_system_1.Freespace)/(1024**3),2)} GB\n" \
           f"Manufacturer: {my_system_2.Manufacturer}\n" \
           f"Model: {my_system_2. Model}\n" \
           f"Owner: {my_system_2.PrimaryOwnerName}\n" \
           f"Number of Processors: {psutil.cpu_count()}\n" \
           f"System Type: {my_system_2.SystemType}"
    return info


def convert_size(used):
    pass


def system_stats():
    cpu_stats = str(psutil.cpu_percent())
    battery_percent = psutil.sensors_battery().percent
    memory_in_use = convert_size(psutil.virtual_memory().used)
    total_memory = convert_size(psutil.virtual_memory().total)
    stats = f"Currently {cpu_stats} percent of CPU, {memory_in_use} of RAM out of total {total_memory} is being used and " \
                f"battery level is at {battery_percent}%"
    return stats

def get_ip(_return=False):
    try:
        response = requests.get(f'http://ip-api.com/json/').json()
        if _return:
            return response
        else:
            return f'Your IP address is {response["query"]}'
    except KeyboardInterrupt:
        return None
    except requests.exceptions.RequestException:
        return None

def get_joke():
    try:
        joke = requests.get('https://v2.jokeapi.dev/joke/Any?format=txt').text
        return joke
    except KeyboardInterrupt:
        return None
    except requests.exceptions.RequestException:
        return None

def get_speedtest():
    print("......")
    internet = speedtest.Speedtest()
    speed = f"Your network's Download Speed is {round(internet.download() / 8388608, 2)}MBps\n" \
            f"Your network's Upload Speed is {round(internet.upload() / 8388608, 2)}MBps"
    return speed

def get_map(query):
    webbrowser.open(f'https://www.google.com/maps/search/{query}')

def open_cmd():
    os.system('start cmd')                

print('Loading your AI personal assistant Angelina')

engine=pyttsx3.init()
voices=engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)    
engine.setProperty('rate',170)    

def speak(text):
    engine.say(text)
    engine.runAndWait()

def wishMe():
    hour=datetime.datetime.now().hour
    if hour>=0 and hour<12:
        speak("Hello,Good Morning")
        print("Hello,Good Morning")
    elif hour>=12 and hour<18:
        speak("Hello,Good Afternoon")
        print("Hello,Good Afternoon")
    else:
        speak("Hello,Good Evening")
        print("Hello,Good Evening")
# SELF=None
class MainThread(QThread):
    def __init__(self):
        super(MainThread,self).__init__()

    def run(self):
        self.TaskExecution()

    def takeCommand(self):
        r=sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening to the command...")
            audio=r.listen(source)

            try:
                print("Recognizing...")
                self.statement=r.recognize_google(audio,language='en-in')
                print(f"user said:{self.statement}\n")

            except Exception as e:
                return "None"
            return self.statement


    def TaskExecution(self):
        flag=True
        flag2=True
        wishMe()
        
        speak("In order to address you, let me know your gender. Are you male or female?")
        self.gdr=self.takeCommand().lower()
        print(self.gdr)
        while self.gdr=="none":
            speak("I could not hear you. can you please repeat it.")
            print("I could not hear you. can you please repeat it.")
            self.gdr = self.takeCommand().lower()
            print(self.gdr)

        addr=""
        if "female" in self.gdr:
            addr="madam"
        else:
            addr="sir"

        speak("Ok"+addr+" Thank you.")
        tab_ops=TabOpt()
        win_ops=WindowOpt()
        speak("How can I help you"+addr+"?")
        self.statement = self.takeCommand().lower()
        while True:
            if flag2:
                while self.statement=="none":
                    speak("Pardon me,"+addr+"I could not hear you. can you please repeat it.")
                    print("Pardon me,I could not hear you. can you please repeat it.")
                    self.statement = self.takeCommand().lower()
                flag2=False
            if self.statement=="none":
                self.statement = self.takeCommand().lower()
                continue

            if "Hello angelina" in self.statement or "hello" in self.statement:
                speak("hello"+addr+"I am there")
                print("Hello I am there")

            elif "good bye" in self.statement or "ok bye" in self.statement or "stop" in self.statement:
                speak("Thank you"+addr+'your personal assistant Angelina is shutting down. Have a nice day. Good bye')
                print('your personal assistant Angelina is shutting down,Good bye')
                break

            elif 'code' in self.statement:
                speak("Where do you want to code"+addr+", Hacker rank, code chef, codeforces, hacker earth or leetcode?")
                newquery = self.takeCommand().lower()
                while newquery=="none":
                    speak("Pardon me, I could not hear you "+addr+". can you please repeat where do you want to code?")
                    print("Pardon me,I could not hear you. can you please repeat where do you want to code?")
                    newquery = self.takeCommand().lower()
                if 'hackerrank' in newquery or 'hacker rank' in newquery:
                    webbrowser.open("hackerrank.com")
                elif 'codechef' in newquery:
                    webbrowser.open("codechef.com")
                elif 'codeforces' in newquery or 'code forces' in newquery:
                    webbrowser.open("codeforces.com")
                elif 'hacker earth' in newquery or 'hackerearth' in newquery:
                    webbrowser.open("hackerearth.com")
                elif 'leetcode' in newquery:
                    webbrowser.open("leetcode.com")
                time.sleep(5)


            elif 'wikipedia' in self.statement:
                speak('Searching Wikipedia...')
                self.statement =self.statement.replace("wikipedia", "")
                results = wikipedia.summary(self.statement, sentences=3)
                speak("According to Wikipedia")
                print(results)
                speak(results)

            elif 'open wikipedia' in self.statement:
                webbrowser.open_new_tab("https://www.wikipedia.com")
                speak("wikipedia is open now")
                time.sleep(5)

            elif 'open youtube' in self.statement:
                webbrowser.open_new_tab("https://www.youtube.com")
                speak("youtube is open now")
                time.sleep(5)

            elif 'open google' in self.statement:
                webbrowser.open_new_tab("https://www.google.com")
                speak("Google chrome is open now")
                time.sleep(5)

            elif 'open gmail' in self.statement:
                webbrowser.open_new_tab("gmail.com")
                speak("Google Mail open now")
                time.sleep(5)

            elif 'send gmail' in self.statement:
                webbrowser.open_new_tab("https://mail.google.com/mail/u/0/#inbox?compose=new")
                speak("Google Mail open now.You can send your mail")
                time.sleep(5)

            elif "command prompt" in self.statement or "cmd" in self.statement:
                print("Opening command prompt")
                speak("Opening command prompt")
                open_cmd()
                time.sleep(5)

            elif "weather" in self.statement:
                api_key="8ef61edcf1c576d65d836254e11ea420"
                base_url="https://api.openweathermap.org/data/2.5/weather?"
                speak("whats the city name")
                city_name=self.takeCommand().lower()
                while city_name=="none":
                    speak("Pardon me"+addr+", please say the city name again")
                    city_name = self.takeCommand().lower()
                complete_url=base_url+"appid="+api_key+"&q="+city_name
                response = requests.get(complete_url)
                x=response.json()
                if x["cod"]!="404":
                    y=x["main"]
                    current_temperature = y["temp"]
                    current_humidiy = y["humidity"]
                    z = x["weather"]
                    weather_description = z[0]["description"]
                    speak(" Temperature is " +
                          str(current_temperature) +" kelvin ")
                    speak("Humidity is " +
                          str(current_humidiy) + "percent.")
                    speak("So the weather condition is " +
                          str(weather_description))
                    print(" Temperature in kelvin unit = " +
                          str(current_temperature) +
                          "\n humidity (in percentage) = " +
                          str(current_humidiy) +
                          "\n description = " +
                          str(weather_description))

                else:
                    speak(" City Not Found. I regret the inconvenience "+addr)



            elif 'time' in self.statement:
                strTime=datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"the time is {strTime}")

            elif 'who are you' in self.statement or 'what can you do' in self.statement:
                speak('I am Angelina your persoanl desktop voice assistant. . I am programmed to minor tasks like'
                      'opening youtube,google chrome,gmail and stackoverflow ,predict time,take a photo,search wikipedia,predict weather' 
                      'in different cities , get top headline news from times of india and you can ask me computational or geographical questions too!')


            elif "who made you" in self.statement or "who created you" in self.statement or "who discovered you" in self.statement:
                speak("I was built by a great team. My creators are Kudari Nikhilesh, Shaik Naseer Ahmed, Shaik Siddiq Parveen and Repudi Srija. I was made under the guidance of Mrs. N. Padmaja Lavanya Kumari.")
                print("I was built by a great team. My creators are Kudari Nikhilesh, Shaik Naseer Ahmed, Shaik Siddiq Parveen and Repudi Srija. I was made under the guidance of Mrs. N. Padmaja Lavanya Kumari.")

            elif "open stackoverflow" in self.statement:
                webbrowser.open_new_tab("https://stackoverflow.com/login")
                speak("Here is stackoverflow")

            elif 'news' in self.statement:
                news = webbrowser.open_new_tab("https://timesofindia.indiatimes.com/home/headlines")
                speak('Here are some headlines from the Times of India,Happy reading')
                time.sleep(6)

            elif "camera" in self.statement or "take a photo" in self.statement:
                ec.capture(0,"robo camera","img.jpg")

            elif 'search'  in self.statement:
                self.statement = self.statement.replace("search", "")
                webbrowser.open_new_tab(statement)
                time.sleep(5)

            elif "map" in self.statement:
                self.statement = self.statement.replace("map", "")
                speak("Getting the map")
                get_map(self.statement)
                time.sleep(6)

            elif "ip" in self.statement:
                ip = get_ip()
                if ip:
                    print(ip)
                    speak(ip)
                else:
                    speak("IP not found, sorry for inconvenience"+addr)

            elif 'date' in self.statement:
                strdate = datetime.datetime.now()
                day=strdate.strftime("%d")
                month=strdate.strftime("%B")
                year=strdate.strftime("%Y")
                weakday=strdate.strftime("%A")
                speak(f"Today date is {day}")
                speak(f"{month}")
                speak(f"{year}")
                speak(f"Its a {weakday}")

            elif "joke" in self.statement:
                joke = get_joke()
                if joke:
                    speak(joke)
                    print(joke)
                else:
                    speak("Sorry jokes are not available at the moment")

            elif "speed test" in self.statement or "speedtest" in self.statement:
                speak("Getting your internet speed, this may take some time")
                speed = get_speedtest()
                if speed:
                    speak(speed)
                    print(speed)

            elif "system statistics" in self.statement:
                stats = system_stats()
                speak(stats)
                print(stats)

            elif "system" in self.statement and ("info" in self.statement or "specs" in self.statement or "information" in self.statement):
                info = systemInfo()
                speak(info)
                print(info)

            elif "switch tab" in self.statement:
                tab_ops.switchTab()
                speak("switched the tab")

            elif "close tab" in self.statement:
                tab_ops.closeTab()
                speak("Tab closed")

            elif "new tab" in self.statement:
                tab_ops.newTab()
                speak("New tab is opened")

            elif "open closed tab" in self.statement or "restore tab" in self.statement:
                tab_ops.openClosedTab()
                speak("Restored the closed tab")

            elif "close window" in self.statement:
                win_ops.closeWindow()
                speak("Window Closed")

            elif "switch window" in self.statement:
                win_ops.switchWindow()
                speak("window switched")

            elif "minimize window" in self.statement:
                win_ops.minimizeWindow()
                speak("Window minimized")

            elif "maximize window" in self.statement:
                win_ops.maximizeWindow()
                speak("window maximized")

            elif "open history" in self.statement:
                win_ops.history()
                speak("I have opened the history for you")
                time.sleep(5)

            elif "log off" in self.statement or "sign out" in self.statement:
                speak("Ok , your pc will log off in 10 sec make sure you exit from all applications")
                time.sleep(5)
                subprocess.call(["shutdown", "/l"])

            else:
                speak("I did not get you, sorry for the inconvienience"+addr)
                self.statement=self.takeCommand().lower()
                continue


            speak("Do you have any other queries")
            self.statement=self.takeCommand().lower()
            if flag:
                if self.statement=="none":
                    speak("Pardon me, do you have any queries")
                    self.statement = self.takeCommand().lower()
                    flag=False
            if "no" == self.statement:
                speak("Thank you")
                break


    time.sleep(3)

startExecution = MainThread()
# s=None
class Main(QMainWindow):
    def __init__(self):
        global s 
        s=self
        super().__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.startTask()
        self.ui.commandLinkButton.clicked.connect(self.close)

    
    def startTask(self):
        self.ui.movie = QtGui.QMovie(":/newPrefix/voiceanimatedgif.gif")
        self.ui.voiceanimated.setMovie(self.ui.movie)
        self.ui.movie.start()
        self.ui.movie = QtGui.QMovie(":/newPrefix/ezgif.com-crop.gif")
        self.ui.label.setMovie(self.ui.movie)
        self.ui.movie.start()
        self.ui.movie = QtGui.QMovie(":/newPrefix/ezgif.com-crop.gif")
        self.ui.label_2.setMovie(self.ui.movie)
        self.ui.movie.start()
        timer = QTimer(self)
        timer.timeout.connect(self.showTime)
        timer.start(1000)
        startExecution.start()
    def showTime(self):
        current_time = QTime.currentTime()
        current_date = QDate.currentDate()
        label_time = current_time.toString('hh:mm:ss')
        label_date = current_date.toString(Qt.ISODate)
       
#********splash screen*****
counter=0
class Splash(QMainWindow):
	def __init__(self):
		super().__init__()
		self.ui = Ui_SplashScreen()
		self.ui.setupUi(self) 
		self.ui.progressBar.setValue(0)
                               
		self.ui.start.clicked.connect(self.start_clicked)               
		
		
		# remove
		self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
		self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
		#shadow
		self.shadow=QGraphicsDropShadowEffect(self)
		self.shadow.setBlurRadius(20)
		self.shadow.setXOffset(0)
		self.shadow.setYOffset(0)
		self.shadow.setColor(QColor(0,0,0,60))
		self.ui.Bg.setGraphicsEffect(self.shadow)
    
	def start_clicked(self):
                
		speak("Loading your AI personal assistant ANGELINA")
		self.timer=QtCore.QTimer()
		self.timer.timeout.connect(self.progress)
		self.timer.start(35)
	def progress(self):
		
		global counter
		self.ui.progressBar.setValue(counter)
		if counter>100:
			self.timer.stop()
			self.main = Main()
			# print(self.main)
    
			self.main.show()
			self.close()
		counter+=1
#--------------------
app=QApplication(sys.argv)
Angelina2=Splash()
Angelina2.show()
sys.exit(app.exec_())




# if __name__=='__main__':
#     TaskExecution()

