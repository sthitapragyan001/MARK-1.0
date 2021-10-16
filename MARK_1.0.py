import subprocess
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import wolframalpha
import pyttsx3
import json
import random
from playsound import playsound
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import pyjokes
import winshell
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
import sounddevice
import monotonic
import math
from gtts import gTTS
from clint.textui import progress
import win32com.client as wincl
from bs4 import BeautifulSoup
from urllib.request import urlopen
import linecache
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
f_loc=os.getcwd()
v = open(f_loc+"/voice.txt",'r')
voice=v.read()
v.close()
engine.setProperty('voice', voices[int(voice)].id)
rate = engine.getProperty('rate')
engine.setProperty('rate', rate-50)            
def goto(linenum):
    global line
    line = linenum
    
def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()
def takeCommand():
    
    r = sr.Recognizer()
    
    with sr.Microphone() as source:
        
        print("Listening...")
        r.adjust_for_ambient_noise(source)
        audio = r.record(source,duration=5)
            
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language ='en-in')
            print(f"\n{query}\n")
                
        except Exception as e:
            print("Unable to Recognize your voice.")
            speak("Please type in your command")
            query=input()
            
    return query
def wakeup():
    r = sr.Recognizer()
    m=2+3
    while m==5:
        with sr.Microphone() as source:

            r.adjust_for_ambient_noise(source)
            audio = r.listen(source)
            
            try:
                query = r.recognize_google(audio, language ='en-in')

                if query=='hi '+assname or query=='hello '+assname or query=='hey '+assname:
                    print(f"\n{query}\n")
                    break
                else:
                    print("Unable to Recognize your voice.")
                    print("Please type hello "+assname+" to start giving commands")
                    query=input()
                    break

            except Exception as e:
                m=5
f = open(f_loc+"/usrname.txt",'r')
u_name=f.read()
f.close()
if u_name!='Sir':
    goto(109)

elif u_name=='Sir':
    speak("What should I call you Sir ?")
    nu_name=takeCommand()
    u_name = u_name.replace(u_name,nu_name)
    f = open(f_loc+"/usrname.txt",'w')
    f.write(u_name)
    f.close()

s = open(f_loc+"/speech.txt",'r')
speechi=s.read()
s.close()

if speechi!='hello':
    goto(126)

elif speechi=='hello':
    speak("How would you like to communicate with me Sir... Typing or Speaking? ")
    spechi=takeCommand()
    speechi = speechi.replace(speechi,spechi)
    s = open(f_loc+"/speech.txt",'w')
    s.write(speechi)
    s.close()
    speak("Our mode of communication has been changed")
    speak("To change it back you can ask me to change our mode of communication")

ag = open(f_loc+"/age.txt",'r')
age=ag.read()
ag.close()

if age!='0':
    goto(145)

elif age=='0':
    speak("What's your age Sir ?")
    nage=takeCommand()
    n_age=int(nage)+3
    n_age=str(n_age)
    speak('Sorry, the minimum age for using Mark 1 is '+n_age+' years')
    speak('I was just kidding')
    age = age.replace(age,nage)
    ag = open(f_loc+"/age.txt",'w')
    ag.write(age)
    ag.close()

e = open(f_loc+"/email_id.txt",'r')
email_id=e.read()
e.close()

if email_id!='email':
    goto(160)

elif email_id=='email':
    speak("Please provide your mail id req. for sending mails ")
    nemail_id=input()
    email_id = email_id.replace(email_id,nemail_id)
    e = open(f_loc+"/email_id.txt",'w')
    e.write(email_id)
    e.close()

p = open(f_loc+"/psword.txt",'r')
psword=p.read()
p.close()

if psword!='password':
    goto(175)

elif psword=='password':
    speak("Please provide the password for the email id ")
    npsword=input()
    psword = psword.replace(psword,npsword)
    p = open(f_loc+"/psword.txt",'w')
    p.write(psword)
    p.close()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>= 0 and hour<12:
        speak("Good Morning "+u_name+" !")

    elif hour>= 12 and hour<18:
        speak("Good Afternoon "+u_name+" !")

    else:
        speak("Good Evening "+u_name+" !")
    if 'typing' in speechi:
        s='type'
    else:
        s='say'
    speak('Just '+s+' hello '+assname+' when you would like me to take commands')

def ntakeCommand():
    
    r = sr.Recognizer()
    
    with sr.Microphone() as source:
        
        print("Listening...")
        r.adjust_for_ambient_noise(source)
        audio = r.record(source)
            
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language ='en-in')
            print(f"\n{query}\n")
                
        except Exception as e:
            print("Unable to Recognize your voice.")
            speak("Please type in your command")
            query=input()
            
    return query

def wakeuptype():
    n=5
    while n==5:
        query=input()
        if query=='hi '+assname or query=='hello '+assname or query=='hey '+assname:
            break
        else:
            n=5
            
def intro():
    speak("I am your Personal Assistant, "+assname)
    speak("I have been created by Sthitapragyan")
    speak("And I am here to help you out today")

g = open(f_loc+"/assname.txt",'r')
assname=g.read()
g.close()
    
def tellDay():
      
    # This function is for telling the
    # day of the week
    day = datetime.datetime.today().weekday() + 1
      
    #this line tells us about the number 
    # that will help us in telling the day
    Day_dict = {1: 'Monday', 2: 'Tuesday', 
                3: 'Wednesday', 4: 'Thursday', 
                5: 'Friday', 6: 'Saturday',
                7: 'Sunday'}
      
    if day in Day_dict.keys():
        day_of_the_week = Day_dict[day]
        speak("The day is " + day_of_the_week)

def tellTime():
      
    # This method will give the time
    time = str(datetime.datetime.now())
      
    # the time will be displayed like 
    # this "2020-06-05 17:50:14.582630"
    #nd then after slicing we can get time
    hour = time[11:13]
    mini = time[14:16]
    h=int(hour)
    if h>12:
        h=h-12
        houri=str(h)
        speak("The time is "+houri+":"+mini+" pm")    
    else:
        speak("The time is "+hour+":"+mini+" am")    

def sendEmail(to, message):
    try:
        smtpObj = smtplib.SMTP('smtp-mail.outlook.com',587)
        smtpObj.ehlo()
        smtpObj.starttls()
        smtpObj.login(email_id,psword)
        smtpObj.sendmail('sthitapragyanmallick@outlook.com', to, message)
        print ("Successfully sent email !")
        smtpObj.close()
    except Exception as e:
        print ("Error: unable to send email ",e)

wishMe()
o=0
if o==0:
    
    
    # This Function will clean any
    # command before execution of this python file
    j=2+3
    while j==5:
        if 'typing' in speechi:
            wakeuptype()
        else:
            wakeup()
        speak("How can I help you ")

        while True:

            if 'typing' in speechi:
                print("listening...")
                query=input()
            else:
                query=takeCommand().lower()

            if "change" in query and "your name" in query:
                speak("What do you want to call me Sir")                
                if 'typing' in speechi:
                    nu_aname=input()
                else:
                    nu_aname=takeCommand()

                assname = assname.replace(assname,nu_aname)
                g = open(f_loc+"/assname.txt",'w')
                g.write(assname)
                g.close()
                
                speak("Thanks for naming me")

            elif "your name" in query:
                speak("My friends call me "+assname)

            elif "time" in query:
                tellTime()

            elif "day" in query:
                tellDay()

            elif 'open youtube' in query:
                speak("Here you go to Youtube\n")
                webbrowser.open("youtube.com")

            elif 'google' in query:
                speak("Here you go to Google\n")
                query = query.replace("google", "")
                query = query.replace("in", "")
                webbrowser.open("https://www.google.com/search?q="+query)

            elif 'my age' in query:
                speak('Sir your age is '+nage)

            elif 'your age' in query or 'your birthdsy' in query:
                speak("Officially I was born on 27th July 2021 ")
                speak("So I am still a baby:)")
            
            elif 'play music' in query or "play song" in query:
                            speak("Here you go with music")
                            # music_dir = "G:\\Song"
                            music_dir = "C:\\Users\\Sthitapragyan\\Music"
                            songs = os.listdir(music_dir)
                            print(songs)   
                            random = os.startfile(os.path.join(music_dir, songs[1]))
            elif 'open zoom' in query:
                speak("opening zoom meetings")
                for r,d,f in os.walk("c:\\"):
                    for files in f:
                        if files == "Zoom.exe":
                            app_loc=os.path.join(r,files)
                            os.startfile(app_loc)
                        break
                    break


            elif 'open' in query and 'chrome' in query:
                speak("opening Google Chrome")
                for r,d,f in os.walk("c:\\"):
                    for files in f:
                        if files == "chrome.exe":
                            app_loc=os.path.join(r,files)
                            os.startfile(app_loc)
                        break
                    break

                
            elif'change' in query and 'communication' in query:
                speak("How would you like to communicate with me Sir... Typing or Speaking? ")
                spechi=takeCommand()
                speechi = speechi.replace(speechi,spechi)
                s = open(f_loc+"/speech.txt",'w')
                s.write(speechi)
                s.close()

            elif'change' in query and 'my name' in query:
                speak("What should I call you Sir ?")
                nu_name=takeCommand()
                u_name = u_name.replace(u_name,nu_name)
                f = open(f_loc+"/usrname.txt",'w')
                f.write(u_name)
                f.close()

            elif 'mail' in query:
                try:
                    speak("What should I say")
                    message = takeCommand()
                    speak("Whom should I send")
                    if 'typing' in speechi:
                        kji=input()
                    else:
                        kji = takeCommand().lower()
                    if 'pratap' in kji or 'papa' in kji:
                        to = 'pratapmallick222@gmail.com'
                    elif 'self' in kji or 'myself' in kji:
                        to = 'sthitapragyanmallick@outlook.com'
                    elif kji=='mama':
                        to = 'jhumpiprince@gmail.com'
                    else :
                        speak("Please type the mail id")
                        to = input()

                    sendEmail(to,message)
                    speak("Email has been sent !")
                except Exception as e:
                    print(e)
                    speak("I am not able to send this email")

            elif 'can you do' in query:
                speak ('I can help you with a lot of things, Like...')
                speak('\n Search for info in wikipedia \n Send a mail \n Tell a joke \n Open any app \n Do calculations \n Tell you the latest news \n open up webpages \n make a google search \n Shutdown your computer \n And many more...') 

            elif 'how are you' in query:
                speak("I am fine, Thank you")
                speak("How are you,Sir")
                m=takeCommand()
                if 'am fine' in m:
                    speak("It's good to know that your fine")
                    
            elif 'joke' in query or 'jokes' in query:
                speak(pyjokes.get_joke())
                playsound(f_loc+"/AudienceLaughing.mp3")

            elif 'exit' in query or 'bye' in query or 'close' in query:
                speak("It was my pleasure to help you")
                exit()

            elif "made you" in query or "created you" in query or 'creator' in query:
                intro()
                
            elif "calculate" in query:
                
                app_id = '46A5W3-44XGQUUKJJ'
                client = wolframalpha.Client('46A5W3-44XGQUUKJJ')
                indx = query.lower().split().index('calculate')
                query = query.split()[indx + 1:]
                res = client.query(' '.join(query))
                answer = next(res.results).text
                speak("The answer is " + answer)

            elif 'name' in query and 'meaning' in query:

                speak('My creator Sthitapragyan says that MARK stands for')
                speak("Multi-user Assistant with Responsive Kernel ")
            
            elif 'search' in query or 'play' in query:
                
                query = query.replace("search", "")
                query = query.replace("play", "")       
                webbrowser.open(query)

            elif "who am i" in query or 'my name' in query:
                speak("If you talk then definitely you are Homo Sapien.")
                speak("To be more specific you are "+u_name+":)")

            elif "you come to this world" in query:
                speak("Thanks to Sthitapragyan. further It's a secret")

            elif 'power point' in query:
                speak("opening Power Point presentation")
                for r,d,f in os.walk("c:\\"):
                    for files in f:
                        if files == "powerpnt.exe":
                            app_loc=os.path.join(r,files)
                            os.startfile(app_loc)
                        break
                    break


            elif "who are you" in query:
                intro()

            elif 'news' in query:
                
                try:
                    jsonObj = urlopen('''https://newsapi.org/v2/everything?q=india&language=en&sortBy=publishedAt&apiKey=9e2a8f7c0ddb4a0b8fc34cfb625b4922''')
                    data = json.load(jsonObj)
                    i = 1
                    k = 1
                                
                    speak('here are some top news from the Times Of India')
                    print("               =============== TIMES OF INDIA ============"+ '\n')
                                 
                    while k<=5:
                        for item in data['articles']:
                            print(item['description'] + '\n')   
                            speak(str(i) + '. ' + item['title'] + '\n')
                            i += 1
                        k+=1
                except Exception as e:
                    print(str(e))
                    
            elif 'shutdown' in query:
                    speak("Enter Number of Seconds to Shutdown System: ")
                    if 'typing' in speechi:
                        sec=int(input())
                    else:
                        sec = int(takeCommand())
                    strOne = "shutdown /s /t "
                    strTwo = str(sec)
                    stri = strOne+strTwo
                    os.system(stri)
                    
            elif 'empty recycle bin' in query:
                winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
                speak("Recycle Bin Recycled")

            elif "listen" in query or "listening" in query:
                speak('Okay, I will stop listening now ')
                speak("When you want me to start taking commands again")
                speak("Just say hello "+assname)
                break
                

            elif "restart" in query:
                subprocess.call(["shutdown", "/r"])
            
            elif "hibernate" in query or "sleep" in query:
                speak("Hibernating")
                subprocess.call("shutdown / h")

            elif "log out" in query or "sign out" in query:
                speak("Make sure all the application are closed before sign-out")
                time.sleep(5)
                subprocess.call(["shutdown", "/l"])

            elif "write a note" in query or "take down notes" in query:
                speak("What should I write, sir")
                if 'typing' in speechi:
                    note=input()
                else:
                    note = ntakeCommand()
                file = open('mark.txt', 'w')
                speak("Sir, Should I include date and time")
                if 'typing' in speechi:
                    snfm=input()
                else:
                    snfm = takeCommand()
                if 'yes' in snfm or 'sure' in snfm:
                    strTime = datetime.datetime.now().strftime("% H:% M:% S")
                    file.write(strTime)
                    file.write(" :- ")
                    file.write(note)
                else:
                    file.write(note)
            
            elif "show note" in query:
                speak("Showing Notes")
                file = open("mark.txt", "r")
                speak(file.read(6))
                        
            # NPPR9-FWDCX-D2C8J-H872K-2YT43
            elif " hello "+assname in query or " hey "+assname in query or " hi "+assname in query:
                
                wishMe()
                speak(assname+" at your service master")

            elif "weather" in query:
                
                # Google Open weather website
                # to get API of Open weather
                base_url = "http://api.openweathermap.org/data/2.5/weather?"
                speak("City name")
                if 'typing' in speechi:
                    city_name = input()
                else:
                    city_name = takeCommand()
                complete_url = base_url + city_name + "&appid=96615d461526e88b8ed77259b29142eba" 
                response = requests.get(complete_url)
                x = response.json()
                
                if x["cod"] != "404":
                    y = x["main"]
                    current_temperature = y['temp']
                    current_pressure = y['pressure']
                    current_humidiy = y['humidity']
                    z = x['weather']
                    weather_description = z[0]['description']
                    speak(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
                else:
                    speak(" City Not Found ")
                
            elif "open wikipedia" in query:
                webbrowser.open("wikipedia.com")

            elif "good morning" in query:
                speak("A warm " +query)
                speak("How are you Master?")
                m=takeCommand()
                if 'am fine' in m:
                    speak("It's good to know that your fine")
                

            # elif "" in query:
                # Command go here
                # For adding more commands

            elif 'open ms word' in query:
                speak("opening ms word")
                for r,d,f in os.walk("c:\\"):
                    for files in f:
                        if files == "WINWORD.exe":
                            app_loc=os.path.join(r,files)
                            os.startfile(app_loc)
                        break
                    break


            elif 'open ms excel' in query:
                speak("opening ms excel")
                for r,d,f in os.walk("c:\\"):
                    for files in f:
                        if files == "EXCEL.exe":
                            app_loc=os.path.join(r,files)
                            os.startfile(app_loc)
                        break
                    break
                        
            elif 'open ms teams' in query:
                speak("opening ms teams")
                for r,d,f in os.walk("c:\\"):
                    for files in f:
                        if files == "Update.exe":
                            app_loc=os.path.join(r,files)
                            os.startfile(app_loc)
                        break
                    break
                        
            elif 'clear chat' in query:
                speak("clearing chat")
                clear()
                
            elif 'website' in query:
                speak('Enter the url of the website you would like to visit')
                wurl=input()
                webbrowser.open(wurl)

            elif 'commands' in query or 'command' in query:
                speak("Here are the commands you can give")
                c = open(f_loc+"/commands.txt",'r')
                commands=c.read()
                print(commands)
                c.close()

            elif 'wikipedia' in query or 'who is' in query or 'what is' in query:
                speak('Searching Wikipedia...')
                query = query.replace("wikipedia", "")
                query = query.replace("who is", "")
                query = query.replace("what is", "")            
                results = wikipedia.summary(query, sentences = 2)
                speak("According to Wikipedia")
                speak(results)

            elif 'voice' in query:
                speak("Which voice would you like me to have...")
                speak('Male voice or Female voice ?')
                k=takeCommand().lower()
                if 'female' in k:
                    nvoice='1'
                else:
                    nvoice='0'
                nvoice= voice.replace(voice,nvoice)
                v = open(f_loc+"/voice.txt",'w')
                v.write(voice)
                v.close()
                engine.setProperty('voice', voices[int(nvoice)].id)
                speak("My voice has been successfully changed")
                


        

            
          



                
