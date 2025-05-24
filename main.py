from datetime import datetime

import speech_recognition as sr
import os
import win32com.client
import webbrowser
import openai
import subprocess


def say(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Cookie"

if __name__ == '__main__':
    say("Hey, I am cookey")
    while True:
        print("Listening....")
        query = takeCommand()

        sites = [
            ["YouTube", "https://www.youtube.com"],
            ["Google", "https://www.google.com"],
            ["Wikipedia", "https://www.wikipedia.org"],
            ["Stack Overflow", "https://stackoverflow.com"],
            ["GitHub", "https://github.com"],
            ["Reddit", "https://www.reddit.com"],
            ["Twitter", "https://twitter.com"],
            ["Instagram", "https://www.instagram.com"],
            ["LinkedIn", "https://www.linkedin.com"],
            ["Amazon", "https://www.amazon.com"],
            ["Netflix", "https://www.netflix.com"],
            ["Coursera", "https://www.coursera.org"],
            ["Khan Academy", "https://www.khanacademy.org"]
        ]

        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])

        apps = [
            ["Chrome", "chrome"],
            ["Firefox", "firefox"],
            ["Edge", "msedge"],
            ["Spotify", r'C:\Users\ASUS\OneDrive\Desktop\Spotify.lnk'],
            ["Notepad", "notepad"]
        ]

        for app in apps:
            if f"open {app[0]}".lower() in query.lower():
                say(f"Opening {app[0]} Vishu...")
                try:
                    os.system(f"start {app[1]}")
                except Exception as e:
                    say(f"Sorry Vishu, I couldn't open {app[0]}")

        if "play" in query.lower():
            song = query.lower().split("play", 1)[1].strip()
            if song:
                say(f"Playing {song} on YouTube, Vishu...")
                search_query = song.replace(" ", "+")
                url = f"https://www.youtube.com/results?search_query={search_query}"
                webbrowser.open(url)
            else:
                say("Kya play karna hai, Vishu?")

        if "the time" in query:
            strfTime = datetime.now().strftime("%H:%M:%S")
            say(f"Vishu time is {strfTime}")


    #say(query)
