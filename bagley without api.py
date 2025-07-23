import os
import webbrowser
import subprocess
import psutil
import pyttsx3
import speech_recognition as sr
import pyautogui
import pyperclip
from datetime import datetime
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
import platform
# Memory file
MEMORY_FILE = "memory.txt"

def load_memory():
    if os.path.exists(MEMORY_FILE):
        with open(MEMORY_FILE, "r", encoding="utf-8") as f:
            return f.read()
    return ""

def save_memory(text):
    with open(MEMORY_FILE, "a", encoding="utf-8") as f:
        f.write(text + "\n")

chatStr = load_memory()

def say(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

def listen():
    r = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source, phrase_time_limit=4)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"You: {query}")
            return query.lower()
        except Exception:
            return ""

# ---------- FEATURE FUNCTIONS ----------
def change_volume(command):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    if "mute" in command:
        volume.SetMute(1, None)
        say("Volume muted")
    elif "unmute" in command:
        volume.SetMute(0, None)
        say("Volume unmuted")
    elif "increase" in command or "raise" in command:
        current_volume = volume.GetMasterVolumeLevel()
        volume.SetMasterVolumeLevel(current_volume + 2.0, None)
        say("Volume increased a bit")
    elif "decrease" in command or "lower" in command:
        current_volume = volume.GetMasterVolumeLevel()
        volume.SetMasterVolumeLevel(current_volume - 2.0, None)
        say("Volume decreased")

def open_application(command):
    apps = {
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
        "chrome": "chrome.exe"
    }
    for app in apps:
        if app in command:
            subprocess.Popen(apps[app])
            say(f"Opening {app}")
            return

def close_application(command):
    for proc in psutil.process_iter():
        if "notepad" in command and proc.name() == "notepad.exe":
            proc.terminate()
            say("Closed Notepad")

def open_folder(command):
    if "documents" in command:
        os.startfile(os.path.expanduser("~/Documents"))
        say("Opened Documents folder")
    elif "downloads" in command:
        os.startfile(os.path.expanduser("~/Downloads"))
        say("Opened Downloads folder")

def system_control(command):
    if "shutdown" in command:
        say("Shutting down the system")
        os.system("shutdown /s /t 1")
    elif "restart" in command:
        say("System is restarting")
        os.system("shutdown /r /t 1")
    elif "log off" in command:
        say("Logging off")
        os.system("shutdown /l")

def clipboard_interaction(command):
    if "copy" in command:
        pyautogui.hotkey("ctrl", "c")
        say("Copied")
    elif "paste" in command:
        pyautogui.hotkey("ctrl", "v")
        say("Pasted")
    elif "show clipboard" in command:
        clipboard_content = pyperclip.paste()
        say(f"This is on the clipboard: {clipboard_content}")

def manage_windows(command):
    if "minimize" in command:
        pyautogui.hotkey("win", "down")
        say("Window minimized")
    elif "maximize" in command:
        pyautogui.hotkey("win", "up")
        say("Window maximized")
    elif "switch" in command:
        pyautogui.hotkey("alt", "tab")
        say("Switching window")

def google_search(query):
    search_query = query.replace("search", "").strip()
    url = f"https://www.google.com/search?q={search_query}"
    webbrowser.open(url)
    say(f"Searching {search_query} on Google")

def type_text(command):
    text_to_type = command.partition("type")[2].strip()
    if text_to_type:
        pyautogui.write(text_to_type)
        say(f"Typing: {text_to_type}")
    else:
        say("Please specify what to type.")

def scroll_screen(command):
    if "down" in command:
        pyautogui.scroll(-500)
        say("Scrolling down")
    elif "up" in command:
        pyautogui.scroll(500)
        say("Scrolling up")
    else:
        say("Please specify scroll direction, up or down.")

def send_message(command):
    # Example: "send message to John hello how are you"
    if "to" in command:
        parts = command.split("to", 1)
        if len(parts) > 1:
            rest = parts[1].strip()
            if " " in rest:
                name, message = rest.split(" ", 1)
                say(f"Sending message to {name}: {message}")
                # Opens WhatsApp Web with pre-filled message
                url = f"https://web.whatsapp.com/send?text={message}"
                webbrowser.open(url)
            else:
                say("Please specify the message after the contact name.")
        else:
            say("Please specify the contact name and message.")
    else:
        say("Please specify who to send the message to.")

def call_someone(command):
    name = command.replace("call", "").strip()
    if name:
        say(f"Calling {name}...")
        # Simulate call (could open Skype/Teams if integrated)
        say(f"Sorry, calling is not supported on this device, but I can open Skype or Teams if you want.")
    else:
        say("Please specify who to call.")

def system_info():
    info = []
    info.append(f"System: {platform.system()}")
    info.append(f"Node Name: {platform.node()}")
    info.append(f"Release: {platform.release()}")
    info.append(f"Version: {platform.version()}")
    info.append(f"Machine: {platform.machine()}")
    info.append(f"Processor: {platform.processor()}")
    ram = psutil.virtual_memory()
    info.append(f"RAM: {ram.percent}% used")
    cpu = psutil.cpu_percent(interval=1)
    info.append(f"CPU Usage: {cpu}%")
    if hasattr(psutil, "sensors_battery"):
        battery = psutil.sensors_battery()
        if battery:
            info.append(f"Battery: {battery.percent}% {'Plugged in' if battery.power_plugged else 'On battery'}")
        else:
            info.append("Battery: Not available")
    say("\n".join(info))
    print("\n".join(info))

# ---------- MAIN LOOP ----------
say("Hello! I am Bagley v3. Ready for your commands.")

low_battery_warned = False

while True:
    # Battery warning (runs every loop)
    if hasattr(psutil, "sensors_battery"):
        battery = psutil.sensors_battery()
        if battery and not battery.power_plugged and battery.percent <= 20:
            if not low_battery_warned:
                say(f"Warning! Battery is low: {battery.percent} percent remaining.")
                low_battery_warned = True
        elif battery and battery.percent > 20:
            low_battery_warned = False

    query = listen()
    if not query:
        continue

    if "hey bagley" in query or "bagley" in query:
        say("Yes, I am listening!")
        continue

    if "system info" in query or "system information" in query:
        system_info()
    elif "volume" in query:
        change_volume(query)
    elif "open" in query or "launch" in query:
        open_application(query)
    elif "close" in query or "exit" in query:
        close_application(query)
    elif "open folder" in query:
        open_folder(query)
    elif any(word in query for word in ["shutdown", "restart", "log off"]):
        system_control(query)
    elif "clipboard" in query:
        clipboard_interaction(query)
    elif "window" in query:
        manage_windows(query)
    elif "search" in query:
        google_search(query)
    elif "type" in query:
        type_text(query)
    elif "scroll" in query:
        scroll_screen(query)
    elif "send message" in query:
        send_message(query)
    elif "call" in query:
        call_someone(query)
    elif "reset chat" in query:
        chatStr = ""
        save_memory("---- Memory Reset ----")
        say("Chat history has been reset")
    elif "quit" in query or "exit" in query:
        say("Bye, see you!")
        break
    else:
        say("Sorry, I can't answer questions without AI features enabled.")