import pyttsx3
import speech_recognition as sr
import webbrowser
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import time
import ctypes
import os
import winsdk.windows.ui.notifications as notifications
import winsdk.windows.data.xml.dom as dom

# Initialize pyttsx3 and set voice properties
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# Global variable to store the user's preferred mode (speak or print)
mode = None


# Function to output text based on the chosen mode
def output(text):
    if mode == 'speak':
        engine.say(text)
        engine.runAndWait()
    elif mode == 'print':
        print(text)


# Function to take user input (either from voice or text)
def take_input():
    if mode == 'speak':
        r = sr.Recognizer()
        with sr.Microphone() as source:
            output("Listening...")
            r.pause_threshold = 1
            audio = r.listen(source)
        try:
            output("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            output(f"You said: {query}")
            return query.lower()
        except Exception:
            output("Say that again, please.")
            return "None"
    elif mode == 'print':
        return input("Enter your query: ").strip().lower()


# Function to open Windows settings
def open_settings():
    output("Opening Windows settings...")
    os.system("start ms-settings:")


# Function to handle PDF input and print summary if the file name is PAPER
def handle_pdf_input(file_path):
    # Get the file name (without extension)
    file_name = os.path.splitext(os.path.basename(file_path))[0]

    # Check if the file name is "PAPER" (case-insensitive)
    if file_name.lower() == "paper":
        # Summary to print
        summary = (
            "Natural fiber reinforced composites are gaining attention as sustainable alternatives to synthetic composites. "
            "This review article examines how various surface treatment methods applied to natural fibers affect the mechanical properties of composites. "
            "Treatments like NaOH, NaHCO3, acetyl acid, silane coupling agents, and maleic anhydride-polypropylene (MAPP) significantly improve flexural properties, "
            "interfacial adhesion, tensile strength, flexural strength, and impact resistance by removing impurities and enhancing chemical bonding. "
            "Alkaline treatments, especially with NaOH, effectively dissolve lignin and hemicellulose, leading to better fiber-matrix bonding and load transfer. "
            "Fiber orientation is another crucial factor, as aligned fibers parallel to applied force yield superior strength. "
            "Off-axis orientations may decrease performance, indicating the need to optimize fiber orientation for better properties. "
            "Overall, these surface treatments and fiber orientation strategies can lead to stronger, more durable natural fiber reinforced composites, "
            "paving the way for their wider use in various industries."
        )
        # Print the summary to console
        output("Summary of PAPER:")
        output(summary)


# Function to lock Windows
def lock_windows():
    output("Locking your workstation now.")
    ctypes.windll.user32.LockWorkStation()


# Function to search Google and print the results
def search_google(query):
    # Prepare the search URL
    url = f'https://www.google.com/search?q={query}'
    headers = {'User-Agent': 'Mozilla/5.0'}

    # Make the request
    response = requests.get(url, headers=headers)

    # Parse the HTML content
    soup = BeautifulSoup(response.text, 'html.parser')

    # Find the search result elements
    results = soup.find_all('div', class_='BNeawe vvjwJb AP7Wnd')[:3]  # Top 3 results
    links = soup.find_all('a')  # Find all anchor tags in the soup

    # Filter links to get top 3 result links
    search_results = []
    for link in links:
        href = link.get('href')
        if href and href.startswith('/url?'):
            # Strip URL to get the actual search result URL
            search_url = href.split('=')[1].split('&')[0]
            # Decode the URL
            search_url = requests.utils.unquote(search_url)
            search_results.append(search_url)
            if len(search_results) == 3:
                break

    # Check if there are results and print them
    if results and search_results:
        print("Here are the top search results for your query:")
        for i, (result, search_url) in enumerate(zip(results, search_results), start=1):
            print(f"Result {i}: {result.get_text()}")
            print(f"Link: {search_url}")
    else:
        print("Sorry, I couldn't find any search results for your query.")


# Function to set an alarm
def set_alarm(alarm_time):
    """Set an alarm for a specific time."""
    while True:
        current_time = datetime.now()
        if current_time >= alarm_time:
            output("Wake up! It's time!")
            break
        time.sleep(1)


# Function to set a timer
def set_timer(duration):
    """Set a timer for a specific duration."""
    output(f"Starting a timer for {duration}...")

    # Parse the duration and calculate total seconds
    if 'minute' in duration:
        minutes = int(duration.split()[0])
        total_seconds = minutes * 60
    elif 'second' in duration:
        total_seconds = int(duration.split()[0])
    else:
        output("Sorry, I didn't understand the timer duration.")
        return

    # Countdown the timer
    while total_seconds > 0:
        minutes, seconds = divmod(total_seconds, 60)
        print(f"Time left: {minutes:02d}:{seconds:02d}")  # Print time left to the console
        time.sleep(1)
        total_seconds -= 1

    # Timer ends, notify the user
    output("Time's up!")


# Function to create and display a Windows notification
def display_notification(title, message):
    """Create and display a Windows notification with the given title and message."""
    # Create a toast notifier
    toast_notifier = notifications.ToastNotificationManager.create_toast_notifier()

    # Create a toast content
    toast_xml = f"""
    <toast>
        <visual>
            <binding template='ToastGeneric'>
                <text>{title}</text>
                <text>{message}</text>
            </binding>
        </visual>
    </toast>
    """

    # Parse the XML content
    doc = dom.XmlDocument()
    doc.load_xml(toast_xml)

    # Create a toast notification
    toast = notifications.ToastNotification(doc)

    # Show the notification
    toast_notifier.show(toast)


# Main function to handle user queries and commands
def main():
    global mode
    # Ask the user for their preferred mode
    mode_preference = input("Choose your mode: type 'speak' or 'print': ").strip().lower()
    if mode_preference == 'speak':
        mode = 'speak'
        engine.say("You have chosen speak mode.")
        engine.runAndWait()
    elif mode_preference == 'print':
        mode = 'print'
        print("You have chosen print mode.")
    else:
        print("Invalid selection. Defaulting to print mode.")
        mode = 'print'

    # Start the main loop
    while True:
        query = take_input()  # Take user input based on the selected mode

        # Handle different queries
        if 'open settings' in query:
            open_settings()
        elif 'open youtube' in query:
            webbrowser.open("https://www.youtube.com")
            output("Opening YouTube...")
        elif 'open google' in query:
            webbrowser.open("https://www.google.com")
            output("Opening Google...")
        elif 'open github' in query:
            webbrowser.open("https://github.com/")
            output("Opening GitHub...")
        elif 'open grammarly' in query:
            webbrowser.open("https://www.grammarly.com/")
            output("Opening Grammarly...")
        elif 'open erp' in query:
            webbrowser.open("https://engg.dpuerp.in/Login.aspx")
            output("Opening ERP...")
        elif 'what is the capital of india' in query:
            output("Delhi")
        elif 'how are you today' in query:
            output("I'm doing well, thank you.")
        elif 'what is your name' in query:
            output("My name is Jarvis")
        elif 'set alarm' in query:
            output("At what time would you like the alarm? (HH:MM)")
            time_query = take_input()
            # Convert the query to a datetime object for the alarm time
            try:
                alarm_time = datetime.strptime(time_query, '%H:%M')
                # Get current date and set the alarm time
                now = datetime.now()
                alarm_time = now.replace(hour=alarm_time.hour, minute=alarm_time.minute, second=0, microsecond=0)
                # If the alarm time has already passed today, schedule it for the next day
                if alarm_time <= now:
                    alarm_time += timedelta(days=1)
                output(f"Alarm set for {alarm_time.strftime('%H:%M')}")
                set_alarm(alarm_time)
            except ValueError:
                output("Sorry, I didn't understand the time format. Please try again.")

        elif 'set timer' in query:
            output("For how long would you like to set the timer? (e.g., 5 minutes or 30 seconds)")
            duration_query = take_input()
            set_timer(duration_query)

        elif 'open gmail' in query:
            webbrowser.open("https://mail.google.com/mail/u/0/#inbox")
            output("Opening Gmail...")

        elif 'open meet' in query:
            webbrowser.open("https://meet.google.com/")
            output("Opening Google Meet...")

        elif 'open zoom' in query:
            webbrowser.open("https://zoom.us/signin#/login")
            output("Opening Zoom sign-in page...")

        elif 'open chrome' in query:
            # Specify the path to Chrome executable
            chrome_path = webbrowser.get('C:/Program Files/Google/Chrome/Application/chrome.exe %s')
            webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
            webbrowser.get(chrome_path).open("https://www.google.com")
            output("Opening Google Chrome...")

        elif 'open firefox' in query:
            # Specify the path to Firefox executable
            firefox_path = webbrowser.get('C:/Program Files/Mozilla Firefox/firefox.exe %s')
            webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(firefox_path))
            webbrowser.get(firefox_path).open("https://www.google.com")
            output("Opening Mozilla Firefox...")

        elif 'open scihub' in query:
            webbrowser.open("https://www.sci-hub.se/")
            output("Opening Sci-Hub...")

        elif 'open mendeley' in query:
            webbrowser.open("https://www.mendeley.com/?interaction_required=true")
            output("Opening Mendeley...")

        elif 'open my courses' in query:
            webbrowser.open("https://www.udemy.com/home/my-courses/learning/")
            output("Opening Udemy My Courses...")

        elif 'play music' in query:
            webbrowser.open("https://open.spotify.com/")
            output("Playing music on Spotify...")

        elif 'open whatsapp' in query:
            webbrowser.open("https://web.whatsapp.com/")
            output("Opening WhatsApp Web...")

        elif 'download video' in query:
            webbrowser.open("https://ssyoutube.com/en169ZG/youtube-video-downloader")
            output("Opening YouTube video downloader...")

        elif 'current weather' in query:
            output("For which location would you like to know the weather?")
            location_query = take_input()
            # Function to get the current weather data should be implemented
            # get_current_weather(location_query)

        elif 'search google' in query:
            output("What would you like to search for?")
            search_query = take_input()
            if search_query != "None":
                search_google(search_query)

        elif 'lock windows' in query:
            lock_windows()

        elif 'pdf input' in query:
            output("Please provide the file path for the PDF:")
            pdf_path = take_input()
            if pdf_path != "None":
                handle_pdf_input(pdf_path)

        elif 'stop' in query:
            output("Stopping the program...")
            break

        elif 'bye' in query:
            output("Goodbye!")
            break

        elif 'display notification' in query:
            output("Please provide the title of the notification:")
            notification_title = take_input()
            output("Call me its  urgent")

        else:
            output("I'm sorry, I don't understand the command.")


if __name__ == '__main__':
    main()

