import os
import psutil
import time
import pandas as pd
import threading
from datetime import datetime
import subprocess
import re

class KeyloggerDataCollector:
    """
    Enhanced data collection system for keylogger behavioral analysis.
    
    Automatically detects keylogger processes and collects comprehensive
    system metrics for ML training and detection development.
    """

    def __init__(self, output_file="behavioral_data.csv", scan_interval=2):
        self.output_file = output_file
        self.scan_interval = scan_interval
        self.running = False
        self.keylogger_pids = set()
        self.suspicious_keywords = [
            'keylog', 'keylogger', 'pynput', 'pyhook', 'keystroke', 
            'logkeys', 'klog', 'hook', 'keyboard', 'keystroke'
        ]
        
        # Initialize data structures
        self._init_data_file()

    def _init_data_file(self):
        """creates the CSV data file with comprehensive headers"""
        if not os.path.exists(self.output_file):
            headers = [
                "timestamp", "pid", "process_name", "is_keylogger",
                "cpu_percent", "memory_rss_mb", "memory_vms_mb", "memory_percent",
                "num_threads", "num_handles", "read_bytes", "write_bytes",
                "read_count", "write_count", "voluntary_ctx_switches", 
                "involuntary_ctx_switches", "cpu_times_user", "cpu_times_system",
                "cmdline", "exe", "nice", "open_files_count"
            ]
            df = pd.DataFrame(columns=headers)
            df.to_csv(self.output_file, index=False)
            print(f"[INFO] Data file initialized: {self.output_file}")

    def detect_keylogger_processes(self):
        """
        Automatically detects potential keylogger processes using multiple methods.
        Returns list of suspicious PIDs.
        """
        suspicious_pids = set()
        
        for proc in psutil.process_iter(['pid', 'name', 'cmdline', 'exe']):
            try:
                proc_info = proc.info
                pid = proc_info['pid']
                
                # Skip system processes
                if pid in (0, 4) or proc_info['name'] in ['System', 'Idle']:
                    continue
                
                # Method 1: Check process name and command line for suspicious keywords
                is_suspicious = self._check_suspicious_keywords(proc_info)
                
                # Method 2: Check for keyboard-related system calls (Windows specific)
                if self._check_keyboard_hooks(pid):
                    is_suspicious = True
                
                # Method 3: Check for hidden windows or minimal UI
                if self._check_stealth_characteristics(proc_info):
                    is_suspicious = True
                
                if is_suspicious:
                    suspicious_pids.add(pid)
                    print(f"[DETECTED] Potential keylogger: PID {pid}, Name: {proc_info['name']}")
                    
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        return list(suspicious_pids)

    def _check_suspicious_keywords(self, proc_info):
        """checks process name, cmdline, and exe path for suspicious keywords"""
        # Safely handle cmdline which can be None
        cmdline = proc_info.get('cmdline', [])
        if cmdline is None:
            cmdline = []
        
        search_fields = [
            proc_info.get('name', '').lower(),
            ' '.join(str(arg) for arg in cmdline).lower(),
            proc_info.get('exe', '').lower()
        ]
        
        for field in search_fields:
            for keyword in self.suspicious_keywords:
                if keyword in field:
                    return True
        return False

    def _check_keyboard_hooks(self, pid):
        """checks if process might be using keyboard hooks (Windows)"""
        try:
            # This is a simplified check - in real implementation you'd use Windows API
            # For now, we'll check for Python processes with keyboard-related imports
            if psutil.Process(pid).name().lower() == 'python.exe':
                return True
        except:
            pass
        return False

    def _check_stealth_characteristics(self, proc_info):
        """checks for stealth characteristics like hidden windows, minimal resource usage"""
        try:
            proc = psutil.Process(proc_info['pid'])
            
            # Check for very low CPU but running
            cpu_percent = proc.cpu_percent()
            memory_info = proc.memory_info()
            
            # Suspicious: Low CPU but active, or Python process without visible window
            if (proc_info['name'].lower() == 'python.exe' and 
                memory_info.rss < 50 * 1024 * 1024):  # Less than 50MB RAM
                return True
                
        except:
            pass
        return False

    def collect_process_metrics(self):
        """
        Collects comprehensive metrics for all running processes.
        Returns list of process data dictionaries.
        """
        process_data = []
        current_pids = set()
        
        for proc in psutil.process_iter(["pid", "name", "exe", "cpu_times", "memory_info", 
                                        "memory_percent", "num_ctx_switches", "num_handles", 
                                        "num_threads", "io_counters", "cmdline", "nice"]):
            try:
                if proc.info['pid'] == 0:
                    continue  # skip System Idle Process
                
                proc_dict = proc.info.copy()
                pid = proc_dict['pid']
                current_pids.add(pid)
                
                # Flatten nested structures
                proc_dict = self._flatten_process_info(proc_dict)
                
                # Label as keylogger or benign
                proc_dict["is_keylogger"] = 1 if pid in self.keylogger_pids else 0
                proc_dict["timestamp"] = datetime.now().isoformat()
                
                process_data.append(proc_dict)
                
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        # Update keylogger PIDs (remove processes that no longer exist)
        self.keylogger_pids = self.keylogger_pids.intersection(current_pids)
        
        return process_data

    def _flatten_process_info(self, proc_dict):
        """flattens nested process information structures"""
        # Handle cmdline safely
        cmdline = proc_dict.get("cmdline", [])
        if cmdline is None:
            cmdline = []
        proc_dict["cmdline"] = ' '.join(str(arg) for arg in cmdline)
        
        # Handle IO counters
        if proc_dict.get("io_counters"):
            io = proc_dict["io_counters"]
            proc_dict["read_count"] = io.read_count if io else 0
            proc_dict["write_count"] = io.write_count if io else 0
            proc_dict["read_bytes"] = io.read_bytes if io else 0
            proc_dict["write_bytes"] = io.write_bytes if io else 0
            del proc_dict["io_counters"]
        else:
            proc_dict.update({"read_count": 0, "write_count": 0, "read_bytes": 0, "write_bytes": 0})
        
        # Handle CPU times
        if proc_dict.get("cpu_times"):
            cpu_times = proc_dict["cpu_times"]
            proc_dict["cpu_times_user"] = cpu_times.user if cpu_times else 0
            proc_dict["cpu_times_system"] = cpu_times.system if cpu_times else 0
            del proc_dict["cpu_times"]
        else:
            proc_dict.update({"cpu_times_user": 0, "cpu_times_system": 0})
        
        # Handle context switches
        if proc_dict.get("num_ctx_switches"):
            ctx = proc_dict["num_ctx_switches"]
            proc_dict["voluntary_ctx_switches"] = ctx.voluntary if ctx else 0
            proc_dict["involuntary_ctx_switches"] = ctx.involuntary if ctx else 0
            del proc_dict["num_ctx_switches"]
        else:
            proc_dict.update({"voluntary_ctx_switches": 0, "involuntary_ctx_switches": 0})
        
        # Handle memory info
        if proc_dict.get("memory_info"):
            mem = proc_dict["memory_info"]
            proc_dict["memory_rss_mb"] = mem.rss / (1024 ** 2) if mem else 0
            proc_dict["memory_vms_mb"] = mem.vms / (1024 ** 2) if mem else 0
            del proc_dict["memory_info"]
        else:
            proc_dict.update({"memory_rss_mb": 0, "memory_vms_mb": 0})
        
        # Get CPU percentage (need to call twice for accurate reading)
        try:
            proc = psutil.Process(proc_dict['pid'])
            proc.cpu_percent()  # First call to initialize
            time.sleep(0.1)
            proc_dict["cpu_percent"] = proc.cpu_percent()
        except:
            proc_dict["cpu_percent"] = 0
        
        # Count open files
        try:
            proc = psutil.Process(proc_dict['pid'])
            proc_dict["open_files_count"] = len(proc.open_files())
        except:
            proc_dict["open_files_count"] = 0
        
        return proc_dict

    def start_collection(self, duration=300):
        """
        Starts the data collection process.
        
        Args:
            duration (int): How long to collect data in seconds
        """
        self.running = True
        start_time = time.time()
        collection_count = 0
        
        print(f"[INFO] Starting data collection for {duration} seconds...")
        print("[INFO] Press Ctrl+C to stop early")
        
        try:
            while self.running and (time.time() - start_time) < duration:
                # Detect keylogger processes
                new_keyloggers = self.detect_keylogger_processes()
                self.keylogger_pids.update(new_keyloggers)
                
                # Collect metrics
                process_data = self.collect_process_metrics()
                
                # Append to CSV
                if process_data:
                    df = pd.DataFrame(process_data)
                    df.to_csv(self.output_file, mode='a', header=False, index=False)
                    collection_count += 1
                    print(f"[COLLECTED] Scan {collection_count}: {len(process_data)} processes, "
                          f"{len(self.keylogger_pids)} keyloggers detected")
                else:
                    print(f"[COLLECTED] Scan {collection_count}: No processes collected")
                
                time.sleep(self.scan_interval)
                
        except KeyboardInterrupt:
            print("\n[INFO] Data collection stopped by user")
        except Exception as e:
            print(f"[ERROR] Collection error: {e}")
        finally:
            self.running = False
            print(f"[INFO] Data collection completed. Total scans: {collection_count}")
            print(f"[INFO] Data saved to: {self.output_file}")

    def analyze_collected_data(self):
        """provides basic analysis of the collected data"""
        if not os.path.exists(self.output_file):
            print("[ERROR] No data file found. Run collection first.")
            return
        
        try:
            df = pd.read_csv(self.output_file)
            print(f"\n[DATA ANALYSIS]")
            print(f"Total records: {len(df)}")
            print(f"Total processes monitored: {df['pid'].nunique()}")
            
            # Fix: Ensure is_keylogger is numeric for the calculation
            df['is_keylogger'] = pd.to_numeric(df['is_keylogger'], errors='coerce').fillna(0)
            
            keylogger_count = int(df['is_keylogger'].sum())
            print(f"Keylogger samples: {keylogger_count}")
            print(f"Benign samples: {len(df) - keylogger_count}")
            
            if keylogger_count > 0:
                keylogger_stats = df[df['is_keylogger'] == 1].describe()
                print("\nKeylogger process statistics:")
                print(keylogger_stats[['cpu_percent', 'memory_rss_mb', 'num_threads']])
        except Exception as e:
            print(f"[ERROR] Analysis failed: {e}")



# Test function
def test_data_collection():
    """Test function for the data collection system"""
    print("Testing Keylogger Data Collection System...")
    
    collector = KeyloggerDataCollector(
        output_file="test_behavioral_data.csv",
        scan_interval=3  # Scan every 3 seconds
    )
    
    # Run for 60 seconds for testing
    collector.start_collection(duration=60)
    
    # Analyze the collected data
    collector.analyze_collected_data()

if __name__ == "__main__":
    test_data_collection()