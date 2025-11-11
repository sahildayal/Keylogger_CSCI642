import pandas as pd
import psutil
import time
from datetime import datetime

def collect_clean_data():
    """Collects clean, properly labeled data for training"""
    print("[INFO] Collecting clean training data...")
    
    # Initialize data structure
    data = []
    
    # Collect data for 30 seconds
    start_time = time.time()
    scan_count = 0
    
    while time.time() - start_time < 30:
        scan_count += 1
        current_pids = set()
        
        for proc in psutil.process_iter(['pid', 'name', 'cpu_percent', 'memory_info', 
                                        'num_threads', 'num_handles', 'io_counters']):
            try:
                proc_info = proc.info
                pid = proc_info['pid']
                
                if pid in [0, 4]:
                    continue
                
                current_pids.add(pid)
                
                # Get CPU percentage properly
                proc.cpu_percent()
                time.sleep(0.05)
                cpu_percent = proc.cpu_percent()
                
                # Label keyloggers (Python processes for demo)
                is_keylogger = 1 if 'python' in proc_info['name'].lower() else 0
                
                # Collect metrics
                process_data = {
                    'timestamp': datetime.now().isoformat(),
                    'pid': pid,
                    'process_name': proc_info['name'],
                    'is_keylogger': is_keylogger,
                    'cpu_percent': cpu_percent,
                    'memory_rss_mb': proc_info['memory_info'].rss / (1024 ** 2) if proc_info['memory_info'] else 0,
                    'num_threads': proc_info['num_threads'],
                    'num_handles': proc_info['num_handles'],
                    'read_bytes': proc_info['io_counters'].read_bytes if proc_info['io_counters'] else 0,
                    'write_bytes': proc_info['io_counters'].write_bytes if proc_info['io_counters'] else 0,
                }
                
                data.append(process_data)
                
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        print(f"[SCAN {scan_count}] Collected {len([d for d in data if d['pid'] in current_pids])} process samples")
        time.sleep(2)
    
    # Save to CSV
    df = pd.DataFrame(data)
    df.to_csv('clean_training_data.csv', index=False)
    
    keylogger_count = df['is_keylogger'].sum()
    print(f"\n[INFO] Data collection complete!")
    print(f"Total samples: {len(df)}")
    print(f"Keylogger samples: {keylogger_count}")
    print(f"Benign samples: {len(df) - keylogger_count}")
    print(f"Saved to: clean_training_data.csv")
    
    return df

if __name__ == "__main__":
    collect_clean_data()