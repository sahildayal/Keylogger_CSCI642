import joblib
import psutil
import time
import os
from sklearn.ensemble import RandomForestClassifier
import pandas as pd

keylogger_pid_list = []

def get_process_memory_usage():
    lst = []
    start = time.time()
    processed_count = 0
    error_count = 0

    for proc in psutil.process_iter(["pid", "name", "exe", "cpu_times", "memory_info", "memory_percent",
                                     "num_ctx_switches", "num_handles", "num_threads", "io_counters", "cmdline",
                                     "nice", "open_files", "cpu_percent"]):
        try:
            if proc.info['pid'] == 0 or proc.info['pid'] == 4:  # Skip System processes
                continue
            
            proc_dict = {}
            
            # Store basic info
            proc_dict["pid"] = proc.info['pid']
            proc_dict["name"] = proc.info['name']
            proc_dict["exe"] = proc.info.get('exe', '')
            proc_dict["cmdline"] = ' '.join(proc.info.get('cmdline', [])) if proc.info.get('cmdline') else ''
            proc_dict["open_files"] = len(proc.info.get('open_files', [])) if proc.info.get('open_files') is not None else 0
            
            # CPU and Memory features (CRITICAL - model needs these)
            try:
                proc.cpu_percent()  # Initialize
                time.sleep(0.1)
                proc_dict["cpu_percent"] = proc.cpu_percent()
            except:
                proc_dict["cpu_percent"] = 0
                
            proc_dict["memory_percent"] = proc.info.get('memory_percent', 0)
            
            # Memory info
            try:
                memory_info = proc.info['memory_info']
                proc_dict["memory_rss"] = memory_info.rss / (1024 ** 2) if memory_info else 0
                proc_dict["memory_vms"] = memory_info.vms / (1024 ** 2) if memory_info else 0
            except:
                proc_dict["memory_rss"] = 0
                proc_dict["memory_vms"] = 0
            
            # Process characteristics
            proc_dict["num_handles"] = proc.info.get('num_handles', 0)
            proc_dict["num_threads"] = proc.info.get('num_threads', 0)
            proc_dict["nice"] = proc.info.get('nice', 0)
            
            # I/O counters
            try:
                io = proc.info['io_counters']
                proc_dict["read_count"] = io.read_count if io else 0
                proc_dict["write_count"] = io.write_count if io else 0
                proc_dict["read_bytes"] = io.read_bytes if io else 0
                proc_dict["write_bytes"] = io.write_bytes if io else 0
            except:
                proc_dict["read_count"] = 0
                proc_dict["write_count"] = 0
                proc_dict["read_bytes"] = 0
                proc_dict["write_bytes"] = 0
            
            # CPU times
            try:
                cpu_times = proc.info['cpu_times']
                proc_dict["cpu_times_user"] = cpu_times.user if cpu_times else 0
                proc_dict["cpu_times_system"] = cpu_times.system if cpu_times else 0
            except:
                proc_dict["cpu_times_user"] = 0
                proc_dict["cpu_times_system"] = 0
            
            # Context switches
            try:
                ctx_switches = proc.info['num_ctx_switches']
                proc_dict["voluntary_ctx_switches"] = ctx_switches.voluntary if ctx_switches else 0
                proc_dict["involuntary_ctx_switches"] = ctx_switches.involuntary if ctx_switches else 0
            except:
                proc_dict["voluntary_ctx_switches"] = 0
                proc_dict["involuntary_ctx_switches"] = 0

            lst.append(proc_dict)
            processed_count += 1
            
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            error_count += 1
            continue

    df = pd.DataFrame(lst)
    print(f"Successfully processed {processed_count} processes, {error_count} errors")
    print("Time taken to get memory usage of all processes: ", time.time() - start)
    return df

def analyze_file_names(proc_data):
    # do stuff with exe, name, cmdline, open_files
    pass

def standardize_features(df):
    """Fix column names to match the trained model"""
    # Rename columns to match what the model expects
    column_mapping = {
        'memory_rss': 'memory_rss_mb',
        'memory_vms': 'memory_vms_mb',
        # Add other mappings if needed
    }
    df = df.rename(columns=column_mapping)
   
    # Get expected features from the model
    try:
        loaded_rf = joblib.load("./retrained_detector.pkl")
        expected_features = loaded_rf.feature_names_in_
       
        # Create missing columns with default values
        for feature in expected_features:
            if feature not in df.columns:
                df[feature] = 0
       
        return df[expected_features]
    except:
        return df
    
def ml_detector_with_confidence(proc_data):
    loaded_rf = joblib.load("./retrained_detector.pkl")
    proc_data_clean = standardize_features(proc_data)
    probabilities = loaded_rf.predict_proba(proc_data_clean)
    
    detect = False
    for i, (prob_benign, prob_keylogger) in enumerate(probabilities):
        if prob_keylogger > 0.7:  #70% confidence threshold
            print(f"Keylogger detected! (Confidence: {prob_keylogger:.1%})")
            print(f"   Process: {proc_data.iloc[i].name}, PID: {proc_data.iloc[i].pid}")
            detect = True
        elif prob_keylogger > 0.4:  #suspicious but not certain
            print(f"Suspicious process: {proc_data.iloc[i]['name']} (PID: {proc_data.iloc[i]['pid']}) (Confidence: {prob_keylogger:.1%})")
            # print(f"Suspicious process: {proc_data.iloc[i].name} (Confidence: {prob_keylogger:.1%})")
    
    if not detect:
        print("No keyloggers detected")
    
    return detect

if __name__ == "__main__":
    proc_data = get_process_memory_usage()
    ml_detector_with_confidence(proc_data)
