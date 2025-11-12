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
            if proc.info['pid'] == 0:
                continue  
            
            proc_dict = proc.info.copy()

            try:
                proc_dict["read_count"] = proc_dict["io_counters"].read_count
                proc_dict["write_count"] = proc_dict["io_counters"].write_count
                proc_dict["read_bytes"] = proc_dict["io_counters"].read_bytes
                proc_dict["write_bytes"] = proc_dict["io_counters"].write_bytes
            except (AttributeError, psutil.NoSuchProcess):
                proc_dict["read_count"] = 0
                proc_dict["write_count"] = 0
                proc_dict["read_bytes"] = 0
                proc_dict["write_bytes"] = 0
            del proc_dict["io_counters"]

            try:
                proc_dict["cpu_times_user"] = proc_dict["cpu_times"].user
                proc_dict["cpu_times_system"] = proc_dict["cpu_times"].system
            except (AttributeError, psutil.NoSuchProcess):
                proc_dict["cpu_times_user"] = 0
                proc_dict["cpu_times_system"] = 0
            del proc_dict["cpu_times"]

            try:
                proc_dict["voluntary_ctx_switches"] = proc_dict["num_ctx_switches"].voluntary
                proc_dict["involuntary_ctx_switches"] = proc_dict["num_ctx_switches"].involuntary
            except (AttributeError, psutil.NoSuchProcess):
                proc_dict["voluntary_ctx_switches"] = 0
                proc_dict["involuntary_ctx_switches"] = 0
            del proc_dict["num_ctx_switches"]

            try:
                proc_dict["memory_rss"] = proc_dict["memory_info"].rss / (1024 ** 2)  # in MB
                proc_dict["memory_vms"] = proc_dict["memory_info"].vms / (1024 ** 2)  # in MB
            except (AttributeError, psutil.NoSuchProcess):
                proc_dict["memory_rss"] = 0
                proc_dict["memory_vms"] = 0
            del proc_dict["memory_info"]
            
            try:
                proc.cpu_percent()
                time.sleep(0.5)  
                proc_dict["cpu_percent"] = proc.cpu_percent()
            except psutil.NoSuchProcess:
                proc_dict["cpu_percent"] = 0

            
            if proc_dict['pid'] in keylogger_pid_list:
                proc_dict["label"] = "keylogger"
            else:
                proc_dict["label"] = "benign"

            lst.append(proc_dict)
            processed_count += 1
            
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            error_count += 1
            continue

    df = pd.DataFrame(lst)
    df.to_csv("process_memory_usage.csv", index=False)

    print(f"Successfully processed {processed_count} processes, {error_count} errors")
    print("Time taken to get memory usage of all processes: ", time.time() - start)
    return df

def analyze_file_names(proc_data):
    # do stuff with exe, name, cmdline, open_files
    pass

def standardize_features(df):
    """Ensure consistent feature order and types across all data"""
    expected_features = [
        'cpu_percent', 'memory_percent', 'num_handles', 'num_threads', 'nice',
        'read_count', 'write_count', 'read_bytes', 'write_bytes', 
        'cpu_times_user', 'cpu_times_system', 'voluntary_ctx_switches',
        'involuntary_ctx_switches', 'memory_rss', 'memory_vms'
    ]
    
    # Create missing columns with default values
    for feature in expected_features:
        if feature not in df.columns:
            df[feature] = 0
    
    # Select and reorder columns
    return df[expected_features]

def ml_detector_with_confidence(proc_data):
    loaded_rf = joblib.load("./random_forest.joblib")
    expected_features = loaded_rf.feature_names_in_
    proc_data_clean = proc_data[expected_features].copy()
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
