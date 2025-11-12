import os
from matplotlib import pyplot as plt
import psutil
import time
import pandas as pd
import seaborn as sns

# TODO : populate this list with actual keylogger PIDs
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

def print_data(file):
    df = pd.read_csv(file)
    print()
    print(df.head())
    print(df.info())
    print(df.describe())
    print(df[df['nice'] == pd.NA])

def extract_keylogger_data():
    key_df  = pd.DataFrame()
    benign_df  = pd.DataFrame()
    with os.scandir("data") as es:
        for e in es:
            if e.is_file() and e.name.endswith('.csv'):
                df = pd.read_csv(e.path)

                # get keylogger data
                keylogger_data = df[df['label'] == 'keylogger']
                if key_df.empty:
                    key_df = keylogger_data
                else:
                    key_df = pd.concat([key_df, keylogger_data], ignore_index=True)

                # get benign data
                benign_data = df[df['label'] == 'keylogger']
                if benign_df.empty:
                    benign_df = benign_data
                else:
                    benign_df = pd.concat([benign_df, benign_data], ignore_index=True)
    key_df.to_csv('benign_data.csv', index=False)
    benign_df.to_csv('benign_data.csv', index=False)

def generate_boxplots():
    df = pd.read_csv("keylogger_data.csv")
    df2 = pd.read_csv("benign_data.csv")
    df = df.drop(columns=['exe', 'cmdline', 'pid', 'name', 'open_files'])
    df2 = df2.drop(columns=['exe', 'cmdline', 'pid', 'name', 'open_files'])
    print(len(df.select_dtypes(include=['float64', 'int64']).columns))

    plt.figure(figsize=(10, 6))

    for i, column in enumerate(df.select_dtypes(include=['float64', 'int64']).columns):  
        # sns.histplot(df[column], kde=True)
        print(column, i,i//3, i%3)
        plt.title(f'Distribution of {column}')
        # sns.histplot(df[column], color='red', label='keylogger', fill=True, alpha=0.5, ax=axes[i//3][i%3])
        # sns.histplot(df2[column], color='blue', label='benign', fill=True, alpha=0.5, ax=axes[i//3][i%3])
        plt.boxplot([df[column], df2[column]], labels=['keylogger', 'benign'], showfliers=False)
        print()
        plt.savefig("graphs/"+column+"_boxplot.png")
        plt.clf()

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

def continuous_data_collection(duration_minutes=60):
    """Run data collection continuously"""
    from datetime import datetime
    all_data = []
    start_time = time.time()
    
    print(f"Starting continuous data collection for {duration_minutes} minutes...")
    
    while time.time() - start_time < duration_minutes * 60:
        print(f"Collection cycle at {datetime.now().strftime('%H:%M:%S')}")
        current_data = get_process_memory_usage()
        all_data.append(current_data)
        print(f"Collected {len(current_data)} process samples")
        time.sleep(30)  
    
    # Combine and save all data
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        filename = f"continuous_data_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
        final_df.to_csv(filename, index=False)
        print(f"Continuous collection complete! Saved to {filename}")
    
    return all_data
    
# get_process_memory_usage()
# generate_boxplots()
# extract_keylogger_data()
continuous_data_collection(duration_minutes=5)
