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

    for proc in psutil.process_iter(["pid", "name", "exe", "cpu_times", "memory_info", "memory_percent",
                                     "num_ctx_switches", "num_handles", "num_threads", "io_counters", "cmdline",
                                     "nice", "open_files", "cpu_percent"]):
        if proc.info['pid'] == 0:
            continue  # skip System Idle Process on Windows
        proc_dict = proc.info.copy()

        # Flatten nested structures
        proc_dict["read_count"] = proc_dict["io_counters"].read_count
        proc_dict["write_count"] = proc_dict["io_counters"].write_count
        proc_dict["read_bytes"] = proc_dict["io_counters"].read_bytes
        proc_dict["write_bytes"] = proc_dict["io_counters"].write_bytes
        del proc_dict["io_counters"]
        proc_dict["cpu_times_user"] = proc_dict["cpu_times"].user
        proc_dict["cpu_times_system"] = proc_dict["cpu_times"].system
        del proc_dict["cpu_times"]
        proc_dict["voluntary_ctx_switches"] = proc_dict["num_ctx_switches"].voluntary
        proc_dict["involuntary_ctx_switches"] = proc_dict["num_ctx_switches"].involuntary
        del proc_dict["num_ctx_switches"]
        proc_dict["memory_rss"] = proc_dict["memory_info"].rss / (1024 ** 2)  # in MB
        proc_dict["memory_vms"] = proc_dict["memory_info"].vms / (1024 ** 2)  # in MB
        del proc_dict["memory_info"]
        
        # you have to call cpu_percent two times to get a valid reading
        proc.cpu_percent()
        time.sleep(0.1)
        proc_dict["cpu_percent"] = proc.cpu_percent()

        # Assign labels based on PID
        if proc_dict['pid'] in keylogger_pid_list:
            proc_dict["label"] = "keylogger"
        else:
            proc_dict["label"] = "benign"

        lst.append(proc_dict)

    df = pd.DataFrame(lst)
    df.to_csv("process_memory_usage.csv", index=False)

    print("time taken to get memory usage of all processes: ", time.time() - start)

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
    
# get_process_memory_usage()
# generate_boxplots()
# extract_keylogger_data()
