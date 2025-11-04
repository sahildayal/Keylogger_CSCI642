import os
import psutil
import time
import pandas as pd

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

def print_data():
    df = pd.read_csv("process_memory_usage.csv")
    print()
    print(df.head())
    print(df.info())
    print(df.describe())
    # print(df.avg_cpu_percent.describe())
    max_row = df[df['cpu_percent'] == df['cpu_percent'].max()]
    # max_row = df[df['avg_cpu_percent'] == df['avg_cpu_percent'].max()]
    print(max_row)
    print("Max CPU Percent: ", df['cpu_percent'].max())
    

get_process_memory_usage()
# print_data()
