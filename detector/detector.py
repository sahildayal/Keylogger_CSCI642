import joblib
import psutil
import time
import os
from sklearn.ensemble import RandomForestClassifier
import pandas as pd

def get_process_memory_usage():
    lst = []
    start = time.time()

    for proc in psutil.process_iter(["pid", "name", "exe", "cpu_times", "memory_info", "memory_percent",
                                     "num_ctx_switches", "num_handles", "num_threads", "io_counters", "cmdline",
                                     "nice", "open_files", "cpu_percent"]):
        if proc.info['pid'] == 0 or proc.info['pid'] == os.getpid():
            continue  # skip System Idle Process on Windows and our program
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

        lst.append(proc_dict)

    # df = pd.DataFrame(columns=["name","num_handles","num_threads","nice","memory_percent","cpu_percent",
    #                            "pid","read_count","write_count","read_bytes","write_bytes","cpu_times_user","cpu_times_system",
    #                            "voluntary_ctx_switches","involuntary_ctx_switches","memory_rss","memory_vms",
    #                            "exe","cmdline","open_files"])
    df = pd.read_csv("data/keylogger_data.csv")
    df = df.iloc[:0]
    df = pd.DataFrame(lst)
 

    print("time taken to get memory usage of all processes: ", time.time() - start)
    return df

def analyze_file_names(proc_data):
    # do stuff with exe, name, cmdline, open_files
    pass

def ml_detector(proc_data):
    loaded_rf = joblib.load("./random_forest.joblib")
    rf = loaded_rf.predict(proc_data.drop(['pid', 'name', 'exe', 'cmdline', 'open_files'], axis=1))
    detect = False
    for i, label in enumerate(rf):
        if label == 'keylogger':
            print("Keylogger detected!")
            print(f"process name: {proc_data.iloc[i].name}, pid: {proc_data.iloc[i].pid}")
            detect = True
    if not detect:
        print("No keyloggers detected.")

if __name__ == "__main__":
    proc_data = get_process_memory_usage()
    ml_detector(proc_data)