import psutil
import time
import pandas as pd
import sys
import ast
import datetime
import os

# # TODO : populate this list with actual keylogger PIDs
# keylogger_pid_list = []

def save_file():
    # path stuff
    phase3_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
    data_dir = os.path.join(phase3_dir, "..", "data")

    # make sure it exists
    os.makedirs(data_dir, exist_ok=True)
    save_file = os.path.join(data_dir, "snapshot.csv")
    return save_file

def snapshot(keylogger_pid_list=[]):
    sf = save_file()

    lst = []
    start = time.time()
    error_count = 0

    for proc in psutil.process_iter(["pid", "name", "exe", "cpu_times", "memory_info", "memory_percent",
                                     "num_ctx_switches", "num_handles", "num_threads", "io_counters", "cmdline",
                                     "nice", "open_files", "cpu_percent"]):
        
        if proc.info['pid'] == 0:
            continue  # skip System Idle Process on Windows
        proc_dict = proc.info.copy()

        # Flatten nested structures
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
        
        # you have to call cpu_percent two times to get a valid reading
        try:
            proc.cpu_percent()
            time.sleep(0.5)  
            proc_dict["cpu_percent"] = proc.cpu_percent()
        except psutil.NoSuchProcess:
            proc_dict["cpu_percent"] = 0

        # Assign labels based on PID
        if len(keylogger_pid_list) > 0:
            if proc_dict['pid'] in keylogger_pid_list:
                proc_dict["label"] = "keylogger"
            else:
                proc_dict["label"] = "benign"

        lst.append(proc_dict)

    df = pd.DataFrame(lst)
    df.to_csv(sf, index=False)

    print(f"Successfully processed {len(lst)} processes, {error_count} errors")
    print("Time taken to get memory usage of all processes: ", time.time() - start)

    return df

def print_data():
    sf = save_file()
    df = pd.read_csv(sf)
    print()
    print(df.head())
    print(df.info())
    print(df.describe())
    # print(df.avg_cpu_percent.describe())
    max_row = df[df['cpu_percent'] == df['cpu_percent'].max()]
    # max_row = df[df['avg_cpu_percent'] == df['avg_cpu_percent'].max()]
    print(max_row)
    print("Max CPU Percent: ", df['cpu_percent'].max())
    
# def continuous_data_collection(duration_minutes=60):
#     """Run data collection continuously"""
#     from datetime import datetime
#     all_data = []
#     start_time = time.time()
    
#     print(f"Starting continuous data collection for {duration_minutes} minutes...")
    
#     while time.time() - start_time < duration_minutes * 60:
#         print(f"Collection cycle at {datetime.now().strftime('%H:%M:%S')}")
#         current_data = snapshot()
#         all_data.append(current_data)
#         print(f"Collected {len(current_data)} process samples")
#         time.sleep(30)  
    
#     # Combine and save all data
#     if all_data:
#         final_df = pd.concat(all_data, ignore_index=True)
#         filename = f"continuous_data_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
#         final_df.to_csv(filename, index=False)
#         print(f"Continuous collection complete! Saved to {filename}")
    
#     return all_data

def continuous_data_collection(duration_seconds=60, interval_seconds=5, output_file="../../data/raw_behavior_data.csv"):

    print(f"[INFO] Collecting raw process data for {duration_seconds}s...")
    start = time.time()
    all_data = []

    while time.time() - start < duration_seconds:
        snap = snapshot()
        all_data.append(snap)
        print(f"[INFO] Snapshot at {datetime.now().strftime('%H:%M:%S')} → {len(snap)} processes")
        time.sleep(interval_seconds)

    final_df = pd.concat(all_data, ignore_index=True)

    if os.path.exists(output_file):
        final_df.to_csv(output_file, mode='a', header=False, index=False)
    else:
        final_df.to_csv(output_file, index=False)

    print(f"[INFO] Saved {len(final_df)} rows to {output_file}")


if __name__ == "__main__":
    """
    Run Options:
        1st arg val - method parameters would be the following args
        examples. 
            python .\\generate_data.py 1
            python .\\generate_data.py 2 "[199, 200, 201]"
            python .\\generate_data.py 4 300 5

        (1) - snapshot()
        (2) - snapshot([<pid>, <pid>, ...])
        (3) - continuous_data_collection()
        (4) - continuous_data_collection(duration_seconds, interval_seconds)
    """
    if len(sys.argv) > 1:
        try:
            call = int(sys.argv[1])
            if call == 1:
                snapshot()
            elif call == 2:
                input_list = ast.literal_eval(sys.argv[2])
                if isinstance(input_list, list):
                    snapshot(input_list)
                else:
                    print("Error: Was not given a valid list as arg 2")
            elif call == 3:
                continuous_data_collection()
            elif call == 4:
                d = int(sys.argv[2])
                i = int(sys.argv[3])
                continuous_data_collection(d, i)
            else:
                print("The first argument must be 1 - 4")
        except Exception as e:
            print("Problem with args provided: ", e)
    else:
        print("Please provide at least 1 arg")