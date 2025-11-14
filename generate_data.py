import psutil
import time
import pandas as pd
import os
from datetime import datetime

FEATURE_COLUMNS = [
    'cpu_percent', 'memory_percent', 'num_handles', 'num_threads', 'nice',
    'read_count', 'write_count', 'read_bytes', 'write_bytes',
    'cpu_times_user', 'cpu_times_system',
    'voluntary_ctx_switches', 'involuntary_ctx_switches',
    'memory_rss', 'memory_vms'
]

META_COLUMNS = ['pid', 'name']
ALL_COLUMNS = META_COLUMNS + FEATURE_COLUMNS


def collect_snapshot():
    rows = []

    for p in psutil.process_iter(['pid']):
        try:
            p.cpu_percent(None)
        except:
            continue

    time.sleep(0.3)

    for proc in psutil.process_iter([
        "pid", "name", "cpu_times", "memory_info", "memory_percent",
        "num_ctx_switches", "num_handles", "num_threads",
        "io_counters", "nice"
    ]):

        try:
            info = proc.info
            row = {}

            row['pid'] = info.get('pid', 0)
            row['name'] = info.get('name', '') or ''

            # CPU %
            try:
                row['cpu_percent'] = proc.cpu_percent(None)
            except:
                row['cpu_percent'] = 0.0

            row['memory_percent'] = info.get('memory_percent', 0.0)
            row['num_handles'] = info.get('num_handles', 0)
            row['num_threads'] = info.get('num_threads', 0)
            row['nice'] = info.get('nice', 0)

            # I/O counters
            io = info.get('io_counters', None)
            if io:
                row['read_count'] = io.read_count
                row['write_count'] = io.write_count
                row['read_bytes'] = io.read_bytes
                row['write_bytes'] = io.write_bytes
            else:
                row['read_count'] = row['write_count'] = 0
                row['read_bytes'] = row['write_bytes'] = 0

            # CPU times
            ct = info.get('cpu_times', None)
            if ct:
                row['cpu_times_user'] = ct.user
                row['cpu_times_system'] = ct.system
            else:
                row['cpu_times_user'] = 0.0
                row['cpu_times_system'] = 0.0

            # Context switches
            cs = info.get('num_ctx_switches', None)
            if cs:
                row['voluntary_ctx_switches'] = cs.voluntary
                row['involuntary_ctx_switches'] = cs.involuntary
            else:
                row['voluntary_ctx_switches'] = 0
                row['involuntary_ctx_switches'] = 0

            # Memory info (MB)
            mem = info.get('memory_info', None)
            if mem:
                row['memory_rss'] = mem.rss / (1024**2)
                row['memory_vms'] = mem.vms / (1024**2)
            else:
                row['memory_rss'] = row['memory_vms'] = 0.0

            rows.append(row)

        except:
            continue

    df = pd.DataFrame(rows)
    df = df.fillna(0)
    return df[ALL_COLUMNS]


def collect_dataset(duration_seconds=60, interval_seconds=5,
                    output_file="raw_behavior_data.csv"):

    print(f"[INFO] Collecting raw process data for {duration_seconds}s...")
    start = time.time()
    all_data = []

    while time.time() - start < duration_seconds:
        snap = collect_snapshot()
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
    collect_dataset(duration_seconds=300, interval_seconds=5)
