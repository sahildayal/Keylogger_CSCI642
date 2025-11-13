import psutil
import time
import pandas as pd
import os
from datetime import datetime

# Canonical feature set used across the entire project
FEATURE_COLUMNS = [
    'cpu_percent', 'memory_percent', 'num_handles', 'num_threads', 'nice',
    'read_count', 'write_count', 'read_bytes', 'write_bytes',
    'cpu_times_user', 'cpu_times_system',
    'voluntary_ctx_switches', 'involuntary_ctx_switches',
    'memory_rss', 'memory_vms'
]

META_COLUMNS = ['pid', 'name', 'label']
ALL_COLUMNS = ['pid', 'name'] + FEATURE_COLUMNS + ['label']


def collect_snapshot(label='unknown'):
    """Collect a single system snapshot with canonical features."""
    rows = []

    # Prime cpu_percent()
    for proc in psutil.process_iter(['pid']):
        try:
            proc.cpu_percent(None)
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

            # Basic metadata
            row['pid'] = info.get('pid', 0)
            row['name'] = info.get('name', '') or ''

            # CPU %
            try:
                row['cpu_percent'] = proc.cpu_percent(None)
            except:
                row['cpu_percent'] = 0.0

            # Memory %
            row['memory_percent'] = info.get('memory_percent', 0.0)

            # Threads / handles / nice
            row['num_handles'] = info.get('num_handles', 0) or 0
            row['num_threads'] = info.get('num_threads', 0) or 0
            row['nice'] = info.get('nice', 0) or 0

            # I/O counters
            io = info.get('io_counters', None)
            if io:
                row['read_count'] = getattr(io, 'read_count', 0)
                row['write_count'] = getattr(io, 'write_count', 0)
                row['read_bytes'] = getattr(io, 'read_bytes', 0)
                row['write_bytes'] = getattr(io, 'write_bytes', 0)
            else:
                row['read_count'] = row['write_count'] = 0
                row['read_bytes'] = row['write_bytes'] = 0

            # CPU times
            ct = info.get('cpu_times', None)
            if ct:
                row['cpu_times_user'] = getattr(ct, 'user', 0.0)
                row['cpu_times_system'] = getattr(ct, 'system', 0.0)
            else:
                row['cpu_times_user'] = 0.0
                row['cpu_times_system'] = 0.0

            # Context switches
            cs = info.get('num_ctx_switches', None)
            if cs:
                row['voluntary_ctx_switches'] = getattr(cs, 'voluntary', 0)
                row['involuntary_ctx_switches'] = getattr(cs, 'involuntary', 0)
            else:
                row['voluntary_ctx_switches'] = 0
                row['involuntary_ctx_switches'] = 0

            # Memory info → MB
            mem = info.get('memory_info', None)
            if mem:
                row['memory_rss'] = getattr(mem, 'rss', 0) / (1024**2)
                row['memory_vms'] = getattr(mem, 'vms', 0) / (1024**2)
            else:
                row['memory_rss'] = row['memory_vms'] = 0.0

            # Label assigned externally
            row['label'] = label

            rows.append(row)

        except:
            continue

    df = pd.DataFrame(rows)
    df = df.fillna(0)
    return df[ALL_COLUMNS]


def collect_dataset(label, duration_seconds=60, interval_seconds=5, output_file="behavior_dataset.csv"):
    """Continuously collect snapshots and append them to one dataset."""
    print(f"[INFO] Collecting {label} data for {duration_seconds}s...")

    start = time.time()
    all_data = []

    while time.time() - start < duration_seconds:
        snap = collect_snapshot(label)
        print(f"[INFO] Snapshot at {datetime.now().strftime('%H:%M:%S')} with {len(snap)} processes")
        all_data.append(snap)
        time.sleep(interval_seconds)

    final_df = pd.concat(all_data, ignore_index=True)

    # Append or create dataset
    if os.path.exists(output_file):
        final_df.to_csv(output_file, mode='a', header=False, index=False)
    else:
        final_df.to_csv(output_file, index=False)

    print(f"[INFO] Saved {len(final_df)} rows to {output_file}")


if __name__ == "__main__":
    # Example: during benign run
    collect_dataset(label="benign", duration_seconds=60, interval_seconds=5)
