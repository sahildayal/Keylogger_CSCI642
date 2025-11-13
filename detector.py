import psutil
import time
import pandas as pd
import joblib

def snapshot_processes():
    """Create live snapshot with the same features used during training."""
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
            row = {
                'pid': info.get('pid', 0),
                'name': info.get('name', '') or ''
            }

            try:
                row['cpu_percent'] = proc.cpu_percent(None)
            except:
                row['cpu_percent'] = 0.0

            row['memory_percent'] = info.get('memory_percent', 0.0)
            row['num_handles'] = info.get('num_handles', 0)
            row['num_threads'] = info.get('num_threads', 0)
            row['nice'] = info.get('nice', 0)

            # I/O
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

            # Memory in MB
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
    return df


def detector(model_path="retrained_detector.pkl"):
    bundle = joblib.load(model_path)
    model = bundle["model"]
    features = bundle["features"]

    df = snapshot_processes()
    proc_meta = df[['pid', 'name']]
    X = df[features].fillna(0)

    probs = model.predict_proba(X)
    classes = model.classes_
    key_index = list(classes).index('keylogger')

    print("\n[RESULTS]")
    detected = False

    for i, p in enumerate(probs):
        prob_key = p[key_index]
        name = proc_meta.iloc[i]['name']
        pid = proc_meta.iloc[i]['pid']

        if prob_key >= 0.80:
            print(f"[ALERT] Keylogger detected → {name} (PID {pid}) | {prob_key:.2f}")
            detected = True
        elif prob_key >= 0.50:
            print(f"[WARN ] Suspicious process → {name} (PID {pid}) | {prob_key:.2f}")

    if not detected:
        print("[INFO] No high-confidence keyloggers detected.")

    return detected


if __name__ == "__main__":
    detector()
