import psutil
import time
import pandas as pd
import joblib

def snapshot():
    rows = []

    for proc in psutil.process_iter([
        "pid", "name", "cpu_times", "memory_info", "memory_percent",
        "num_ctx_switches", "num_handles", "num_threads",
        "io_counters", "nice"
    ]):

        try:
            row = {}
            info = proc.info

            row['pid'] = info.get('pid', 0)
            row['name'] = info.get('name', '') or ''

            row['cpu_percent'] = proc.cpu_percent()
            time.sleep(0.01)

            row['memory_percent'] = info.get('memory_percent', 0.0)
            row['num_handles'] = info.get('num_handles', 0)
            row['num_threads'] = info.get('num_threads', 0)
            row['nice'] = info.get('nice', 0)

            io = info.get('io_counters', None)
            if io:
                row['read_count'] = io.read_count
                row['write_count'] = io.write_count
                row['read_bytes'] = io.read_bytes
                row['write_bytes'] = io.write_bytes
            else:
                row['read_count'] = row['write_count'] = 0
                row['read_bytes'] = row['write_bytes'] = 0

            ct = info.get('cpu_times', None)
            if ct:
                row['cpu_times_user'] = ct.user
                row['cpu_times_system'] = ct.system
            else:
                row['cpu_times_user'] = row['cpu_times_system'] = 0.0

            cs = info.get('num_ctx_switches', None)
            if cs:
                row['voluntary_ctx_switches'] = cs.voluntary
                row['involuntary_ctx_switches'] = cs.involuntary
            else:
                row['voluntary_ctx_switches'] = row['involuntary_ctx_switches'] = 0

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

    df = snapshot()
    meta = df[['pid','name']]
    X = df[features].fillna(0)

    probs = model.predict_proba(X)
    classes = model.classes_

    key_idx = list(classes).index('keylogger')

    print("\n[DETECTION RESULTS]\n")

    found = False

    for i, (prob) in enumerate(probs):
        p_key = prob[key_idx]
        pid = meta.iloc[i]['pid']
        name = meta.iloc[i]['name']

        if p_key >= 0.90:
            print(f"[ALERT] Keylogger detected → {name} (PID {pid}) | {p_key:.2f}")
            found = True
        elif p_key >= 0.60:
            print(f"[WARN ] Suspicious → {name} (PID {pid}) | {p_key:.2f}")

    if not found:
        print("[INFO] No keylogger confidently detected.")

    return found


if __name__ == "__main__":
    detector()
