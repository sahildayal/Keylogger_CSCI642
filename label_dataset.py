import pandas as pd
import psutil

def detect_possible_keylogger_pids():
    candidates = []

    for proc in psutil.process_iter(["pid", "name", "exe", "cmdline"]):
        try:
            name = (proc.info["name"] or "").lower()
            cmd = " ".join(proc.info.get("cmdline") or []).lower()

            if "key_logger" in cmd or "stealth_keylogger" in cmd:
                candidates.append(proc.info["pid"])
            elif "python" in name and "key" in cmd:
                candidates.append(proc.info["pid"])
        except:
            continue

    return list(set(candidates))


def label_dataset(input_csv="raw_behavior_data.csv",
                  output_csv="labeled_behavior_data.csv"):

    df = pd.read_csv(input_csv)
    print(f"[INFO] Loaded {len(df)} rows.")

    print("\n[INFO] Auto-detecting keylogger PIDs from running processes...")
    auto_pids = detect_possible_keylogger_pids()

    print(f"Auto-detected possible keylogger PIDs: {auto_pids}")

    user_pid = input("Enter CONFIRMED keylogger PID (or leave empty to use auto-detect): ").strip()

    if user_pid:
        keylogger_pids = [int(user_pid)]
    else:
        keylogger_pids = auto_pids

    if not keylogger_pids:
        print("[WARN] No keylogger PID detected. Labeling skipped.")
        return

    print(f"[INFO] Using keylogger PIDs: {keylogger_pids}")

    df['label'] = df['pid'].apply(lambda p: "keylogger" if p in keylogger_pids else "benign")

    df.to_csv(output_csv, index=False)
    print(f"[INFO] Saved labeled dataset → {output_csv}")
