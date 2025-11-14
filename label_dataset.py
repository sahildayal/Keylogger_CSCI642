import pandas as pd
import psutil

def detect_possible_keylogger_pids():
    """Automatically scan running processes for possible keylogger PIDs."""
    candidates = []

    for proc in psutil.process_iter(["pid", "name", "exe", "cmdline"]):
        try:
            name = (proc.info["name"] or "").lower()
            cmd = " ".join(proc.info.get("cmdline") or []).lower()

            # Direct script names
            if "key_logger.py" in cmd or "stealth_keylogger.py" in cmd:
                candidates.append(proc.info["pid"])

            # Python-based keylogger heuristics
            elif "python" in name and ("key" in cmd or "stealth" in cmd):
                candidates.append(proc.info["pid"])

        except:
            continue

    return list(set(candidates))


def label_dataset(input_csv="raw_behavior_data.csv",
                  output_csv="labeled_behavior_data.csv"):
    """Label ONLY the actual keylogger PID as 'keylogger' and all others as 'benign'."""

    df = pd.read_csv(input_csv)
    print(f"[INFO] Loaded {len(df)} total rows from {input_csv}")

    print("\n[INFO] Attempting auto-detection of keylogger PIDs...")
    auto_pids = detect_possible_keylogger_pids()
    print(f"[INFO] Auto-detected possible keylogger PIDs: {auto_pids}")

    # Ask user which one is the REAL keylogger PID
    user_pid = input("Enter CONFIRMED keylogger PID (or press Enter to use auto-detected ones): ").strip()

    if user_pid:
        keylogger_pids = [int(user_pid)]
    else:
        keylogger_pids = auto_pids

    if not keylogger_pids:
        print("[WARN] No keylogger PID provided or detected. Dataset will NOT be labeled.")
        return

    print(f"[INFO] Labeling PID(s) {keylogger_pids} as keylogger...")

    df['label'] = df['pid'].apply(lambda p: "keylogger" if p in keylogger_pids else "benign")

    df.to_csv(output_csv, index=False)
    print(f"[INFO] Labeled dataset saved → {output_csv}")


# -------------------------------
#            MAIN
# -------------------------------
if __name__ == "__main__":
    print("\n=== Keylogger Labeling Script (Hybrid Mode) ===\n")
    inp = input("Enter raw CSV filename (default: raw_behavior_data.csv): ").strip()

    if inp == "":
        inp = "raw_behavior_data.csv"

    out = input("Enter output CSV filename (default: labeled_behavior_data.csv): ").strip()

    if out == "":
        out = "labeled_behavior_data.csv"

    label_dataset(input_csv=inp, output_csv=out)