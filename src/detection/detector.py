import joblib
import os
import sys
import psutil

# path stuff to get it so we don't have to rewrite func
cur_dir = os.path.dirname(os.path.abspath(__file__))
phase_dir = os.path.abspath(os.path.join(cur_dir, ".."))
sys.path.append(phase_dir)
from utils.generate_data import snapshot

def ml_detector_with_confidence(snapshot, model_path = "retrained_detector.pkl"):
    """
    Tests a the snapshot of the running processes against the trainted Random Forest classifier.
    Prints processes that are detected as keyloggers or suspected of being keyloggers along with
    the probability. Returns a list of process ids that are detected or suspected. 
    
    :param snapshot: dataframe containing information about all running processes on the system
    :param model_path: the path to the trained model file
    :return: list of detected keylogger process ids, list of suspected keylogger process ids
    """

    bundle = joblib.load(model_path)
    model = bundle["model"]
    features = bundle["features"]
    
    meta = snapshot[['pid', 'name']]
    X = snapshot[features].fillna(0)

    probs = model.predict_proba(X)
    classes = model.classes_

    key_idx = list(classes).index('keylogger')
    
    print("\n[DETECTION RESULTS]\n")
    found = False
    keylogger_pids = []
    sus_keylogger_pids = []
    for i, (prob) in enumerate(probs):
        p_key = prob[key_idx]
        pid = meta.iloc[i]['pid']
        name = meta.iloc[i]['name']

        if p_key >= 0.70:
            print(f"[ALERT] Keylogger detected → {name} (PID {pid}) | {p_key:.2f}")
            found = True
            keylogger_pids.append(pid)
        elif p_key >= 0.40:
            print(f"[WARN] Suspicious → {name} (PID {pid}) | {p_key:.2f}")
            sus_keylogger_pids.append(pid)

    if not found:
        print("[INFO] No keylogger confidently detected.")

    return keylogger_pids, sus_keylogger_pids 

def terminate_keyloggers(pids, severity):
    """
    Asks the user if they would like to terminate each of the detected or suspected keylogger processes

    :param pids: process ids that are being considered for termination
    :param severity: whether the process ids have the level of detected or suspected
    """

    for p in pids:
        terminate = input(f"Would you like to terminate this [{severity}] keylogger: PID {p}? (Y/N) ")
        if terminate == "Y" or terminate == "y":
            try:
                process = psutil.Process(p)
                process.terminate()
                print(f"[INFO] Process {p} terminated successfully")
            except psutil.NoSuchProcess:
                print(f"[INFO] Process {p} not found")
            except psutil.AccessDenied:
                print(f"[INFO] Authorized to terminate process {p}")
            except Exception as e:
                print(f"[INFO] Process {p} could not be terminated: {e}")
        else:
            print(f"[INFO] Process {p} was not terminated")

def run_detector():
    """
    Runs the detection program. First gets a snapshot of current system processes, classifies 
    those against trained model, and asks the user if they would like to terminate the detected
    and suspected keylogger processes
    """

    proc_data = snapshot()
    found_keyloggers, sus_keyloggers = ml_detector_with_confidence(proc_data)
    if len(found_keyloggers) > 0:
        terminate_keyloggers(found_keyloggers, "detected")
    if len(sus_keyloggers) > 0:
        terminate_keyloggers(sus_keyloggers, "suspected")

if __name__ == "__main__":
    run_detector()
