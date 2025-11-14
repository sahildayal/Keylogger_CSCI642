import joblib
import os
import sys

# path stuff to get it so we don't have to rewrite func
cur_dir = os.path.dirname(os.path.abspath(__file__))
phase3_dir = os.path.abspath(os.path.join(cur_dir, ".."))
sys.path.append(phase3_dir)
from utils.generate_data import snapshot

def ml_detector_with_confidence(proc_data):
    loaded_rf = joblib.load("./random_forest.joblib")
    expected_features = loaded_rf.feature_names_in_
    proc_data_clean = proc_data[expected_features].copy()
    probabilities = loaded_rf.predict_proba(proc_data_clean)
    
    print("\n[DETECTION RESULTS]\n")
    found = False
    for i, (prob_benign, prob_keylogger) in enumerate(probabilities):
        if prob_keylogger > 0.7:  #70% confidence threshold; TODO: does this need to be adjusted?
            print(f"Keylogger detected! (Confidence: {prob_keylogger:.1%})")
            print(f"   Process: {proc_data.iloc[i].name}, PID: {proc_data.iloc[i].pid}")
            found = True
        elif prob_keylogger > 0.4:  #suspicious but not certain
            print(f"Suspicious process: {proc_data.iloc[i]['name']} (PID: {proc_data.iloc[i]['pid']}) (Confidence: {prob_keylogger:.1%})")
            # print(f"Suspicious process: {proc_data.iloc[i].name} (Confidence: {prob_keylogger:.1%})")
    
    if not found:
        print("No keyloggers detected") 
    
    return found

if __name__ == "__main__":
    proc_data = snapshot()
    ml_detector_with_confidence(proc_data)
