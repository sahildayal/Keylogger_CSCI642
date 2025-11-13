import psutil
import joblib
import pandas as pd

def check_running_keyloggers():
    print("=" * 50)
    print("CHECK 1: RUNNING KEYLOGGERS")
    print("=" * 50)
    
    processes = [p for p in psutil.process_iter(['pid', 'name', 'cmdline']) if 'python' in p.info['name'].lower()]
    print(f'Python processes running: {len(processes)}')
    
    if len(processes) == 0:
        print("NO Python processes found - keyloggers aren't running!")
    else:
        print("Python processes found:")
        for p in processes:
            print(f'  PID: {p.info["pid"]}, Name: {p.info["name"]}')
            if p.info['cmdline']:
                cmd = " ".join(p.info['cmdline'])
                print(f'    Command: {cmd}')
                if 'keylog' in cmd.lower() or 'stealth' in cmd.lower():
                    print('    THIS IS A KEYLOGGER!')
    print()

def check_model_features():
    print("=" * 50)
    print("CHECK 2: MODEL FEATURES")
    print("=" * 50)
    
    try:
        model = joblib.load('retrained_detector.pkl')
        print(f'Model loaded successfully')
        print(f'Model features ({len(model.feature_names_in_)}): {list(model.feature_names_in_)}')
        print(f'Model classes: {model.classes_}')
    except Exception as e:
        print(f'Failed to load model: {e}')
        try:
            model = joblib.load('random_forest.joblib')
            print(f' Fallback model loaded')
            print(f'Model features ({len(model.feature_names_in_)}): {list(model.feature_names_in_)}')
        except:
            print('No model files found')
    print()

def check_labeled_data():
    print("=" * 50)
    print("CHECK 3: LABELED TRAINING DATA")
    print("=" * 50)
    
    try:
        df = pd.read_csv('labeled_continuous_data.csv')
        print(f'Labeled data loaded: {len(df)} samples')
        print(f'Keyloggers: {(df["label"] == "keylogger").sum()}')
        print(f'Benign: {(df["label"] == "benign").sum()}')
        print(f'Columns: {list(df.columns)}')
        
        if (df["label"] == "keylogger").sum() > 0:
            print('\nSample keylogger processes:')
            keyloggers = df[df['label'] == 'keylogger'][['name', 'cpu_percent', 'memory_rss']].head(3)
            print(keyloggers.to_string(index=False))
        else:
            print('NO keyloggers found in training data!')
            
    except Exception as e:
        print(f'Failed to load labeled data: {e}')
        print('Trying continuous_data_20251113_1404.csv...')
        try:
            df = pd.read_csv('continuous_data_20251113_1404.csv')
            print(f'Continuous data: {len(df)} samples')
            python_procs = df[df['name'].str.contains('python', case=False, na=False)]
            print(f'Python processes in continuous data: {len(python_procs)}')
        except:
            print('No data files found')
    print()

def check_detector_features():
    print("=" * 50)
    print("CHECK 4: DETECTOR FEATURE MATCH")
    print("=" * 50)
    
    # Simulate what the detector sees
    try:
        sample_proc = next(p for p in psutil.process_iter(['pid', 'name', 'cpu_percent', 'memory_info']) if p.info['pid'] != 0)
        print(f'Sample process: {sample_proc.info["name"]} (PID: {sample_proc.info["pid"]})')
        
        # Try to load model and check feature alignment
        model = joblib.load('retrained_detector.pkl')
        detector_features = model.feature_names_in_
        print(f'Model expects: {len(detector_features)} features')
        
        # Check if we can collect these features
        available_features = []
        for feature in detector_features:
            try:
                # Try to access each feature
                if hasattr(sample_proc.info, feature) or feature in sample_proc.info:
                    available_features.append(feature)
            except:
                pass
                
        print(f'Features available from processes: {len(available_features)}/{len(detector_features)}')
        if len(available_features) != len(detector_features):
            print('Feature mismatch detected!')
            missing = set(detector_features) - set(available_features)
            print(f'Missing features: {missing}')
        
    except Exception as e:
        print(f'Error checking features: {e}')
    print()

if __name__ == "__main__":
    print("RUNNING KEYLOGGER DETECTION DEBUG CHECKS")
    print()
    
    check_running_keyloggers()
    check_model_features()
    check_labeled_data()
    check_detector_features()
    
    print("=" * 50)
    print("DEBUG SUMMARY")
    print("=" * 50)