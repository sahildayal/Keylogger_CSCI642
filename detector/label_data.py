import pandas as pd

def label_continuous_data():
    # Load your continuous data
    df = pd.read_csv('continuous_data_20251113_1404.csv')
    
    # Label Python processes as keyloggers (since you ran keyloggers during collection)
    df['label'] = df['name'].apply(
        lambda x: 'keylogger' if 'python' in str(x).lower() else 'benign'
    )
    
    # Check the labeling
    keylogger_count = (df['label'] == 'keylogger').sum()
    print(f"Labeled {keylogger_count} processes as keyloggers")
    print(f"Labeled {len(df) - keylogger_count} processes as benign")
    
    # Save labeled data
    df.to_csv('labeled_continuous_data.csv', index=False)
    print("✅ Saved as 'labeled_continuous_data.csv'")
    
    return df

if __name__ == "__main__":
    label_continuous_data()