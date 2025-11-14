import os
import pandas as pd

def extract_keylogger_data():
    key_df  = pd.DataFrame()
    benign_df  = pd.DataFrame()
    with os.scandir("../data/data_vm2") as es:
        for e in es:
            if e.is_file() and e.name.endswith('.csv'):
                df = pd.read_csv(e.path)

                # get keylogger data
                keylogger_data = df[df['label'] == 'keylogger']
                if key_df.empty:
                    key_df = keylogger_data
                else:
                    key_df = pd.concat([key_df, keylogger_data], ignore_index=True)

                # get benign data
                benign_data = df[df['label'] == 'benign']
                if benign_df.empty:
                    benign_df = benign_data
                else:
                    benign_df = pd.concat([benign_df, benign_data], ignore_index=True)
    key_df.to_csv('keylogger_data_vm2.csv', index=False)
    benign_df.to_csv('benign_data_vm2.csv', index=False)


def clean_keylogger_data():
    # clean data sets to remove sensitive information that was gathered

    datasets = ["../../data/sorted_data/benign_data.csv", "../../data/sorted_data/benign_data_vm2.csv", "../../data/sorted_data/keylogger_data.csv", "../../data/sorted_data/keylogger_data_vm2.csv", 
                "../../data/exe_data/vm1_5instances.csv", "../../data/exe_data/vm1_5instances_2.csv", "../../data/exe_data/vm1_5instances_3.csv", "../../data/exe_data/vm1_5instances_4.csv", "../../data/exe_data/vm1_5instances_5.csv", "../../data/exe_data/vm1_5instances_6.csv", "../../data/exe_data/vm1_5instances_7.csv", 
                "../../data/exe_data/vm1_5instances_8.csv", "../../data/exe_data/vm1_5instances_9.csv", "../../data/exe_data/vm1_5instances_10.csv", "../../data/exe_data/vm1_5instances_11.csv", "../../data/exe_data/vm1_5instances_12.csv", "../../data/exe_data/vm1_5instances_13.csv", "../../data/exe_data/vm1_5instances_14.csv", 
                "../../data/exe_data/vm1_5instances_15.csv", "../../data/exe_data/vm1_5instances_16.csv", "../../data/exe_data/vm1_5instances_17.csv", "../../data/exe_data/vm1_5instances_18.csv", "../../data/exe_data/vm1_5instances_19.csv", "../../data/exe_data/vm1_5instances_20.csv"]
    drop_cols = ["cmdline", "open_files", "exe", ]
    df_list = []

    for data in datasets:
        df = pd.read_csv(data)
        df = df.drop(columns=drop_cols)
        df_list.append(df)
    
    combined = pd.concat(df_list, ignore_index=True)
    combined.to_csv("../../data/sorted_data/cleaned_data.csv", index=False)
        
def merge_datasets():
    datasets = ["../../data/sorted_data/cleaned_data.csv", "../../data/ming/benign_data.csv", "../../data/ming/keylogger_data.csv", "../../data/sahil/labeled_behavior_data.csv"]
    canonical_columns = [
        "pid", "name", "cpu_percent", "memory_percent",
        "num_handles", "num_threads", "nice",
        "read_count", "write_count", "read_bytes", "write_bytes",
        "cpu_times_user", "cpu_times_system",
        "voluntary_ctx_switches", "involuntary_ctx_switches",
        "memory_rss", "memory_vms",
        "label"
        ]

    dfs = []

    for file in datasets:
        if not os.path.exists(file):
            print(f"[WARN] Skipping missing file: {file}")
            continue

        print(f"[INFO] Loading {file}")
        df = pd.read_csv(file)

        # Fix column order so everything aligns
        df = df[canonical_columns]

        dfs.append(df)

    # Combine everything
    combined = pd.concat(dfs, ignore_index=True)

    # Remove duplicates if any
    combined = combined.drop_duplicates()

    # Shuffle rows for better ML training
    combined = combined.sample(frac=1, random_state=42).reset_index(drop=True)

    # Save final dataset
    combined.to_csv("../../data/merged_data/combined_dataset.csv", index=False)

# extract_keylogger_data()
# clean_keylogger_data()
merge_datasets()
