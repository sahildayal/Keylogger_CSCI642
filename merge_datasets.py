import pandas as pd
import os

# List all datasets you want to merge
datasets = [
    "labeled_behavior_data.csv",      # your own dataset
    "benign_data.csv",         # rename her files to avoid confusion
    "keylogger_data.csv"
]

# OPTIONAL (your raw_behavior_data should NOT be added unless it's labeled)
# datasets.append("raw_behavior_data.csv")   # only if it contains labels

# The canonical column order for your project
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
combined.to_csv("combined_dataset.csv", index=False)

print("[INFO] Combined dataset saved as combined_dataset.csv")
print("[INFO] Final shape:", combined.shape)
print(combined['label'].value_counts())
