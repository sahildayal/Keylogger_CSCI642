import pandas as pd

INPUT_FILE = "combined_dataset.csv"
OUTPUT_FILE = "balanced_dataset.csv"

# Set ratio of benign:keylogger
# Example: if keylogger=200 rows, benign=600 rows (3:1 ratio)
BENIGN_RATIO = 15     # change this if you want 1:1 or 2:1

print("[INFO] Loading combined dataset...")
df = pd.read_csv(INPUT_FILE)

# Split classes
benign_df = df[df["label"] == "benign"]
keylogger_df = df[df["label"] == "keylogger"]

print(f"[INFO] Benign count: {len(benign_df)}")
print(f"[INFO] Keylogger count: {len(keylogger_df)}")

# Number of benign rows to sample
target_benign = min(len(benign_df), len(keylogger_df) * BENIGN_RATIO)

print(f"[INFO] Sampling {target_benign} benign rows...")

benign_sample = benign_df.sample(n=target_benign, random_state=42)

# Combine everything
balanced_df = pd.concat([benign_sample, keylogger_df], ignore_index=True)

# Shuffle
balanced_df = balanced_df.sample(frac=1, random_state=42).reset_index(drop=True)

# Save
balanced_df.to_csv(OUTPUT_FILE, index=False)

print("[INFO] Balanced dataset saved →", OUTPUT_FILE)
print("[INFO] Final label counts:")
print(balanced_df["label"].value_counts())
