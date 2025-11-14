import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import joblib
import os

CSV_PATH = "balanced_dataset.csv"
MODEL_PATH = "retrained_detector.pkl"

# =========================================================
# Utility: Load Data + Create /graphs folder
# =========================================================
def load_data():
    if not os.path.exists(CSV_PATH):
        raise FileNotFoundError(f"{CSV_PATH} does not exist. Run label_dataset first.")
    
    df = pd.read_csv(CSV_PATH)
    os.makedirs("graphs", exist_ok=True)
    return df


# =========================================================
# 1. Label Distribution
# =========================================================
def plot_label_distribution(df):
    plt.figure(figsize=(6, 4))
    df["label"].value_counts().plot(
        kind="bar",
        color=["#4CAF50", "#F44336"]
    )
    plt.title("Label Distribution (Benign vs Keylogger)")
    plt.xlabel("Class")
    plt.ylabel("Count")
    plt.tight_layout()
    plt.savefig("graphs/label_distribution.png")
    plt.close()


# =========================================================
# 2. Boxplots for ALL numeric features
# =========================================================
def plot_feature_boxplots(df):
    numeric_cols = df.select_dtypes(include=["float64", "int64"]).columns

    for col in numeric_cols:
        if col in ["pid"]:  # skip pid
            continue

        plt.figure(figsize=(7, 4))
        sns.boxplot(x=df["label"], y=df[col])
        plt.title(f"Boxplot: {col} (Benign vs Keylogger)")
        plt.tight_layout()
        plt.savefig(f"graphs/{col}_boxplot.png")
        plt.close()


# =========================================================
# 3. Histograms for Top 5 Most Relevant Features
# =========================================================
TOP_FEATURES = [
    "cpu_percent",
    "read_bytes",
    "write_bytes",
    "read_count",
    "write_count"
]

def plot_feature_histograms(df):
    for col in TOP_FEATURES:
        if col not in df.columns:
            continue

        plt.figure(figsize=(7, 4))
        sns.histplot(df[df["label"]=="benign"][col], kde=True, color="blue", label="Benign", alpha=0.5)
        sns.histplot(df[df["label"]=="keylogger"][col], kde=True, color="red", label="Keylogger", alpha=0.5)
        plt.legend()
        plt.title(f"Histogram: {col}")
        plt.tight_layout()
        plt.savefig(f"graphs/{col}_hist.png")
        plt.close()


# =========================================================
# 4. Correlation Heatmap
# =========================================================
def plot_correlation_heatmap(df):
    numeric = df.select_dtypes(include=["float64", "int64"])
    plt.figure(figsize=(12, 10))
    sns.heatmap(numeric.corr(), cmap="coolwarm", annot=False)
    plt.title("Feature Correlation Heatmap")
    plt.tight_layout()
    plt.savefig("graphs/correlation_heatmap.png")
    plt.close()


# =========================================================
# 5. Feature Importance (Random Forest)
# =========================================================
def plot_feature_importance():
    if not os.path.exists(MODEL_PATH):
        print("Model not found. Train model first.")
        return
    
    bundle = joblib.load(MODEL_PATH)
    model = bundle["model"]
    feature_names = bundle["features"]

    importances = model.feature_importances_

    plt.figure(figsize=(10, 6))
    sns.barplot(x=importances, y=feature_names, palette="viridis")
    plt.title("Random Forest Feature Importance")
    plt.xlabel("Importance")
    plt.ylabel("Feature")
    plt.tight_layout()
    plt.savefig("graphs/feature_importance.png")
    plt.close()


# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":
    df = load_data()
    
    print("[INFO] Generating graphs in /graphs folder...")

    plot_label_distribution(df)
    plot_feature_boxplots(df)
    plot_feature_histograms(df)
    plot_correlation_heatmap(df)
    plot_feature_importance()

    print("[INFO] Done! All graphs saved in 'graphs/' folder.")
