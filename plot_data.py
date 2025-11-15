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
    numeric_cols = df.select_dtypes(include=["int64", "float64"]).columns

    for col in numeric_cols:
        if col in ["pid"]:
            continue

        plt.figure(figsize=(8, 5))
        sns.boxplot(
            data=df,
            x="label",
            y=col,
            palette={"benign": "#4CAF50", "keylogger": "#F44336"},
            width=0.5,
            linewidth=1.2,
            showfliers=True
        )

        # APPLY LOG SCALE (IMPORTANT)
        plt.yscale("log")

        plt.title(f"{col} — Box & Whisker (Log Scale)")
        plt.xlabel("Class")
        plt.ylabel(f"{col} (log scale)")
        plt.tight_layout()
        plt.savefig(f"graphs/{col}_boxplot_log.png", dpi=300)
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
# 6. Combined Multi-Feature Pairplot (Best 3 Features)
# =========================================================
def plot_combined_pairplot(df):
    features = ["read_bytes", "write_bytes", "num_handles"]

    # Check all features exist
    for f in features:
        if f not in df.columns:
            print(f"[WARN] Skipping pairplot. Missing feature: {f}")
            return

    # Map your label colors
    palette = {
        "benign": "#4CAF50",      # green
        "keylogger": "#F44336"    # red
    }

    # Create folder
    os.makedirs("graphs", exist_ok=True)

    print("[INFO] Creating combined pairplot for read_bytes, write_bytes, num_handles...")

    sns.pairplot(
        df,
        vars=features,
        hue="label",
        palette=palette,
        diag_kind="kde",
        corner=True
    )

    plt.suptitle("Combined Feature Pairplot: read_bytes, write_bytes, num_handles", y=1.02)
    plt.tight_layout()
    plt.savefig("graphs/combined_pairplot.png", dpi=300)
    plt.close()

    print("[INFO] Saved: graphs/combined_pairplot.png")


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
    plot_combined_pairplot(df)


    print("[INFO] Done! All graphs saved in 'graphs/' folder.")
