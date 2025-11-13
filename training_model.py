import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report, accuracy_score
import joblib

FEATURE_COLUMNS = [
    'cpu_percent', 'memory_percent', 'num_handles', 'num_threads', 'nice',
    'read_count', 'write_count', 'read_bytes', 'write_bytes',
    'cpu_times_user', 'cpu_times_system',
    'voluntary_ctx_switches', 'involuntary_ctx_switches',
    'memory_rss', 'memory_vms'
]


def train_detector(csv_path="labeled_behavior_data.csv",
                   model_path="retrained_detector.pkl"):

    df = pd.read_csv(csv_path)
    df = df[df['label'].isin(['benign','keylogger'])]

    X = df[FEATURE_COLUMNS].fillna(0)
    y = df['label']

    print("[INFO] Label distribution:")
    print(y.value_counts())

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.25, random_state=42, stratify=y
    )

    model = RandomForestClassifier(
        n_estimators=200,
        max_depth=16,
        class_weight='balanced',
        random_state=42
    )

    model.fit(X_train, y_train)

    preds = model.predict(X_test)

    print("[INFO] Accuracy:", accuracy_score(y_test, preds))
    print("\n[INFO] Classification Report:\n", classification_report(y_test, preds))

    bundle = {"model": model, "features": FEATURE_COLUMNS}
    joblib.dump(bundle, model_path)

    print(f"[INFO] Saved detector → {model_path}")


if __name__ == "__main__":
    train_detector()
