import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, classification_report
import joblib

# Data that is collected in the csv files about system processes have these columns
FEATURE_COLUMNS = [
    'cpu_percent', 'memory_percent', 'num_handles', 'num_threads', 'nice',
    'read_count', 'write_count', 'read_bytes', 'write_bytes',
    'cpu_times_user', 'cpu_times_system',
    'voluntary_ctx_switches', 'involuntary_ctx_switches',
    'memory_rss', 'memory_vms'
]

def train_model(csv_path = "../../data/merged_data/combined_dataset.csv", 
                joblib_path = "./random_forest.joblib", 
                pkl_path = './random_forest.pkl'):
    """
    Trains a Random Forest Classifier model based on the dataset located at csv_path. 
    This model is then saved to the joblib_path and pkl_path in .joblib and .pkl formats respectively. 

    Provides information about the Random Forest model including:
        - model accuracy on test data
        - classification report
        - feature importance

    :param csv_path: file location to the training dataset
    :param joblib_path: file location where joblib file will be saved
    :param pkl_path: file location where pkl file will be saved
    """
    
    # Load the dataset 
    print("[INFO] Loading dataset...")
    data = pd.read_csv(csv_path)

    # shuffle the frames 
    data = data.sample(frac=1, random_state=42).reset_index(drop=True)
    
    data = data[data['label'].isin(['benign', 'keylogger'])]

    # already cleaned exe, cmdline, open_files
    # drop_cols = ['label', 'pid', 'name', 'exe', 'cmdline', 'open_files']
    # drop_cols = ['label', 'pid', 'name']
    # X = data.drop(drop_cols, axis=1)

    X = data[FEATURE_COLUMNS].fillna(0)
    y = data['label']  # Labels
    # y_numerical = y.map({'benign': 0, 'keylogger': 1})
    print("[INFO] Label Distribution:")
    print(y.value_counts())

    # Split the dataset into training and testing sets
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42, stratify=y)

    # Initialize the Random Forest model; n_jobs=-1 means all cpu's work on it (quickest)
    rf_model = RandomForestClassifier(
        n_estimators=20, 
        class_weight='balanced',
        max_depth=15, random_state=42,
        n_jobs=-1
    )

    # Train the model
    rf_model.fit(X_train, y_train)

    # Make predictions on the test set
    y_pred = rf_model.predict(X_test)

    # Evaluate the model using accuracy
    accuracy = accuracy_score(y_test, y_pred)

    # Print the results
    print(f"[INFO] Random Forest Model Accuracy: {accuracy * 100:.2f}%")
    print("\n[INFO] Classification Report:\n", classification_report(y_test, y_pred))

    bundle = {"model": rf_model, "features": FEATURE_COLUMNS}

    # Optionally, you can also check feature importance:
    for feature, importance in zip(rf_model.feature_names_in_, rf_model.feature_importances_):
        print(f"Feature: {feature}, Importance: {importance}")

    joblib.dump(rf_model, joblib_path)
    joblib.dump(bundle, pkl_path)
    print(f"[INFO] Saved detector model to {joblib_path} and {pkl_path}")

if __name__ == "__main__":
    train_model()