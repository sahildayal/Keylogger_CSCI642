import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, classification_report
import joblib

def retrain_with_continuous_data():
    print("RETRAINING MODEL WITH CONTINUOUS DATA")
    
    labeled_df = pd.read_csv('labeled_continuous_data.csv')
    print(f"Loaded labeled continuous data: {len(labeled_df)} samples")

    # Use this as our main dataset
    combined_df = labeled_df
    
    # Label Python processes as keyloggers in the continuous data
    # This is a simplification - in real scenario you'd have proper labels
    combined_df['label'] = combined_df['label'].fillna('benign')  # Fill any missing labels
    
    # If no labels in continuous data, create them based on process names
    if 'label' not in combined_df.columns or combined_df['label'].isna().all():
        combined_df['label'] = combined_df['name'].apply(
            lambda x: 'keylogger' if 'python' in str(x).lower() else 'benign'
        )
    
    print(f"Final dataset: {len(combined_df)} total samples")
    print(f"Keyloggers: {combined_df['label'].value_counts().get('keylogger', 0)}")
    print(f"Benign: {combined_df['label'].value_counts().get('benign', 0)}")
    
    # Prepare features - use only numeric columns
    feature_columns = [
        'cpu_percent', 'memory_percent', 'num_handles', 'num_threads', 'nice',
        'read_count', 'write_count', 'read_bytes', 'write_bytes', 
        'cpu_times_user', 'cpu_times_system', 'voluntary_ctx_switches',
        'involuntary_ctx_switches', 'memory_rss', 'memory_vms'
    ]
    
    # Select only existing numeric columns
    X_columns = [col for col in feature_columns if col in combined_df.columns]
    X = combined_df[X_columns].fillna(0)
    y = combined_df['label']
    
    print(f"Using features: {X_columns}")
    
    # Train new model
    rf_model = RandomForestClassifier(
        n_estimators=100,
        max_depth=10,
        class_weight='balanced',
        random_state=42
    )
    
    rf_model.fit(X, y)
    
    # Save the new model
    joblib.dump(rf_model, 'retrained_detector.pkl')
    print("New model saved as 'retrained_detector.pkl'")
    
    # Test accuracy
    accuracy = rf_model.score(X, y)
    print(f"Training accuracy: {accuracy:.3f}")
    
    # Show feature importance
    print("\nFeature Importances:")
    for feature, importance in zip(X_columns, rf_model.feature_importances_):
        print(f"  {feature}: {importance:.4f}")
    
    return rf_model, X_columns

if __name__ == "__main__":
    retrain_with_continuous_data()