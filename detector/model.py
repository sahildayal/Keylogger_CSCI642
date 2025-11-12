# Import necessary libraries
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score
from sklearn.datasets import load_iris  # For example, using the Iris dataset
import joblib

# Load the dataset (for demonstration, we use the Iris dataset)
benign_data = pd.read_csv('data/benign_data.csv')
keylogger_data = pd.read_csv('data/keylogger_data.csv')
test_data = pd.concat([benign_data.iloc[50:], keylogger_data.iloc[50:]], ignore_index=True)
data = pd.concat([benign_data.iloc[:50], keylogger_data.iloc[:50]], ignore_index=True)
X = data.drop(['label', 'pid', 'name'], axis=1)  # Features
y = data['label']  # Labels

# Split the dataset into training and testing sets
# X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0., random_state=42)

# Initialize the Random Forest model
rf_model = RandomForestClassifier(n_estimators=100, random_state=42)

# Train the model
rf_model.fit(X, y)

# Make predictions on the test set
y_pred = rf_model.predict(test_data.drop(['label', 'pid', 'name'], axis=1))

# Evaluate the model using accuracy
accuracy = accuracy_score(test_data['label'], y_pred)

# Print the results
print(f"Random Forest Model Accuracy: {accuracy * 100:.2f}%")

# Optionally, you can also check feature importance:
# print("features:", rf_model.feature_names_in_)
# print("Feature importances:", rf_model.feature_importances_)
for feature, importance in zip(rf_model.feature_names_in_, rf_model.feature_importances_):
    print(f"Feature: {feature}, Importance: {importance}")

joblib.dump(rf_model, "./random_forest.joblib")