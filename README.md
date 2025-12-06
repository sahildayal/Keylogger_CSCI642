# Group3_CSCI642_Project: Improving User Space Keylogger Detection Through Behavioral Analysis

These are the parts to this project:
1. Creating a Python user space keylogger that runs on a Windows 10 system
2. Creating a detection system for the keylogger developed (and similar keyloggers already developed)

## Install and Requirements
It is recommented to not only use a Virtual Machine (VM), running Windows 10, but also to use a virtual environment in order to keep track of all packages. This can be done by doing the following:

```
python -m venv <name of virtual environment>
.\<name of virtual environment>\Scripts\activate
```

Then install the required packages using `pip` and `requirements.txt`:
```pip install -r requirements.txt```

## System Architecture
This displays the following important directories and files for this project. Some files have been excluded for abstraction and clarity purposes (ex. there are many data sample files that are not included).
```
Keylogger_CSCI642/
├── data
│   ├── dev1_collected_data
│   ├── merged_data
│   │   ├── balanced_dataset.csv        # preprocessed dataset used for training
│   │   └── combined_dataset.csv
│   ├── dev2_collected_data
│   └── dev3_collected_data
├── securecodingenv
├── src
│   ├── detection
│   │   ├── __init__.py
│   │   ├── detector.py                 # behavioral detection of live system metrics
│   │   ├── model.py                    # trains Random Forest model
│   │   ├── random_forest.joblib
│   │   ├── random_forest.pkl
│   │   └── retrained_detector.pkl      # trained classifier used for report and analysis
│   ├── exe_files
│   │   ├── build
│   │   │   ├── generate_data
│   │   │   ├── keylogger
│   │   │   └── stealth_keylogger
│   │   ├── dist
│   │   │   ├── generate_data.exe
│   │   │   ├── keylogger.exe
│   │   │   └── stealth_keylogger.exe
│   │   ├── generate_data.spec
│   │   ├── keylogger.spec
│   │   └── stealth_keylogger.spec
│   ├── graphs
│   ├── keylogger
│   │   ├── keylogger.py
│   │   └── stealth_keylogger.py
│   ├── utils
│   │   └── generate_data.py
├── .gitignore
├── CONTRIBUTIONS
├── README.md
├── README.txt
├── requirements.txt
```
## Pipeline Overview
The following describes how the project was developed and may be useful for recreating the results
- Run either the ```keylogger.py``` or ```stealth_keylogger.py```
- Generate keylogger data by utilizing ```generate_data.py```
- Conglomerate and clean all generated data (```clean_sort_merge.py``` may be of some use)
- Utilize that generated dataset to train the model using ```model.py```
- Run real-time detection wtih ```detector.py```

## Keylogger Development and Deployment
There are 2 ways to run the keylogger:
1. In the keylogger directory, run the command ```python .\keylogger.py```
    - This will log all keystrokes to a file called `keystroke_log.csv` in the directory in which the file is run (i.e. keylogger)
    - To stop the keylogger, simply press the ESC key

2. A `.exe` file has been provided in the exe_files directory. This was generated with the command ```pyinstaller --onefile .\keylogger.py```. This would be more similar to a real world attack, as not every computer would have Python installed (which is needed for the previous way to run the keylogger) and the `keylogger.exe` can be run in the command prompt/powershell of Window 10 computers

### Stealth Keylogger
Similar steps can be taken to run the stealth keylogger by replacing all instances of  ```keylogger``` in the above steps with ```stealth_keylogger```

## Generating Training Data
The program `generate_data.py` (located in utils directory) collects information on all running processes on the device. This data is used to train the keylogger detector model. It labels processes as either 'benign' or 'keylogger' (this is based on user-inputted process id values (pid))

Information Collected:
- CPU usage
- memory
- threads
- I/O operations
- open files
- context switches
- command line arguments

There are 2 different ways to run this program. The first is just a single snapshot and the second is continuous collection over a certain duration in an interval. The following describes how to run each section:

1. Snapshot
- `python .\generated_data.py 1`: Takes a snapshot of all current running processes on the machine without labeling the data. This data is stored in `snapshot.csv` in the data folder
- `python .\generated_data.py 2 "[<pid>, <pid>, ...]"`: Takes a snapshot of all current running processes on the machine. The pids listed will be labeled as 'keylogger' and all other processes as 'benign'.This data is stored in `snapshot.csv` in the data folder

2. Continuous
- `python .\generated_data.py 3`: Takes continuous snapshots of all current running processes on the machine without labeling. This defaults to 1 minute duration with 5 second intervals. This data is stored in `raw_behavior_data.csv` in the data folder
- `python .\generated_data.py 3 <duration> <interval>`: Takes continuous snapshots of all current running processes on the machine without labeling using the provided duration and interval (which are in seconds). This data is stored in `raw_behavior_data.csv` in the data folder


In order to run the program, follow the steps below:
1. Determine the pid of the thread running the keylogger (this can be done using Task Manager)
2. Use the pid(s) as a list argument when running the following command in the terminal where the `generate_data.py` file is located (phase3 directory)
- ```python .\generate_data.py "[<pid>, <pid>, ...]"```
- if no pid(s) are given, then it will label all processes as 'benign'

## Detector Development and Deployment
### Training the AI Model
A Random Forest model was developed and trained off of data from multiple Virtual Machines running Windows 10. This takes a snapshot of the current machine and runs the data against the trained model and outputs (with confidence) if there is a suspected keylogger or detected keylogger.
- Suspected Keylogger: Confidence of 40% - 70%
- Detected Keylogger: Confidence of >70%

To train the Random Forest model, run the following (from the detection directory):
```python .\model.py```
This loads the data and outputs random_forest.joblib and random_forest.pkl. When testing, the data loaded into the program was ```balanced_dataset.csv```, which trained a model to ≈0.86 accuracy and outputted the ```retrained_detector.pkl``` file.

### Running the Detector
Utilizing the Random Forest model trained, ```detector.py``` collects a live snapshot of all running processes and tests agains the model to classify process activity as:
- Detected Keylogger
- Suspected Keylogger
- Benign Activity
The user will then be able to go through the list of detected and suspected keylogger processes and decide whether or not to terminate them. 

The following shows how to run the detector (from the detection directory):
```python .\detector.py```

## Files in Utils
The files located in the utils directory (besides ```generate_data.py```) were used as helper scripts at certain points in the project and may or may not be useful when trying to get similar results. They will need to be edited to fit your specific needs. 

## Other Keyloggers Used for Data Training and Testing
These keyloggers were picked to help train the model and test if they could be detected as they are similar to the keylogger that was developed in this project. It's important that this tool does not only detect our keylogger, but other ones as well, as this tool's intended use is to prevent user space keyloggers on Windows 10 (not just the one developed)
- https://github.com/elxecutor/keylogger/tree/main
- https://github.com/ramprasathmk/keylogger
- https://github.com/creekmar/python-keylogger 

## Legal and Ethical Disclaimer
This keylogger was designed for **educational** and **research** purposes **only**!
