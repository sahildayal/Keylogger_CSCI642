# Group3_CSCI642_Project: Using Behavioral Analysis to Detect Keyloggers

These are the parts to this project:
1. Creating a keylogger that runs on a Windows 10 system
2. Creating a detection system for the keylogger developed (and similar keyloggers already developed)

**Note**: Multiple attempts were explored to determine the best way to create the keylogger and detection system. The paper is based off of work done in the paper_files directory

## Install and Requirements
It is recommended to use a virtual environment in order to keep track of all packages
This can be done by doing the following:

```
python -m venv <name of virtual environment>
.\<name of virtual environment>\Scripts\activate
```

Then install the required packages using `pip` and `requirements.txt`:
```pip install -r requirements.txt```

## System Architecture
This contains only the important elements to highlight for phase 3 of the project (and will be updated for final submission)
```
├── data
│   ├──merged_data
│   │   └── combined_dataset.csv
├── phase 2
├── phase 3
│   ├── detection
│   │   ├── detector.py
│   │   ├── model.py
│   │   ├── random_forest.joblib
│   │   └── random_forest.pkl
│   ├── exe_files
│   ├── keylogger
│   │   ├── keylogger.py
│   │   └── stealth_keylogger.py
│   ├── utils
│   │   └── generate_data.py
│   ├── graphs
│   ├── paper_files
│   │   ├── balanced_dataset.csv
│   │   ├── random_forest.pkl
│   │   ├── stealth_keylogger.py
│   │   └── training_model.py
├── securecodingenv
├── requirements.txt
```
## Paper Files



**Note**: The following describes how one would implement another attempt (whose data was not utilized in the research paper), which also utilizes some code from the paper_files directory, however, not all were integrated completely (this will be fully merged for the final deliverable in phase 4)
## Keylogger Development and Deployment
There are 2 ways to run the keylogger:
1. In the phase3 directory, run the command ```python .\keylogger.py```
    - This will log all keystrokes to a file called `keystroke_log.csv` in the directory in which the file is run (i.e. phase3)
    - To stop the keylogger, simply press the ESC key

2. A `.exe` file has been provided in the phase3 directory. This was generated with the command ```pyinstaller --onefile .\keylogger.py```. This would be more similar to a real world attack, as not every computer would have Python installed (which is needed for the previous way to run the keylogger) and the `keylogger.exe` can be run in the command prompt/powershell of Window 10 computers

## Generating Training Data
The program `generate_data.py` (located in phase3 directory) collects information on all running processes on the device. This data is used to train the keylogger detector model. It labels processes as either 'benign' or 'keylogger' (this is based on user-inputted process id values (pid))

Information Collected:
- cpu
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
A Random Forest Model was developed and trained off of data from multiple Virtual Machines running Windows 10. This takes a snapshot of the current machine and runs the data against the trained model and outputs (with confidence) if there is a suspected keylogger or detected keylogger.
- Suspected Keylogger: Confidence of 40% - 70%
- Detected Keylogger: Confidence of >70%

The following shows how to run the detector (from the detection directory):
```python .\detector.py```

## Other Keyloggers Used for Data Training and Testing
These keyloggers were picked to help train the model and test if they could be detected as they are similar to the keylogger that was developed in this project. It's important that this tool does not only detect our keylogger, but other ones as well, as this tool's intended use is to prevent user space keyloggers on Windows 10 (not just the one developed)
- https://github.com/elxecutor/keylogger/tree/main
- https://github.com/ramprasathmk/keylogger
- https://github.com/creekmar/python-keylogger 

## Legal and Ethical Disclaimer
This keylogger was designed for **educational** and **research** purposes **only**!
