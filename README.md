# Group3_CSCI642_Project

There are two parts to this project. 
1. Creating a keylogger that runs on Windows 10 without the system knowing
2. Creating a keylogger detector to detect that the keylogger is running

## Keylogger
In order to run the keylogger type the following into the terminal
```
pip install pynput pywin32 psutil
python keylogger/research_keylogger.py
```
press ESC to stop

## Gathering Training Data
### generate_data.py
This is a program to collect data (including cpu, memory, threads, io, open_files, context switching, and command line arguments) on all processes running on the device for the purpose of using it to train the keylogger detector model. It will label processes as benign unless it is a recognized keylogger pid. 

In order to run the program, follow the steps below.

1. Inside generate_data.py add known keylogger pids to keylogger_pid_list
2. Run the following on the terminal
```
pip install pandas psutil
python detector/generate_data.py
```

## Detector
There will be two phases to this detector. 
1. Signature-based detection where the process's name and open files will be subject to keyword identification
2. Behavior-based detection where a proccess's information and behavior will be fed into a random forest classifier