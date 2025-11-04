# Keylogger_Detector_CSCI642

## generate_data.py
This is a program to collect data on all processes running on the device for the purpose of using it to train the keylogger detector model. It will label processes as benign unless it is a recognized keylogger pid. In order to run the program, follow the steps below.

1. Inside generate_data.py add known keylogger pids to keylogger_pid_list
2. Run the following on the terminal
```
pip install pandas psutil
python generate_data.py
```

