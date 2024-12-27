import os
import subprocess
import sys


def launch_dashboard():
    # Launch the ADDDELETE_1_Excel.py script
    script_path = r"ADDDELETE_1_Excel.py"
    subprocess.Popen(['python', script_path])

if __name__ == "__main__":
    # Start the dashboard
    launch_dashboard()
