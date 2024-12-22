import subprocess
import argparse
import threading

# Set up argument parser
parser = argparse.ArgumentParser(description='Execute Python files in sequence, with an optional date argument for market_data_fetch.py and economic_data_fetch.py.')
parser.add_argument('--day', type=str, help='Optional date to be passed as --day to market_data_fetch.py and economic_data_fetch.py')

# Parse the arguments
args = parser.parse_args()

# Function to run a sequence of scripts
def run_sequence(files):
    for file in files:
        try:
            subprocess.run(["python", file], check=True)
        except subprocess.CalledProcessError as e:
            print(f"Error occurred while executing {file}: {e}")

# List of Python files to be executed before threading
initial_files = [
    "kill.py",
    "nuke.py",
    "market_data_fetch.py"
]

# Execute initial files
for file in initial_files:
    try:
        if file == "market_data_fetch.py":
            if args.day:
                subprocess.run(["python", file, "--day", args.day], check=True)
            else:
                subprocess.run(["python", file], check=True)
        else:
            subprocess.run(["python", file], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while executing {file}: {e}")

# Create two threads for the simultaneous execution
thread1 = threading.Thread(target=run_sequence, args=(["market_stats.py", "market_stats_header_flattener.py"],))
thread2 = threading.Thread(target=run_sequence, args=(["ta.py", "ta_gap_counter.py"],))

# Start both threads
thread1.start()
thread2.start()

# Wait for both threads to complete
thread1.join()
thread2.join()

# Continue with remaining files after threads are done
remaining_files = [
    "fluffcutter.py",
    "datefinder.py",
    "economic_data_fetch.py",
    "economic_stats.py",
    "economic_hedgecutter.py",
    "colorizer1.py",
    "batchmaker1.py",
    "batch_date_pruner.py"
]

for file in remaining_files:
    try:
        if file == "economic_data_fetch.py":
            if args.day:
                subprocess.run(["python", file, "--day", args.day], check=True)
            else:
                subprocess.run(["python", file], check=True)
        else:
            subprocess.run(["python", file], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while executing {file}: {e}")
