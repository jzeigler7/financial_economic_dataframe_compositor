import os
import glob

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.realpath(__file__))

# File patterns for Excel spreadsheets in the script's directory
file_patterns = [os.path.join(script_dir, "*.xlsx"), os.path.join(script_dir, "*.xls")]

# Loop through the patterns and delete the files
for pattern in file_patterns:
    files = glob.glob(pattern)
    for file in files:
        try:
            os.remove(file)
            print(f"Deleted: {file}")
        except Exception as e:
            print(f"Failed to delete {file}: {e}")

# File to delete in the same directory
file_to_delete = os.path.join(script_dir, "ta_gap_count.txt")

# Delete market_stats_gap_count.txt if it exists
if os.path.exists(file_to_delete):
    try:
        os.remove(file_to_delete)
        print(f"Deleted: {file_to_delete}")
    except Exception as e:
        print(f"Failed to delete {file_to_delete}: {e}")
else:
    print(f"File not found: {file_to_delete}")

print("All specified files in the script's directory have been deleted.")

