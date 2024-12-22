import win32com.client
import psutil
import os

def close_all_excel_files_without_saving():
    try:
        # Connect to any running Excel application
        excel_app = win32com.client.Dispatch("Excel.Application")

        # Get the list of all open workbooks
        workbooks = excel_app.Workbooks

        # Loop through each workbook and close it without saving
        for workbook in workbooks:
            workbook.Close(SaveChanges=False)

        # Optionally, you can quit the application if needed
        excel_app.Quit()

        print("All open Excel files have been closed without saving.")
    except Exception as e:
        print(f"An error occurred while closing Excel files: {e}")

def close_all_txt_and_csv_files():
    # Define the common applications that open text and CSV files
    target_processes = ["notepad.exe", "excel.exe"]

    # Iterate over all running processes
    for process in psutil.process_iter(['pid', 'name', 'open_files']):
        try:
            # Check if the process name matches one of the target processes
            if process.info['name'].lower() in target_processes:
                open_files = process.info['open_files']

                if open_files:
                    # Iterate through all open files for this process
                    for file in open_files:
                        file_path = file.path.lower()
                        # Check if the file is a .txt or .csv file
                        if file_path.endswith('.txt') or file_path.endswith('.csv'):
                            # Terminate the process if any .txt or .csv files are open
                            process.terminate()
                            print(f"Closed process '{process.info['name']}' which had open {file_path}")
                            break
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

if __name__ == "__main__":
    close_all_excel_files_without_saving()
    close_all_txt_and_csv_files()

