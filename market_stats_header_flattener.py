import pandas as pd
import sys
import os

def flatten_headers(input_file, output_file=None, separator='_'):
    """
    Flattens multi-level column headers in an Excel file by combining them into a single level.

    Parameters:
    - input_file: Path to the input Excel file.
    - output_file: Path to save the modified Excel file. If None, overwrites the input file.
    - separator: String to separate the combined header levels.
    """
    if output_file is None:
        output_file = input_file  # Overwrite the original file

    # Check if the input file exists
    if not os.path.exists(input_file):
        print(f"Error: The file '{input_file}' does not exist.")
        sys.exit(1)

    try:
        # Read the Excel file with multi-level headers
        df = pd.read_excel(input_file, header=[0, 1], index_col=0)
    except ValueError as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

    # Check if the columns are MultiIndex
    if not isinstance(df.columns, pd.MultiIndex):
        print("The Excel file does not have multi-level headers. No changes made.")
        sys.exit(0)

    # Flatten the MultiIndex columns by joining with the specified separator
    df.columns = [f"{ticker}{separator}{measure}" if measure else f"{ticker}"
                  for ticker, measure in df.columns]

    # Save the flattened DataFrame to Excel
    df.to_excel(output_file, index=True)
    print(f"Headers flattened and saved to '{output_file}'.")

if __name__ == "__main__":
    # Define the input and output file paths
    input_file = r'C:\Users\jacob\OneDrive\Documents\ZAP\market_stats.xlsx'

    # Optionally, define a different output file to avoid overwriting
    # For example:
    # output_file = r'C:\Users\jacob\OneDrive\Documents\ZAP\market_stats_flattened.xlsx'
    # If you want to overwrite the original file, set output_file to input_file
    output_file = input_file  # Overwrite the original file

    # Call the function to flatten headers
    flatten_headers(input_file, output_file)
