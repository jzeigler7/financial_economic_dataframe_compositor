import yfinance as yf
import pandas as pd
import json
from datetime import datetime
import argparse
import os
from tqdm import tqdm  # For progress bars

# Setup argument parser to accept a custom day argument
parser = argparse.ArgumentParser(description='Download stock data for tickers and save it to an Excel file.')
parser.add_argument('--day', type=str, help='Specify the current calendar day in YYYY-MM-DD format (optional)')
args = parser.parse_args()

# Load the JSON file with ticker symbols
with open('tickers.json', 'r') as file:
    tickers_data = json.load(file)

# Extract tickers from the market_data section
tickers = list(tickers_data["market_data"].keys())

# Set the start date
start_date = "2007-08-01"

# Determine the current "day"
if args.day:
    try:
        end_date = datetime.strptime(args.day, '%Y-%m-%d').strftime('%Y-%m-%d')
    except ValueError:
        print("Error: Invalid date format. Please use YYYY-MM-DD.")
        exit(1)
else:
    end_date = datetime.today().strftime('%Y-%m-%d')

# Check if the Excel file already exists
output_file = 'market_data.xlsx'
existing_data = None

if os.path.exists(output_file):
    print(f"Found existing file: {output_file}. Loading it to check for gaps.")

    # Load the existing spreadsheet
    existing_data = pd.read_excel(output_file)

    # Find the most recent date in the existing file
    last_recorded_date = existing_data['Date'].max()

    # Set the start date for new data to be the day after the last recorded date
    start_date = (pd.to_datetime(last_recorded_date) + pd.DateOffset(1)).strftime('%Y-%m-%d')

    print(f"Fetching data starting from {start_date} to fill the gap.")

# If there's no gap or nothing to append, exit the program
if start_date >= end_date:
    print(f"Data is already up-to-date. No new data to fetch.")
    exit(0)

# Create a dictionary to hold all ticker data
all_ticker_data = {}

# Loop through all tickers and download the historical market data for the gap period
print("\nDownloading ticker data:")
for ticker in tqdm(tickers):
    data = yf.download(ticker, start=start_date, end=end_date, progress=False)

    # Ensure we only keep the desired columns (Open, High, Low, Close, Adj Close, Volume)
    if not data.empty:
        all_ticker_data[ticker] = data[['Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']]
    else:
        print(f"No data found for {ticker}.")

# Concatenate all the data into one DataFrame with multi-level columns
if all_ticker_data:
    # Align all data to the same date index to ensure consistency
    all_dates = pd.date_range(start=start_date, end=end_date, freq='B')  # Using business days
    for ticker in all_ticker_data:
        all_ticker_data[ticker] = all_ticker_data[ticker].reindex(all_dates)

    new_data_df = pd.concat(all_ticker_data.values(), keys=all_ticker_data.keys(), axis=1)

    # Flatten the multi-index columns for easier writing to Excel
    new_data_df.columns = ['_'.join(col).strip() for col in new_data_df.columns.values]

    # Reset index to make the date a column
    new_data_df.reset_index(inplace=True)
    new_data_df.rename(columns={'index': 'Date'}, inplace=True)

    # **Remove rows with at least 25% missing data before combining with existing data**
    data_columns = new_data_df.columns.drop('Date')

    # Calculate fraction of missing data per row
    fraction_missing = new_data_df[data_columns].isnull().mean(axis=1)

    # Number of rows before pruning
    num_rows_before = new_data_df.shape[0]

    # Remove rows with at least 25% missing data
    new_data_df = new_data_df[fraction_missing < 0.25].reset_index(drop=True)

    # Number of rows after pruning
    num_rows_after = new_data_df.shape[0]

    print(f"\nNew data: Rows before pruning: {num_rows_before}, Rows after pruning: {num_rows_after}")

    if num_rows_before > num_rows_after:
        num_rows_removed = num_rows_before - num_rows_after
        print(f"Pruned {num_rows_removed} rows from new data due to insufficient data.")
    else:
        print("No rows were pruned from new data.")

    # If there is existing data, append the new data
    if existing_data is not None:
        # **Process existing_data to remove rows with at least 25% missing data**
        existing_data_columns = existing_data.columns.drop('Date')
        fraction_missing_existing = existing_data[existing_data_columns].isnull().mean(axis=1)

        # Number of rows before pruning
        num_rows_before_existing = existing_data.shape[0]

        existing_data = existing_data[fraction_missing_existing < 0.25].reset_index(drop=True)

        # Number of rows after pruning
        num_rows_after_existing = existing_data.shape[0]

        print(f"\nExisting data: Rows before pruning: {num_rows_before_existing}, Rows after pruning: {num_rows_after_existing}")

        if num_rows_before_existing > num_rows_after_existing:
            num_rows_removed_existing = num_rows_before_existing - num_rows_after_existing
            print(f"Pruned {num_rows_removed_existing} rows from existing data due to insufficient data.")
        else:
            print("No rows were pruned from existing data.")

        # Combine existing and new data
        combined_data = pd.concat([existing_data, new_data_df], ignore_index=True)
        combined_data.drop_duplicates(subset='Date', inplace=True)
    else:
        combined_data = new_data_df

    # **Final check to remove any rows with at least 25% missing data after combining**
    data_columns_combined = combined_data.columns.drop('Date')
    fraction_missing_combined = combined_data[data_columns_combined].isnull().mean(axis=1)

    # Number of rows before final pruning
    num_rows_before_combined = combined_data.shape[0]

    combined_data = combined_data[fraction_missing_combined < 0.25].reset_index(drop=True)

    # Number of rows after final pruning
    num_rows_after_combined = combined_data.shape[0]

    print(f"\nCombined data: Rows before final pruning: {num_rows_before_combined}, Rows after final pruning: {num_rows_after_combined}")

    if num_rows_before_combined > num_rows_after_combined:
        num_rows_removed_combined = num_rows_before_combined - num_rows_after_combined
        print(f"Pruned {num_rows_removed_combined} rows from combined data due to insufficient data.")
    else:
        print("No rows were pruned from combined data.")

    # Save the updated data back to the Excel file
    combined_data.to_excel(output_file, index=False)
    print(f"\nNew data appended and saved to {output_file}")
else:
    print("No new data to append.")
