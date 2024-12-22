import sys
import pandas as pd
import numpy as np
from scipy.stats import skew, kurtosis
import warnings
from datetime import datetime

# Suppress warnings related to invalid values (like log of non-positive numbers)
warnings.filterwarnings("ignore", category=RuntimeWarning)

# Load the Excel data
file_path = 'economic_data.xlsx'  # Assuming the spreadsheet is in the same directory as the script
excel_data = pd.read_excel(file_path, header=0, index_col=0)

# Initialize results dictionary
results = {}

# Process each indicator (1-dimensional data for each ticker)
for indicator in excel_data.columns:
    df = excel_data[indicator].dropna()  # Ensure valid data is available
    df.index = pd.to_datetime(df.index)

    print(f"Processing {indicator}: data points = {len(df)}")  # Debug: check number of data points

    if df.empty or df.isna().all():
        print(f"Skipping {indicator}: no valid data")  # Debug: indicate which indicator is skipped
        continue

    try:
        # Calculate relevant statistics for 1-dimensional economic data
        results[indicator] = {
            'Daily Return': df.pct_change(),
            'Log Return': np.log(df / df.shift(1)).replace([np.inf, -np.inf], np.nan),
            'Cumulative Return': (1 + df.pct_change()).cumprod() - 1,
            'Rolling Mean': df.rolling(window=30).mean(),
            'Rolling Std': df.rolling(window=30).std(),
            'Rolling Skewness': df.rolling(window=30).apply(lambda x: skew(x), raw=True),
            'Rolling Kurtosis': df.rolling(window=30).apply(lambda x: kurtosis(x), raw=True),
            'Z-Score': (df - df.mean()) / df.std(),
            'Annualized Volatility': df.pct_change().rolling(window=30).std() * np.sqrt(252),
            'Max Drawdown': (1 + df.pct_change()).cumprod() / (1 + df.pct_change()).cumprod().cummax() - 1,
            'EMA': df.ewm(span=30, adjust=False).mean(),
        }
        print(f"Successfully processed {indicator}")  # Debug: indicate success
    except Exception as e:
        print(f"Error processing {indicator}: {e}")  # Debug: indicate if an error occurs
        continue

# Check if any results were generated
if not results:
    print("No valid data was processed, results are empty.")
    sys.exit(1)

# Combine results into a DataFrame
try:
    combined_results = pd.concat(
        {(indicator, stat): pd.Series(result) if isinstance(result, (np.float64, float)) else result
         for indicator, indicator_results in results.items() for stat, result in indicator_results.items()},
        axis=1)
except ValueError as e:
    print(f"Concatenation error: {e}")
    sys.exit(1)

# Flatten the MultiIndex by joining the column levels
combined_results.columns = ['_'.join(col).strip() for col in combined_results.columns.values]

# Extract the original dates from the excel_data
dates = excel_data.index.strftime('%Y-%m-%d')

# Insert the 'Date' column with actual dates from the original data
combined_results.insert(0, 'Date', dates)

# Save the results to an Excel file
output_file = 'economic_stats.xlsx'  # Output file will be saved in the same directory
with pd.ExcelWriter(output_file) as writer:
    combined_results.to_excel(writer, index=False)

# Print completion message
print(f"Expanded statistical analysis completed and saved to {output_file}.")


