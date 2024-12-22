import sys
import pandas as pd
import numpy as np
from scipy.stats import skew, kurtosis
import warnings
from datetime import datetime
import yfinance as yf
from openpyxl import load_workbook

# Suppress warnings related to invalid values (like log of non-positive numbers)
warnings.filterwarnings("ignore", category=RuntimeWarning)

# Load the Excel data
file_path = r'C:\Users\jacob\OneDrive\Documents\ZAP\market_data.xlsx'  # Path to the input file
excel_data = pd.read_excel(file_path, header=0, index_col=0)

# Print the existing date range before any modifications
print("Original data date range:")
print(excel_data.index.min(), "to", excel_data.index.max())

# Ensure the index (dates) includes all business days from August 2, 2010 to today
today = datetime.today().strftime('%Y-%m-%d')
expected_dates = pd.date_range(start='2007-08-01', end=today, freq='B')  # Business days

# Reindex to include missing dates
excel_data = excel_data.reindex(expected_dates)

# Function to extract ticker and attribute from combined labels
def extract_ticker_and_attribute(label):
    parts = label.rsplit('_', 1)  # Split the label by the last underscore
    return parts[0], parts[1] if len(parts) > 1 else None

# Parse and store data for tickers and attributes
parsed_results = {}
for label in excel_data.columns:
    ticker, attribute = extract_ticker_and_attribute(label)
    if ticker not in parsed_results:
        parsed_results[ticker] = {}
    parsed_results[ticker][attribute] = excel_data[label]

# Print the tickers in the data and check their Close price availability
print("Tickers in data:", list(parsed_results.keys()))
results = {}
benchmark_ticker = '^GSPC'  # Adjust this to match your data

# Check if benchmark is in data; if not, download it
if benchmark_ticker in parsed_results and 'Close' in parsed_results[benchmark_ticker]:
    benchmark = parsed_results[benchmark_ticker]['Close'].dropna()
else:
    print(f"Benchmark ticker '{benchmark_ticker}' not found in the data. Downloading benchmark data.")
    benchmark_data = yf.download(benchmark_ticker, start='2010-08-01', end=today)
    benchmark = benchmark_data['Close'].dropna()

# Ensure benchmark index is datetime
benchmark.index = pd.to_datetime(benchmark.index)

# Process each ticker
for ticker, data in parsed_results.items():
    if 'Close' in data:
        df = data['Close'].dropna()
        df.index = pd.to_datetime(df.index)
        print(f"Processing {ticker}, Close data available")

        # Align the ticker data with the benchmark
        aligned_df, aligned_benchmark = df.align(benchmark, join='inner')

        # Calculate various statistics for the ticker
        results[ticker] = {
            'Daily Return': df.pct_change(),
            'Log Return': np.log(df / df.shift(1)),
            'Cumulative Return': (1 + df.pct_change()).cumprod() - 1,
            'Rolling Mean': df.rolling(window=30).mean(),
            'Rolling Std': df.rolling(window=30).std(),
            'Rolling Variance': df.rolling(window=30).var(),
            'Rolling Skewness': df.rolling(window=30).apply(lambda x: skew(x), raw=True),
            'Rolling Kurtosis': df.rolling(window=30).apply(lambda x: kurtosis(x), raw=True),
            'Auto-Correlation': df.rolling(window=30).apply(lambda x: x.autocorr(lag=1)),
            'Z-Score': (df - df.mean()) / df.std(),
            'Annualized Volatility': df.pct_change().rolling(window=30).std() * np.sqrt(252),
            'Sharpe Ratio': ((df.pct_change() - 0.01 / 252).rolling(window=30).mean() /
                             df.pct_change().rolling(window=30).std() * np.sqrt(252)),
            'Max Drawdown': (1 + df.pct_change()).cumprod() / (1 + df.pct_change()).cumprod().cummax() - 1,
            'EMA': df.ewm(span=30, adjust=False).mean(),
            'Rolling Beta': df.pct_change().rolling(window=30).cov(benchmark.pct_change()) /
                            benchmark.pct_change().rolling(window=30).var(),
            'Rolling Correlation': df.pct_change().rolling(window=30).corr(benchmark.pct_change()),
        }

# Check if any results were generated
if not results:
    print("No valid data was processed, results are empty.")
    sys.exit(1)

# Combine results into one DataFrame
combined_results = pd.concat(
    {(ticker, stat): pd.Series(result) if isinstance(result, (np.float64, float)) else result
     for ticker, ticker_results in results.items() for stat, result in ticker_results.items()},
    axis=1)

# Convert the datetime index to date-only strings
combined_results.index = combined_results.index.strftime('%Y-%m-%d')

# Save the results to an Excel file
output_file = r'C:\Users\jacob\OneDrive\Documents\ZAP\market_stats.xlsx'  # Path to the output file
with pd.ExcelWriter(output_file) as writer:
    combined_results.to_excel(writer)

# Delete the third row if needed
wb = load_workbook(output_file)
ws = wb.active
ws.delete_rows(3)
wb.save(output_file)

print(f"Expanded statistical analysis completed and saved to {output_file}.")
