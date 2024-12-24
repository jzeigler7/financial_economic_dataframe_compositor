# Financial Data Analysis and Processing Suite

## Overview
This project is a comprehensive suite of Python scripts designed to automate the collection, processing, and analysis of financial and economic data. Leveraging data from sources like Yahoo Finance and FRED, the tools generate insights through technical indicators, economic statistics, and visualizations, culminating in structured, consolidated reports.

## Features
### Data Collection and Fetching
- **Market Data Fetch** (`market_data_fetch.py`): Collects historical market data for stocks and indices using Yahoo Finance.
- **Economic Data Fetch** (`economic_data_fetch.py`): Fetches economic indicators from the FRED database.
- **Date Extraction** (`datefinder.py`): Extracts unique dates from market data for consistency across datasets.

### Data Analysis
- **Technical Analysis** (`ta.py`): Computes indicators like SMA, EMA, RSI, and Bollinger Bands.
- **Market Statistics** (`market_stats.py`): Analyzes market trends through rolling statistics, Sharpe ratios, and drawdowns.
- **Economic Statistics** (`economic_stats.py`): Processes economic data to derive metrics like rolling skewness, kurtosis, and max drawdown.
- **Gap Counting** (`ta_gap_counter.py`): Identifies rows with missing data in technical analysis files.

### Data Cleaning and Formatting
- **Fluff Cutter** (`fluffcutter.py`): Removes rows with excessive missing data across datasets.
- **Economic Hedge Cutter** (`economic_hedgecutter.py`): Aligns economic datasets by removing rows without corresponding dates in a reference file.
- **Batch Date Pruner** (`batch_date_pruner.py`): Deletes specific columns from a consolidated dataset.
- **Header Flattener** (`market_stats_header_flattener.py`): Simplifies multi-level column headers in Excel files.

### Data Consolidation
- **Batch Maker** (`batchmaker1.py`): Merges multiple datasets into a single Excel file for streamlined reporting.
- **File Coloring** (`colorizer1.py`): Applies color coding to Excel sheets for better readability.

### Automation and Management
- **Data Collection Orchestrator** (`data_collection.py`): Automates the execution of scripts in a defined workflow, supporting threading for parallel execution.
- **Kill Script** (`kill.py`): Closes open Excel and text file processes.
- **Nuke Script** (`nuke.py`): Deletes all temporary and intermediary files in the working directory.

## Installation
1. **Clone the Repository**:
   ```bash
   git clone <repository_url>
   cd <repository_name>
   ```
2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Configure Files**:
   - **API Keys**: Update `FRED_API_KEY` in `economic_data_fetch.py` with your FRED API key.
   - **Ticker Configuration**: Ensure `tickers.json` is present with the appropriate structure:
     ```json
     {
       "market_data": {
         "AAPL": "Apple Inc.",
         "GOOGL": "Alphabet Inc."
       },
       "economic_data": {
         "GDP": "Gross Domestic Product",
         "CPIAUCSL": "Consumer Price Index"
       }
     }
     ```

## Usage
### Running Individual Scripts
Execute individual scripts directly to process specific tasks. For example:
```bash
python market_data_fetch.py --day YYYY-MM-DD
```

### Automating the Workflow
Use the `data_collection.py` script to run the entire pipeline:
```bash
python data_collection.py --day YYYY-MM-DD
```

### Consolidating Data
Merge datasets into a single batch file using:
```bash
python batchmaker1.py
```

## Workflow Example
1. **Fetch Data**:
   - `market_data_fetch.py` downloads market data to `market_data.xlsx`.
   - `economic_data_fetch.py` retrieves economic indicators to `economic_data.xlsx`.
2. **Analyze Data**:
   - `ta.py` generates technical indicators in `ta_data.xlsx`.
   - `market_stats.py` computes rolling statistics.
3. **Clean and Format**:
   - `fluffcutter.py` removes rows with missing values.
   - `economic_hedgecutter.py` aligns economic datasets by date.
4. **Consolidate and Colorize**:
   - `batchmaker1.py` merges datasets into `batch1.xlsx`.
   - `colorizer1.py` applies color coding for visualization.
5. **Finalize**:
   - Run `market_stats_header_flattener.py` to simplify headers.

## Key Files
### Configuration Files
- **`tickers.json`**: Maps market and economic tickers to descriptive names.

### Data Processing Scripts
- **`market_data_fetch.py`**: Fetches stock and index data.
- **`economic_data_fetch.py`**: Downloads economic indicators.
- **`ta.py`**: Computes technical indicators.
- **`market_stats.py`**: Processes market statistics.
- **`economic_stats.py`**: Analyzes economic data.

### Cleaning and Formatting
- **`fluffcutter.py`**: Removes incomplete rows.
- **`economic_hedgecutter.py`**: Filters rows by date alignment.
- **`batch_date_pruner.py`**: Deletes specific columns from Excel files.
- **`market_stats_header_flattener.py`**: Simplifies multi-level headers.

### Automation and Utility
- **`data_collection.py`**: Manages the end-to-end workflow.
- **`kill.py`**: Closes Excel and text editor processes.
- **`nuke.py`**: Cleans up intermediary files.
- **`colorizer1.py`**: Applies color formatting to Excel files.

## Dependencies
- `pandas`, `numpy`, `scipy`
- `openpyxl`
- `yfinance`
- `fredapi`
- `tqdm`
- `argparse`
