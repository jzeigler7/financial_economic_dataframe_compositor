import pandas as pd

# Load the Excel file
file_path = 'market_data.xlsx'

# Load the data from the correct sheet (as found, it's named 'Sheet')
df = pd.read_excel(file_path, sheet_name='Sheet')

# Extract the 'Date' column and remove duplicates
dates_df = df[['Date']].drop_duplicates()

# Save the unique dates to a CSV file
output_file_path = 'found_dates.csv'
dates_df.to_csv(output_file_path, index=False)

print(f"CSV file '{output_file_path}' has been created.")
