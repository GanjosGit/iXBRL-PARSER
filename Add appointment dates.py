"import os
import re
import logging
from bs4 import BeautifulSoup
import pandas as pd

# Setup logging to monitor the progress
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Paths
input_excel = '/Volumes/SamsungSSD/Companies House Accounts Scan YTD 24.xlsx'  # Path to your input XLSX file
output_excel = '/Volumes/SamsungSSD/Companies_House_Accounts_Scan_YTD_24_Results.xlsx'  # Path to save the results XLSX file

# Define the appointed pattern for both 2023 and 2024 (e.g., appointed 5th October 2023 or appointed 7th July 2024)
appointed_pattern = r'appointed\s(\d{1,2}(?:st|nd|rd|th)?\s(?:[A-Za-z]+|\d{1,2})\s(?:2023|2024))'

# Function to extract all appointed dates in the ""appointed day-month year"" format (for both 2023 and 2024)
def find_appointed_dates(content):
    """"""Searches for the keyword 'appointed' followed by valid date patterns for both 2023 and 2024.""""""
    soup = BeautifulSoup(content, 'lxml')  # Using lxml for faster parsing
    text = soup.get_text(separator=' ').lower()  # Convert to lower case for case-insensitive search

    # Search for all appointed date patterns (both 2023 and 2024)
    matches = re.findall(appointed_pattern, text, re.IGNORECASE)
    
    if matches:
        # Join all matches (dates) as a comma-separated string
        return ', '.join(matches)
    
    return None

# Step 1: Load the input Excel file into a Pandas DataFrame
df = pd.read_excel(input_excel)

# Step 2: Add a new column for 'Appointed Dates'
df['Appointed Dates'] = ''

# Step 3: Iterate over each row, process the file, and extract the appointed dates
for index, row in df.iterrows():
    file_path = row['File Path']  # Ensure 'File Path' column exists in your Excel

    if os.path.exists(file_path):  # Check if the file exists
        try:
            logging.info(f""Processing file: {file_path}"")

            # Open the file and search for appointed dates
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

                # Search for appointed dates (both 2023 and 2024)
                appointed_dates = find_appointed_dates(content)

                # Update the DataFrame with the appointed dates result
                df.at[index, 'Appointed Dates'] = appointed_dates or 'N/A'

        except Exception as e:
            logging.error(f""Error processing {file_path}: {e}"")
            df.at[index, 'Appointed Dates'] = 'Error'
    else:
        logging.warning(f""File not found: {file_path}"")
        df.at[index, 'Appointed Dates'] = 'File Not Found'

# Step 4: Save the updated DataFrame to a new Excel file
df.to_excel(output_excel, index=False)
logging.info(f""Results saved to {output_excel}"")"