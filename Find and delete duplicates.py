"import csv
import os
import re
import logging
from bs4 import BeautifulSoup
import shutil
import pandas as pd

# Setup logging to see what's going on
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Paths
input_excel = '/Volumes/SamsungSSD/23_24 Final.xlsx'  # Path to your input XLSX file
backup_excel = '/Volumes/SamsungSSD/23_24_Final_backup.xlsx'  # Backup XLSX file
output_csv = '/Volumes/SamsungSSD/23_24_Final_results.csv'  # Output CSV file

# Step 1: Create a backup of the original Excel file
shutil.copy(input_excel, backup_excel)
logging.info(f""Backup created: {backup_excel}"")

# Define patterns or keywords for each item
turnover_patterns = [r'turnover', r'total revenue', r'net revenue']
company_name_patterns = [r'([A-Za-z0-9\s\-\(\)&\']+?\s(?:Limited|Ltd|LLP|Incorporated|Inc|Company|Holdings))']
registration_number_patterns = [r'registration number', r'company number', r'number:\s?\d{5,}', r'number\s?\d{5,}']

# Function to extract matching text based on patterns
def find_match(content, patterns):
    """"""Searches for a keyword or pattern in the content and returns the first match found.""""""
    for pattern in patterns:
        match = re.search(pattern, content, re.IGNORECASE)
        if match:
            return match.group(0)
    return None

# Function to extract the value near a keyword (scanning 1-3 cells rightwards)
def extract_value_near_keyword(content, keyword):
    """"""Searches for the keyword and extracts the value next to it.""""""
    soup = BeautifulSoup(content, 'lxml')  # Using lxml for faster parsing
    text = soup.get_text(separator=' ').lower()  # Convert to lower case for case-insensitive search
    keyword_index = text.find(keyword.lower())
    
    if keyword_index != -1:
        # Extract surrounding text after the keyword
        surrounding_text = text[keyword_index:].split()

        # Look 1-3 positions ahead for the first valid numerical value with multiple digits
        for i in range(1, 4):
            if len(surrounding_text) > i and surrounding_text[i].replace(',', '').isdigit() and len(surrounding_text[i]) > 3:
                return surrounding_text[i]  # Return the first valid number
    return None

# Function to search for registration number in the header
def extract_registration_from_header(content):
    """"""Attempts to find the registration number within the first few lines or pages.""""""
    soup = BeautifulSoup(content, 'lxml')  # Parse with lxml for better performance
    text = soup.get_text(separator=' ').lower()
    
    # Limit the search to the first 500 characters for efficiency (considering headers)
    header_section = text[:500]  
    match = find_match(header_section, registration_number_patterns)
    
    return match

# Step 2: Load the input Excel file into a Pandas DataFrame
df = pd.read_excel(input_excel)

# Step 3: Create new columns for 'Company Name Check', 'Registration Number', and 'Turnover'
df['Company Name Check'] = ''
df['Registration Number'] = ''
df['Turnover'] = ''

# Step 4: Iterate over each row and process the file path
for index, row in df.iterrows():
    file_path = row['File Path']  # Ensure 'File Path' column exists in your Excel

    if os.path.exists(file_path):  # Check if file exists
        try:
            logging.info(f""Processing file: {file_path}"")

            # Open the file and search for company name, turnover, and registration number
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()

                # Search for company name
                company_name = find_match(content, company_name_patterns)

                # Search for turnover (look for the value 1-3 positions right of the keyword)
                turnover = None
                for keyword in turnover_patterns:
                    turnover = extract_value_near_keyword(content, keyword)
                    if turnover:
                        break

                # Search for registration number (prefer header)
                reg_number = extract_registration_from_header(content)
                if not reg_number:  # If not in the header, try a broader search
                    reg_number = find_match(content, registration_number_patterns)

                # Update the DataFrame with the results
                df.at[index, 'Company Name Check'] = company_name or 'N/A'
                df.at[index, 'Turnover'] = turnover or 'N/A'
                df.at[index, 'Registration Number'] = reg_number or 'N/A'

        except Exception as e:
            logging.error(f""Error processing {file_path}: {e}"")
            df.at[index, 'Company Name Check'] = 'Error'
            df.at[index, 'Turnover'] = 'Error'
            df.at[index, 'Registration Number'] = 'Error'

    else:
        logging.warning(f""File not found: {file_path}"")
        df.at[index, 'Company Name Check'] = 'File Not Found'
        df.at[index, 'Turnover'] = 'File Not Found'
        df.at[index, 'Registration Number'] = 'File Not Found'

# Step 5: Save the updated DataFrame to a new CSV file
df.to_csv(output_csv, index=False)
logging.info(f""Results saved to {output_csv}"")"