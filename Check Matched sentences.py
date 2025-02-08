"import os
import re
import csv
import logging
from bs4 import BeautifulSoup

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Define the parent directory where your subfolders are located
parent_directory = '/Volumes/SamsungSSD/2324'  # Updated to the TEST directory

# Create an output CSV file
output_csv = '/Volumes/SamsungSSD/cyber_scan_results_test.csv'  # Save results to SSD

# Define the keywords to search for (cyber-related)
cyber_keywords = [
    'cyber', 'cybersecurity', 'data breach', 'IT security', 'information security'
]

def clean_html(content):
    """"""Cleans HTML content and returns plain text.""""""
    soup = BeautifulSoup(content, 'html.parser')
    return soup.get_text(separator=' ').strip()

def extract_company_name(content):
    """"""Extracts the company name from the content.""""""
    match = re.search(r'([A-Za-z0-9\s\-\(\)&\']+?\s(?:Limited|Ltd|LLP|Incorporated|Inc|Company|Holdings))', content, re.IGNORECASE)
    if match:
        company_name = match.group(1).strip()
        return company_name
    return None

def extract_matched_sentence(content, keyword):
    """"""Extracts the matched sentence containing the keyword.""""""
    sentences = re.split(r'\.|\n', content)
    for sentence in sentences:
        if keyword.lower() in sentence.lower():
            return sentence.strip()
    return None

def scan_html_files_for_cyber(parent_directory):
    """"""Scans HTML files for cyber-related data.""""""
    results = []
    file_count = 0

    for subdir, _, files in os.walk(parent_directory):
        for filename in files:
            if filename.endswith('.html') or filename.endswith('.xhtml'):
                file_path = os.path.join(subdir, filename)
                logging.info(f""Scanning file: {filename} in {subdir}"")

                try:
                    with open(file_path, 'r', encoding='utf-8') as file:
                        content = file.read()
                        clean_content = clean_html(content)

                        for keyword in cyber_keywords:
                            matched_sentence = extract_matched_sentence(clean_content, keyword)
                            if matched_sentence:
                                logging.info(f""Keyword '{keyword}' found in: {filename}"")

                                company_name = extract_company_name(clean_content)
                                unaudited = ""Yes"" if ""unaudited"" in clean_content.lower() else ""No""
                                dormant = ""Yes"" if ""dormant"" in company_name.lower() else ""No""

                                results.append([filename, matched_sentence, company_name or 'N/A', unaudited, dormant])
                                break  # Break after the first keyword match

                except FileNotFoundError as e:
                    logging.error(f""File not found: {file_path}. Error: {str(e)}"")
                except Exception as e:
                    logging.error(f""An error occurred while processing {file_path}. Error: {str(e)}"")

                file_count += 1

    if results:
        with open(output_csv, mode='w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(['File Name', 'Matched Sentence (Cyber Context)', 'Company Name', 'Unaudited', 'Dormant'])
            writer.writerows(results)
            writer.writerow(['Total Files Processed', file_count])

        logging.info(f""Scan completed. Results saved to {output_csv}"")
    else:
        logging.warning(""No files containing relevant cyber mentions were found."")

# Run the function
scan_html_files_for_cyber(parent_directory)"