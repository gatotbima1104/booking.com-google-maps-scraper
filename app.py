import requests
from bs4 import BeautifulSoup
import re
import pandas as pd

def extract_emails(url):
    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an HTTPError for bad responses

        soup = BeautifulSoup(response.text, 'html.parser')
        email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
        emails = re.findall(email_pattern, soup.get_text())
        unique_emails = set(emails)
        return unique_emails
    except requests.RequestException as e:
        print(f"Failed to retrieve content from {url}. Error: {e}")
        return set()

# Read the Excel file using pandas
excel_file_path = 'hotel.xlsx'
df = pd.read_excel(excel_file_path)

# Assuming URLs are in column C starting from the second row (index 1)
urls_to_scrape = df.iloc[0:20, 3].tolist()

# Create a dictionary to store website URLs and corresponding emails as arrays
website_emails_dict = {}

# Iterate over the URLs and extract emails
for index, url in enumerate(urls_to_scrape, start=1):
    emails = extract_emails(url)

    # Add all websites to the dictionary, even if they have no emails
    website_emails_dict[url] = list(emails)

    print(f"Link number {index}: {url}")

# Create a new DataFrame with website URLs and email arrays
result_df = pd.DataFrame({'Website': list(website_emails_dict.keys()), 'Emails': list(website_emails_dict.values())})

# Replace null websites with a placeholder, e.g., 'no-value'
result_df['Website'] = result_df['Website'].fillna('no-value')

# Convert the 'Emails' column to strings and remove square brackets and empty strings
result_df['Emails'] = result_df['Emails'].apply(lambda emails: ', '.join(emails) if emails else '')

# Save the result to a new Excel file
result_excel_file = 'coba.xlsx'
result_df.to_excel(result_excel_file, index=False)

print(f"Extracted emails saved to {result_excel_file}")
