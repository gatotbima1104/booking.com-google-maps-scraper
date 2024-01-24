import pandas as pd
from emailextractor import EmailExtractor

def extract_emails(url):
    try:
        extractor = EmailExtractor(url)
        emails = extractor.get_emails()

        # You can filter or process the emails further if needed
        return emails
    except Exception as e:
        print(f"Failed to extract emails from {url}. Error: {e}")
        return []

# Read the Excel file using pandas
excel_file_path = 'hotel.xlsx'
df = pd.read_excel(excel_file_path)

# Assuming URLs are in column J starting from the second row (index 1)
urls_to_scrape = df.iloc[1:, 3].tolist()

# Create a dictionary to store website URLs and corresponding emails as arrays
website_emails_dict = {}

# Iterate over the URLs and extract emails
for index, url in enumerate(urls_to_scrape, start=2):  # Starting from the second row
    emails = extract_emails(url)

    # Add all websites to the dictionary, even if they have no emails
    website_emails_dict[url] = emails

    print(f"Link number {index}: {url} - Extracted emails: {emails}")

# Create a new DataFrame with website URLs and email arrays
result_df = pd.DataFrame({'Website': list(website_emails_dict.keys()), 'Emails': list(website_emails_dict.values())})

# Replace null websites with a placeholder, e.g., 'no-value'
result_df['Website'] = result_df['Website'].fillna('no-value')

# Convert the 'Emails' column to strings and remove square brackets and empty strings
result_df['Emails'] = result_df['Emails'].apply(lambda emails: ', '.join(emails) if emails else '')

# Save the result to a new Excel file
result_excel_file = 'result-with-py.xlsx'
result_df.to_excel(result_excel_file, index=False)

print(f"Extracted emails saved to {result_excel_file}")
