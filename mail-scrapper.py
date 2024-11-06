import re
import requests
from bs4 import BeautifulSoup
import pandas as pd
import concurrent.futures as cf
import time
from datetime import datetime as dt

start_time = time.time()

def find_emails(text):
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    return re.findall(email_pattern, text)

def get_emails_from_website(url):
    with requests.Session() as session:
        try:
            response = session.get(url, timeout=10)
            # response = requests.get(url, timeout=10)
            response.raise_for_status()  # Check for HTTP errors
            soup = BeautifulSoup(response.text, 'html.parser')

            # Collect emails from mailto links
            emails = set()
            for mailto in soup.select('a[href^=mailto]'):
                email = mailto['href'].replace('mailto:', '')
                emails.add(email)

            # Collect emails in visible text
            """ page_text = soup.get_text()
            emails.update(find_emails(page_text)) """

            return ', '.join(emails) if emails else 'No emails found'

        except Exception as e:
            print(f"Error fetching {url}: {e}")
            return "Error fetching email"

# Read URLs from Excel
df = pd.read_excel('websites.xlsx', sheet_name='Sheet1')  # Adjust sheet name if needed

# Ensure 'Email' column exists or create it
if 'Email' not in df.columns:
    df['Email'] = ""

results = {}

# Fetch emails for each URL and update DataFrame
""" for i, row in df.iterrows():
    url = row['URL']
    print(f"Fetching emails from {url}...")
    email = get_emails_from_website(url)
    df.at[i, 'Email'] = email
    print(f"Found email(s): {email}") """

with cf.ThreadPoolExecutor(max_workers=5) as executor:
    future_to_url = {executor.submit(get_emails_from_website, row['URL']) : row['URL'] for _, row in df.iterrows()}

    for future in cf.as_completed(future_to_url):
        url = future_to_url[future]
        try:
            email = future.result()
            results[url] = email
            print(f"Fetched from {url} email: {email}")
        except Exception as e:
            print(f"Error in fetching from {url}")
            results[url] = f"Error: {e}"

df['Email'] = df['URL'].map(results)

timestamp = dt.now().strftime('%Y%m%d_%H%M%S')  # Format: YYYYMMDD_HHMMSS
filename = f"urls_with_emails_{timestamp}.xlsx"

# Write updated DataFrame back to Excel
df.to_excel(filename, index=False)
print(f"Email extraction complete. Results saved to {filename}")

end_time = time.time()
duration = end_time-start_time

print(f"Elapsed time: {duration:2f} seconds")
