from bs4 import BeautifulSoup
import pandas as pd
import time
from datetime import datetime as dt
import aiohttp
import asyncio
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import os

def extract_mails(html):
    soup = BeautifulSoup(html, 'html.parser')
    emails = set()
    for mailto in soup.select('a[href^=mailto]'):
        email = mailto['href'].replace('mailto:', '')
        emails.add(email)

    return emails

async def fetch_mails(session, url):
    try:
        async with session.get(url, timeout=10) as response:            
            response.raise_for_status()  # Check for HTTP errors
            html = await response.text()
            return extract_mails(html)           

    except Exception as e:
        print(f"Error fetching {url}: {e}")
        return "Error fetching email"
    
async def main(urls, output_display):
    async with aiohttp.ClientSession() as session:
        tasks = [fetch_mails(session, url) for url in urls]
        email_result = await asyncio.gather(*tasks)

        for url, mails in zip(urls, email_result):
            output_display.insert(tk.END, f"fetched {url}: {', '.join(mails) or 'No mails found'}\n")
            output_display.see(tk.END)

        return email_result
    
def start_scraping(input_file, output_display):
    if not input_file:
        messagebox.showerror("Error", "Please select an input file.")
        return

    df = pd.read_excel(input_file, sheet_name='Sheet1')
    urls = df['URL'].to_list()

    start_time = time.time()
    output_display.insert(tk.END, f"Starting scraping at {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(start_time))}\n")

    timestamp = dt.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(os.path.dirname(input_file), f"output_emails_{timestamp}.xlsx")

    emails_list = asyncio.run(main(urls, output_display))

    if 'Email' not in df.columns:
        df['Email'] = ""

    df['Email'] = [', '.join(emails) for emails in emails_list]

    df.to_excel(output_file, index = False)
    end_time = time.time()
    duration = end_time-start_time
    output_display.insert(tk.END, f"Finished scraping at {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(end_time))}\n")
    output_display.insert(tk.END, f"Total duration: {duration:.2f} seconds\n")
    output_display.see(tk.END)

    messagebox.showinfo("Completed", f"Scraping completed successfully!\nOutput file: {output_file}")

def setup_gui():
    root = tk.Tk()
    root.title("Email Scraper")
    root.geometry("600x400")

    # Labels and Buttons
    tk.Label(root, text="Select Input Excel File:").grid(row=0, column=0, padx=10, pady=10)
    input_entry = tk.Entry(root, width=40)
    input_entry.grid(row=0, column=1)
    
    output_display = ScrolledText(root, wrap=tk.WORD, width=70, height=15)
    output_display.grid(row=2, column=0, columnspan=3, padx=10, pady=10)

    def browse_input():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        input_entry.delete(0, tk.END)
        input_entry.insert(0, file_path)

    tk.Button(root, text="Browse", command=browse_input).grid(row=0, column=2, padx=5)
    tk.Button(root, text="Start Scraping", command=lambda: start_scraping(input_entry.get(), output_display)).grid(row=1, column=1, pady=20)

    root.mainloop()

if __name__ == "__main__":
    setup_gui()
    """ input_file = "websites.xlsx"
    timestamp = dt.now().strftime('%Y%m%d_%H%M%S')  # Format: YYYYMMDD_HHMMSS
    output_file = f"urls_with_emails_{timestamp}.xlsx"

    scrape_mails_from_excel(input_file, output_file) """
    
