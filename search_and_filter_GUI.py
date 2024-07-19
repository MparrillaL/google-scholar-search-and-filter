import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import requests
from bs4 import BeautifulSoup
from serpapi import GoogleSearch
import datetime

# Function to perform a Google Scholar search using SerpApi and save results to an Excel file
def search_and_save_to_excel(api_key, query, file_name, year_lower_bound, year_upper_bound, num_results=60):
    all_results = []

    # Define search parameters
    params = {
        "engine": "google_scholar",
        "q": query,
        "num": 20,  # Assuming the API returns up to 20 results per page
        "as_ylo": year_lower_bound,
        "as_yhi": year_upper_bound,
        "api_key": api_key
    }

    for start in range(0, num_results, 20):
        params["start"] = start
        search = GoogleSearch(params)
        results = search.get_dict()

        if 'organic_results' in results:
            all_results.extend(results['organic_results'])
        else:
            print("No more results found.")
            break

    if all_results:
        df = pd.json_normalize(all_results)
        df.to_excel(f'{file_name}.xlsx', index=False)
        print(f"Results saved to {file_name}.xlsx")
    else:
        print("No organic results found.")

# Function to check if a webpage contains specific keywords
def check_keywords_in_page(url, keywords):
    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        page_text = soup.get_text().lower()

        for keyword in keywords:
            if keyword.lower() in page_text:
                return True
        return False
    except requests.RequestException as e:
        print(f"Error fetching {url}: {e}")
        return False

# Function to filter an Excel file by checking keywords within the webpage content
def filter_excel_by_page_content(input_file, keywords, output_file):
    df = pd.read_excel(input_file)
    
    if 'link' not in df.columns:
        print("'link' column not found in the DataFrame.")
        return

    filtered_links = []
    for url in df['link']:
        if check_keywords_in_page(url, keywords):
            filtered_links.append(url)

    if filtered_links:
        filtered_df = df[df['link'].isin(filtered_links)]
        filtered_df.to_excel(output_file, index=False)
        print(f"Filtered results saved to {output_file}")
    else:
        print("No pages contained the specified keywords.")

def filter_excel_by_topic(input_file, topic, output_file):
    df = pd.read_excel(input_file)
    
    # Print the columns of the DataFrame to verify which columns are present
    print("Columns in the DataFrame:", df.columns)
    
    # Filter by both 'Snippet' and 'Title' columns if they exist
    if 'Snippet' in df.columns and 'Title' in df.columns:
        filtered_df = df[df['Snippet'].str.contains(topic, case=False, na=False) | df['Title'].str.contains(topic, case=False, na=False)]
    elif 'Snippet' in df.columns:
        filtered_df = df[df['Snippet'].str.contains(topic, case=False, na=False)]
    elif 'Title' in df.columns:
        filtered_df = df[df['Title'].str.contains(topic, case=False, na=False)]
    else:
        print("Neither 'Snippet' nor 'Title' column found.")
        return
    
    filtered_df.to_excel(output_file, index=False)
    print(f"Filtered results saved to {output_file}")

class SearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Search and Filter App")

        tk.Label(root, text="SerpApi Key:").grid(row=0, column=0, padx=10, pady=10)
        self.api_key_entry = tk.Entry(root, width=50)
        self.api_key_entry.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(root, text="Search Query:").grid(row=1, column=0, padx=10, pady=10)
        self.query_entry = tk.Entry(root, width=50)
        self.query_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(root, text="Topic:").grid(row=2, column=0, padx=10, pady=10)
        self.topic_entry = tk.Entry(root, width=50)
        self.topic_entry.grid(row=2, column=1, padx=10, pady=10)

        tk.Label(root, text="Keywords (comma separated):").grid(row=3, column=0, padx=10, pady=10)
        self.keywords_entry = tk.Entry(root, width=50)
        self.keywords_entry.grid(row=3, column=1, padx=10, pady=10)

        tk.Label(root, text="Year Lower Bound:").grid(row=4, column=0, padx=10, pady=10)
        self.year_lower_entry = tk.Entry(root, width=50)
        self.year_lower_entry.grid(row=4, column=1, padx=10, pady=10)

        tk.Label(root, text="Year Upper Bound:").grid(row=5, column=0, padx=10, pady=10)
        self.year_upper_entry = tk.Entry(root, width=50)
        self.year_upper_entry.grid(row=5, column=1, padx=10, pady=10)

        self.search_button = tk.Button(root, text="Search and Filter", command=self.run_search_and_filter)
        self.search_button.grid(row=6, column=0, columnspan=2, pady=20)

    def run_search_and_filter(self):
        api_key = self.api_key_entry.get()
        query = self.query_entry.get()
        topic = self.topic_entry.get()
        keywords = self.keywords_entry.get().split(',')
        year_lower_bound = int(self.year_lower_entry.get())
        year_upper_bound = int(self.year_upper_entry.get())

        # Generate a unique file name based on the current date and time
        file_name = f"serpapi_results_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        initial_filtered_file_name = "initial_filtered_results.xlsx"
        final_filtered_file_name = "final_filtered_results.xlsx"

        # Perform the search and save results to an Excel file
        search_and_save_to_excel(api_key, query, file_name, year_lower_bound, year_upper_bound)
        
        # Filter the initial results by topic in the snippet or title
        filter_excel_by_topic(file_name + '.xlsx', topic, initial_filtered_file_name)
        
        # Further filter the initial filtered results by checking if the webpage content contains the specified keywords
        filter_excel_by_page_content(initial_filtered_file_name, keywords, final_filtered_file_name)
        
        messagebox.showinfo("Completed", f"Filtered results saved to {final_filtered_file_name}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SearchApp(root)
    root.mainloop()
