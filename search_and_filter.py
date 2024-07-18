import pandas as pd
import openpyxl
from serpapi import GoogleSearch
import datetime

# Define your SerpApi key
API_KEY = "YOUR_KEY_HERE"

# Function to perform a Google Scholar search using SerpApi and save results to an Excel file
def search_and_save_to_excel(api_key, query, file_name, num_results, year_lower_bound, year_upper_bound):
    # Define search parameters
    params = {
        "engine": "google_scholar",
        "q": query,
        "num": num_results,
        "as_ylo": year_lower_bound,
        "as_yhi": year_upper_bound,
        "api_key": api_key
    }
    
    # Perform the search
    search = GoogleSearch(params)
    results = search.get_dict()

    # Check if 'organic_results' exist in results
    if 'organic_results' in results:
        # Convert results to a DataFrame
        df = pd.json_normalize(results['organic_results'])
        
        # Save results to an Excel file
        df.to_excel(f'{file_name}.xlsx', index=False)
        print(f"Results saved to {file_name}.xlsx")
    else:
        print("No organic results found.")

# Function to filter an Excel file by a specific topic and save the filtered results to a new Excel file
def filter_excel_by_topic(input_file, topic, output_file):
    # Read the Excel file
    df = pd.read_excel(input_file)
    
    # Print the columns of the DataFrame to verify which columns are present
    print("Columns in the DataFrame:", df.columns)
    
    # Check if the 'Snippet' column exists and filter by topic
    if 'Snippet' in df.columns:
        filtered_df = df[df['Snippet'].str.contains(topic, case=False, na=False)]
    # If 'Snippet' column does not exist, check for 'title' column (adjust as needed)
    elif 'Description' in df.columns:
        filtered_df = df[df['title'].str.contains(topic, case=False, na=False)]
    else:
        print("Neither 'Snippet' nor 'title' column found.")
        return
    
    # Save the filtered DataFrame to a new Excel file
    filtered_df.to_excel(output_file, index=False)
    print(f"Filtered results saved to {output_file}")

if __name__ == "__main__":
    # Define the search query and the topic to filter by
    query = "define the query"
    topic = "topic"
    
    # Generate a unique file name based on the current date and time
    file_name = f"serpapi_results_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    # Define the file name for the filtered results
    filtered_file_name = "filtered_results.xlsx"

    # Define the number of results and the year range
    num_results = 50
    year_lower_bound = datetime.datetime.now().year - 5
    year_upper_bound = datetime.datetime.now().year

    # Perform the search and save results to an Excel file
    search_and_save_to_excel(API_KEY, query, file_name, num_results, year_lower_bound, year_upper_bound)
    
    # Filter the Excel file by topic and save the filtered results to a new Excel file
    filter_excel_by_topic(f'{file_name}.xlsx', topic, filtered_file_name)
