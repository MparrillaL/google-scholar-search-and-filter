import os
import pandas as pd
from serpapi import GoogleSearch
from datetime import datetime

# Configure your SerpAPI key
API_KEY = 'your_serpapi_api_key'

# Function to perform the search using SerpAPI and save results to Excel
def search_and_save_excel(query, file_name):
    # Set up the search parameters
    search = GoogleSearch({
        'q': query,
        'api_key': API_KEY,
        'output': 'json'
    })

    # Perform the search and get results as a dictionary
    results = search.get_dict()

    # Extract relevant data from the search results
    data = []
    for result in results.get('organic_results', []):
        title = result.get('title')
        link = result.get('link')
        snippet = result.get('snippet')
        data.append({'Title': title, 'Link': link, 'Snippet': snippet})

    # Convert the data to a Pandas DataFrame
    df = pd.DataFrame(data)

    # Save the DataFrame to an Excel file
    excel_file = f'{file_name}.xlsx'
    df.to_excel(excel_file, index=False)
    print(f'Results saved to {excel_file}')

# Function to load the Excel file, filter data by topic, and save filtered results
def filter_excel_by_topic(file_name, topic, filtered_file_name):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_name)

    # Filter the DataFrame by the topic
    filtered_df = df[df['Snippet'].str.contains(topic, case=False, na=False)]

    # Save the filtered DataFrame to a new Excel file
    filtered_excel_file = f'{filtered_file_name}.xlsx'
    filtered_df.to_excel(filtered_excel_file, index=False)
    print(f'Filtered results saved to {filtered_excel_file}')

if __name__ == '__main__':
    # Define your search query and file name
    query = 'Machine Learning applications in healthcare'
    file_name = f'serpapi_results_{datetime.now().strftime("%Y%m%d_%H%M%S")}'

    # Call the function to search and save to Excel
    search_and_save_excel(query, file_name)

    # Define the topic to filter by and the filtered file name
    topic = 'healthcare'
    filtered_file_name = f'filtered_results_{datetime.now().strftime("%Y%m%d_%H%M%S")}'

    # Call the function to load, filter, and save the filtered Excel
    filter_excel_by_topic(f'{file_name}.xlsx', topic, filtered_file_name)
