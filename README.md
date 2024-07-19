# Google Scholar Article Search and Filter

## Description

This project provides a tool that automates the search for academic articles on Google Scholar using the SerpApi and filters results based on specific keywords and topics. The application features an intuitive graphical user interface (GUI) built with `tkinter`, allowing users to input their API key, define search queries, and specify filtering criteria. Results are saved to Excel files and can be further filtered by examining the content of linked web pages.

## Tools Used

- **Python:** Primary programming language.
- **pandas:** Data manipulation and analysis.
- **openpyxl:** Working with Excel files.
- **requests:** Making HTTP requests to extract web page content.
- **BeautifulSoup:** Parsing HTML content to extract text and search for keywords.
- **serpapi:** API for accessing Google Scholar search results.
- **tkinter:** Creating the graphical user interface (GUI).

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/MparrillaL/google-scholar-search-and-filter.git
Navigate to the project directory:

sh
Copiar código
cd google-scholar-search-and-filter
Create and activate a virtual environment (optional but recommended):

sh
Copiar código
python -m venv venv
source venv/bin/activate  # On Windows use `venv\Scripts\activate`
Install dependencies:

sh
Copiar código
pip install -r requirements.txt
Usage
Run the main script:

sh
Copiar código
python app.py
In the graphical interface, enter your SerpApi key, search query, topic to filter by, keywords to search in the web pages, and desired year ranges.

Click "Search and Filter" to get the results.

Results will be saved to Excel files. The final file will be filtered based on the content of the linked web pages.

Requirements
Python 3.6 or higher
Dependencies listed in requirements.txt
