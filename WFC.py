import os
import re
from collections import Counter
import pandas as pd
import fitz  # PyMuPDF

def extract_text_from_pdf(pdf_path):
    """
    Extract text from a PDF file using PyMuPDF (fitz).

    Args:
    pdf_path (str): Path to the PDF file.

    Returns:
    str: Extracted text from the PDF.
    """
    text = ''
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text += page.get_text()
    return text

def count_word_frequency(text, keywords):
    """
    Count word frequency for exact matches of keywords in text.
    
    Args:
    text (str): Text to count word frequency from.
    keywords (list): List of keywords to search for.
    
    Returns:
    dict: Dictionary containing keyword frequencies.
    """
    word_freq = Counter()
    for keyword in keywords:
        pattern = r'\b{}\b'.format(re.escape(keyword))
        matches = re.findall(pattern, text.lower())
        word_freq[keyword] = len(matches)
    return word_freq

def process_folder(folder_path):
    """
    Process all PDF files in a folder.
    
    Args:
    folder_path (str): Path to the folder containing PDF files.
    """
    theme_name = input("Enter the name of the theme: ")
    fiscal_year = input("Enter the fiscal year: ")
    search_query = input("Enter keywords separated by comma: ").lower()
    search_queries = [query.strip() for query in search_query.split(',')]  # Split input by comma
    
    # Initialize lists to store data for DataFrame
    rows = []
    
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            text = extract_text_from_pdf(pdf_path)
            print(f'File: {filename}')
            word_freq = count_word_frequency(text, search_queries)
            for keyword, frequency in word_freq.items():
                print(f"The frequency of '{keyword}' is: {frequency}")
            print()
            row = [theme_name, fiscal_year, filename] + [word_freq[keyword] for keyword in search_queries]
            row.append(sum(word_freq.values()))  # Append total frequency to the row
            rows.append(row)
    
    # Create DataFrame
    columns = ['Theme', 'Year', 'File Name'] + search_queries + ['Keyword total']
    df = pd.DataFrame(rows, columns=columns)
    
    # Create output folder if it doesn't exist
    output_folder = os.path.join(os.path.dirname(__file__), 'output')
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Define Excel file path
    excel_file = os.path.join(output_folder, f"{theme_name}_{fiscal_year}_word_frequency_report.xlsx")
    
    # Export DataFrame to Excel file
    df.to_excel(excel_file, index=False)
    print(f"Excel file '{excel_file}' generated successfully!")

def main():
    folder_path = 'pdf_files'  # Relative path to the folder containing PDF files
    folder_path = os.path.join(os.path.dirname(__file__), folder_path)  # Get absolute path
    process_folder(folder_path)

if __name__ == "__main__":
    main()