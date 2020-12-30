# Import Dependencies
import pytesseract
import argparse
import cv2
import os
import numpy as np
import pandas as pd
import time
from pdf2image import convert_from_path
from PIL import Image

# Set pytesseract path on local machine
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Define Individual Functions Per File Type
def df_to_txt(df):
    # String to append to, specify source type
    text = "Pandas DataFrame: \n"
    
    # Get list of columns
    df_cols = df.columns.tolist()
    
    # Extract data from each column
    for item in df_cols:
        text = text + ' '.join(df[item].tolist())
        text = text + '\n'
        time.sleep(2)
    return text

def csv_to_txt(path):
    # Initialize string
    doc_text = ''

    # Read in csv
    df_csv = pd.read_csv(path)

    # Extract text
    df_csv = df_csv.applymap(str).copy()
    doc_text = df_to_txt(df_csv)
    
    return doc_text

def excel_to_txt(path):
    # Initialize string
    doc_text = ''
    
    # Read in csv
    df_excel = pd.read_excel(path)
    
    # Extract text
    df_excel = df_excel.applymap(str).copy()
    doc_text = df_to_txt(df_excel)
    
    return doc_text

def ocr_pdf(path):
    # Initialize string to append to and get pages
    doc_text = ''
    pages = convert_from_path(path)
    
    # Initialize page counter, and loop through pages
    page_counter = 1
    
    for page in pages:   
        # Create image with unique page name
        filename = "page_" + str(page_counter) + ".jpg"
        page.save(filename,"JPEG")
        page_counter = page_counter + 1
        
        # Read text from image and discard image
        text = pytesseract.image_to_string(Image.open(filename))
        os.remove(filename)
        
        # Concatenate text to one string
        if doc_text == '':
            doc_text = text
        else:
            doc_text = doc_text + "Page Number: " + str(page_counter) + '\n' + text
            
    # Return text result
    return doc_text

def docs_to_txt(path):
    # Go into grant folder
    os.chdir(path)

    # Get list of PDF documents
    doc_files = os.listdir()
    grant_docs = []

    # Go through documents 
    for doc in doc_files:
        # Initiialize doc
        doc_text = ""
        
        # Handle different file type cases
        if doc.endswith('.pdf'):
            doc_text = ocr_pdf(doc)
        elif doc.endswith('.csv'):
            doc_text = csv_to_txt(doc)
        elif doc.endswith('.xlsx'):
            doc_text = excel_to_txt(doc)
        elif doc.endswith('.xls'):
            doc_text = excel_to_txt(doc)
        else:
            extension = doc.split('.')[1]
            doc_text = 'File Format ' + extension + " not currently supported."
        
        # Get file name
        print('\n***** DOC PROCESSED: ')
        file_name = doc.split('.')[0]
        print(doc + ": " + doc_text[:100])
        
        # Write two new text file, with same name
        f = open(f'{file_name}.txt',"w+",encoding="utf-16",errors='ignore')
        f.write(doc_text)
        f.close()
        
        # Append to list of strings/docs
        grant_docs.append(doc_text)
        
        time.sleep(2)
        
    # Leave grant folder
    os.chdir('..')

def ocr_directories(path):
    # Got into parent directory
    os.chdir(path)
    
    # Get list of grants/directories
    grant_folders = os.listdir()
    content = []
    
    # Go through each grant directory
    for g_path in grant_folders:
        # Get unique grant id
        grant_id = g_path
        
        # Create .txt versions of each doc in dir
        grant_docs = docs_to_txt(g_path)
        content.append(grant_docs)
        print('************** GRANT PROCESSED')
    
    # grant_dict = {'Grant_ID':grant_folders, 'Content':content}
    # df = pd.DataFrame(grant_dict)
    os.chdir('..')
    # return(df)

def clear_jpgs(path):
    # Clear leftover jpg files
    os.chdir(path)

    grant_folders = os.listdir()
    for g_path in grant_folders:
        os.chdir(g_path)
        files_in_directory = os.listdir()
        filtered_files = [file for file in files_in_directory if file.endswith(".jpg")]
        for file in filtered_files:
            os.remove(file)
        os.chdir('..')
    os.chdir('..')

def clear_txt(path):
    # Clear leftover txt files
    os.chdir(path)

    grant_folders = os.listdir()
    for g_path in grant_folders:
        os.chdir(g_path)
        files_in_directory = os.listdir()
        filtered_files = [file for file in files_in_directory if file.endswith(".txt")]
        for file in filtered_files:
            os.remove(file)
        os.chdir('..')
    os.chdir('..')

# Clear residual files for new run
clear_jpgs('proposal_docs_for_ocr')
clear_txt('proposal_docs_for_ocr')

# Start new run
ocr_directories('proposal_docs_for_ocr')

# ***** Dataframe export functionality commented out,
# ***** Uncommented in functions above to reactivate

# ***** Add count of documents per grant id

# doc_counts = []
# for e,row in df.iterrows():
#     doc_counts.append(len(row.Content))
# df[Doc_count] = doc_counts
# df.head()

# ***** Export df
# df.to_csv('grant_content.csv')