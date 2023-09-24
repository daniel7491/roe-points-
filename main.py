import cv2
import pytesseract
import pandas as pd
import re

# Specify the pytesseract executable path based on your project's virtual environment
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Rest of your code remains the same
image_path = r"C:\Users\dnai7\Desktop\1.jpg"
image = cv2.imread(image_path)
extracted_text = pytesseract.image_to_string(image)

# Sample extracted data (replace with your actual extracted data)
extracted_data = [
    'Name: Jabes',
    'Contribution: 104,015',
    'Ranking: 8',
    'Season points: 58,792 58,792',
    'Declaration Point: 0 0',
    'Stone Donation: 0 0',
    'Total War Joined ie): 0',
    'Total War Demolition Value :0 ie,)',
    'Total War Feat ie): 0',
    'Demolition Value in Eden: 552,086 27,604',
    'Feat in Eden: 1,624,916,276 16,249',
    'Guild Contribution EXP (e) ie)',
    'Wonder CP: 0 0',
    'Occupy Enemy Territory: 1,370 0',
]

# Define column names
columns = ['Name', 'Contribution', 'Ranking','Season points','Demolition Value in Eden','Feat in Eden','Occupy Enemy Territory']  # Add more columns as needed

# Initialize a dictionary to store data
data_dict = {column: '' for column in columns}

# Process the extracted data and populate the dictionary
for line in extracted_data:
    for column in columns:
        if line.startswith(column + ':'):
            data_dict[column] = line.split(': ')[1]

# Create a DataFrame from the dictionary
df = pd.DataFrame([data_dict])

# Export to Excel
excel_file = '1.xlsx'
df.to_excel(excel_file, index=False)

# # Export to Excel
# excel_file = 'output.xlsx'
# data = {'Extracted Data': [extracted_text]}
# df = pd.DataFrame(data)
# df.to_excel(excel_file, index=False)
