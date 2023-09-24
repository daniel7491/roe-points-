import cv2
import pytesseract
import pandas as pd
import re
import os

# Specify the pytesseract executable path based on your project's virtual environment
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Define the folder path containing your images

# Define the folder path containing your images
x = input("input the path to the images: ")
folder_path = x  # Remove the rf and use x directly

# Create an empty list to store the data for all images
all_data = []

# Define the columns you want in your Excel output
columns = ['Name', 'Contribution', 'Ranking', 'Season points', 'Demolition Value in Eden', 'Feat in Eden', 'Occupy Enemy Territory']

# Regular expression patterns to match key-value pairs
patterns = {
    'Name': r'\[21X\] \* (.*) \*',
    'Contribution': r'Contribution (\d+(?:,\d+)*)',
    'Ranking': r'Ranking (\d+(?:,\d+)*)',
    'Season points': r'Season points (\d+(?:,\d+)*)',
    'Demolition Value in Eden': r'Demolition Value in Eden (\d+(?:,\d+)*)',
    'Feat in Eden': r'Feat in Eden (\d+(?:,\d+)*)',
    'Occupy Enemy Territory': r'Occupy Enemy Territory (\d+(?:,\d+)*)',
}

# Iterate through all files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".jpg") or filename.endswith(".png"):
        # Load the image
        image_path = os.path.join(folder_path, filename)
        image = cv2.imread(image_path)

        # Extract text from the image
        extracted_text = pytesseract.image_to_string(image)

        # Initialize a dictionary to store the extracted data for this image
        data_dict = {col: None for col in columns}

        # Extract data using regular expressions
        for key, pattern in patterns.items():
            match = re.search(pattern, extracted_text)
            if match:
                if key == 'Name':
                    data_dict[key] = match.group(1)  # Name doesn't need formatting
                else:
                    data_dict[key] = match.group(1).replace(',', '')  # Remove commas from numbers

        # Append the data for this image to the list
        all_data.append(data_dict)

# Create a DataFrame from the list of data dictionaries
df = pd.DataFrame(all_data)

# Export to Excel
excel_file = 'all_images_data.xlsx'
df.to_excel(excel_file, index=False)
