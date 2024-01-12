import cv2
import pytesseract
import pandas as pd
import re
import os
from tqdm import tqdm

# Specify the pytesseract executable path based on your project's virtual environment
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Input validation for folder path
folder_path = input("Input the path to the images: ")
if not os.path.exists(folder_path):
    print("Error: The specified folder path does not exist.")
    exit()

# Create an empty list to store the data for all images
all_data = []

# Define the regex pattern for name extraction
name_pattern = r'\[CNt\]\s*([^CNt#]+)CNt\s*#'

# Define the regex patterns for other data extraction
data_patterns = {
    'Contribution': r'Contribution\s*(\d+(?:,\d+)*)',
    'Ranking': r'Ranking\s*(\d+)',
    'Season points': r'Season points\s*(\d+(?:,\d+)*)',
    'Demolition Value in Eden': r'Demolition Value in Eden\s*(\d+(?:,\d+)*)',
    'Merit': r'Merit\s*(\d+(?:,\d+)*)',
    'Wonder CP': r'Wonder CP\s*:\s*(\d+(?:,\d+)*)',
    'Occupy Enemy Territory': r'Occupy Enemy Territory\s*(\d+)',
}

# Iterate through all files in the folder with a progress bar
for filename in tqdm(os.listdir(folder_path), desc="Processing images"):
    if filename.lower().endswith((".jpg", ".png")):
        # Load the image
        image_path = os.path.join(folder_path, filename)
        try:
            image = cv2.imread(image_path)
            # Convert to grayscale, apply threshold, etc. for preprocessing if needed
            # ...
            # Extract text with pytesseract
            extracted_text = pytesseract.image_to_string(image)
            # Search for the name using the regex pattern
            name_match = re.search(name_pattern, extracted_text)
            name = name_match.group(1).strip() if name_match else 'Name not found'

            # Initialize data dictionary with name
            data_dict = {'Name': name}

            # Extract other data using their respective regex patterns
            for key, pattern in data_patterns.items():
                match = re.search(pattern, extracted_text, re.IGNORECASE)
                data_dict[key] = int(match.group(1).replace(',', '')) if match else 0

            # Append the data for this image to the list
            all_data.append(data_dict)
        except Exception as e:
            print(f"Error processing image {filename}: {e}")

# Convert to DataFrame
df = pd.DataFrame(all_data)

# Export to Excel
excel_file = 'all_images_data.xlsx'
df.to_excel(excel_file, index=False)
print(f"Data extracted from images has been saved to {excel_file}")
