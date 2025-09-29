
'''col1,col2 is wrong and
   col3 is correct '''

import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import pandas as pd
import re

# Configure
PDF_PATH = "/Users/shubham_working/PycharmProjects/pdf_to_excel_extractor/input_pdfs/56851 Bundles.pdf"
pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"  # Adjust as needed

# Output data
data = []

# Convert PDF to images
images = convert_from_path(PDF_PATH, dpi=300)

for page_index, image in enumerate(images):
    # Rotate image to horizontalize vertical text (rotate counter-clockwise 90°)
    rotated_image = image.rotate(270, expand=True)

    # OCR on rotated image
    custom_config = r'--oem 3 --psm 6'
    rotated_text = pytesseract.image_to_string(rotated_image, config=custom_config)

    # Extract col1
    match_col1 = re.search(r"\*\*\s*(B[\s\-X\d]+of\s\d+)", rotated_text)
    col1 = match_col1.group(1).strip() if match_col1 else f"B ? of ? (Page {page_index+1})"

    # Extract col2 (bundle number)
    match_col2 = re.search(r'\b(\d{5})\b', rotated_text)
    col2 = match_col2.group(1) if match_col2 else ""

    # Extract rotated col3 value like 2-00-08
    match_col3 = re.search(r'\b\d-\d{2}-\d{2}\b', rotated_text)
    col3 = match_col3.group(0) if match_col3 else ""

    data.append((col1, col2, col3))

# Save results
df = pd.DataFrame(data, columns=['col1', 'col2', 'col3'])
print(df)
df.to_excel("output_excels/rotated_col3_fixed.xlsx", index=False)
