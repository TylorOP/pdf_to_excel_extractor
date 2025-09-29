import os
import re
import pytesseract
import pdfplumber
import pandas as pd
from pdf2image import convert_from_path
from PIL import Image
from datetime import datetime

# Folders
INPUT_FOLDER = "input_pdfs"
OUTPUT_FOLDER = "output_excels"
LOG_FOLDER = "logs"

# Tesseract path
pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"  # Adjust if needed

# Ensure folders exist
os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)


def extract_info_from_page(page, page_index, errors, pdf_name):
    try:
        # Text-based extraction using pdfplumber (for col1 and col2)
        text = page.extract_text()
        if not text:
            errors.append(f"{pdf_name} Page {page_index + 1}: No text found.")
            return ["NOT FOUND", "NOT FOUND", "NOT FOUND"]

        # Extract col1
        match_col1 = re.search(r"\bB[\s\-X\d]+of\s\d+\b", text)
        col1 = match_col1.group(0).strip() if match_col1 else f"B ? of ? (Page {page_index + 1})"

        # Extract col2
        match_col2 = re.search(r'\b\d{5}\b', text)
        col2 = match_col2.group(0) if match_col2 else "NOT FOUND"

        # Convert the page to an image, rotate for OCR
        image = page.to_image(resolution=300).original
        rotated_image = image.rotate(270, expand=True)

        # OCR on rotated image
        custom_config = r'--oem 3 --psm 6'
        ocr_text = pytesseract.image_to_string(rotated_image, config=custom_config)

        # Extract col3: value like 2-00-08
        match_col3 = re.search(r'\b\d-\d{2}-\d{2}\b', ocr_text)
        col3 = match_col3.group(0) if match_col3 else "NOT FOUND"

        print(f"DEBUG Page {page_index + 1}: col1={col1}, col2={col2}, col3={col3}")
        return [col1, col2, col3]

    except Exception as e:
        errors.append(f"{pdf_name} Page {page_index + 1} Exception: {str(e)}")
        return ["ERROR", "ERROR", "ERROR"]


def process_pdf(file_path):
    records = []
    errors = []
    pdf_name = os.path.basename(file_path)

    with pdfplumber.open(file_path) as pdf:
        for i, page in enumerate(pdf.pages):
            col1, col2, col3 = extract_info_from_page(page, i, errors, pdf_name)
            if "NOT FOUND" in (col1, col2, col3):
                errors.append(f"{pdf_name} Page {i + 1}: Missing value(s) - col1: {col1}, col2: {col2}, col3: {col3}")

            records.append({
                "col1": col1,
                "col2": col2,
                "col3": col3
            })

    # Save Excel
    output_file = os.path.join(OUTPUT_FOLDER, pdf_name.replace(".pdf", ".xlsx"))
    pd.DataFrame(records).to_excel(output_file, index=False)

    # Save logs
    if errors:
        log_file = os.path.join(LOG_FOLDER, pdf_name.replace(".pdf", ".log"))
        with open(log_file, "w") as f:
            for err in errors:
                f.write(err + "\n")

    print(f"✅ Processed: {pdf_name}")


def main():
    pdf_files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print("⚠️ No PDF files found in input_pdfs folder.")
        return

    for pdf_file in pdf_files:
        process_pdf(os.path.join(INPUT_FOLDER, pdf_file))

    print(f"🎉 All PDFs processed. Excel files in '{OUTPUT_FOLDER}'.")


if __name__ == "__main__":
    main()
