import os
import re
import time
import pytesseract
import pdfplumber
import pandas as pd
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


# --- Fallback function for col3 if not found initially ---
def extract_col3_with_fallback(page, page_index):
    try:
        # Convert and rotate image
        image = page.to_image(resolution=300).original.convert("RGB")
        rotated = image.rotate(270, expand=True)

        # Crop region of interest
        width, height = rotated.size
        crop_box = (
            int(width * 0.05),
            int(height * 0.20),
            int(width * 0.55),
            int(height * 0.70)
        )
        cropped = rotated.crop(crop_box)

        # OCR
        custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789-'
        ocr_text = pytesseract.image_to_string(cropped, config=custom_config)
        print(f"🔁 Fallback OCR (Page {page_index + 1}): '{ocr_text.strip()}'")

        # 1. Direct match
        match = re.search(r'\b\d-\d{2}-\d{2}\b', ocr_text)
        if match:
            return match.group(0)

        # 2. Rebuild from lines
        lines = [line.strip() for line in ocr_text.splitlines() if line.strip()]
        for i in range(len(lines) - 2):
            f, s, t = lines[i:i+3]
            if re.fullmatch(r'\d-', f) and re.fullmatch(r'\d{2}', s) and re.fullmatch(r'\d{2}', t):
                return f"{f[0]}-{s}-{t}"
            if re.fullmatch(r'\d', f) and re.fullmatch(r'\d{2}', s) and re.fullmatch(r'\d{2}', t):
                return f"{f}-{s}-{t}"

        # 3. Rebuild from characters
        chars = ''.join(c for c in ocr_text if c.isdigit() or c == '-')
        match = re.search(r'\d-\d{2}-\d{2}', chars)
        if match:
            parts = match.group(0).split("-")
            if all(p.isdigit() for p in parts) and int(parts[1]) <= 31 and int(parts[2]) <= 31:
                return match.group(0)


        return "NOT FOUND"
    except Exception as e:
        print(f"⚠️ Fallback OCR failed on Page {page_index + 1}: {str(e)}")
        return "ERROR"


# --- Main extraction function per page ---
def extract_info_from_page(page, page_index, errors, pdf_name):
    try:
        # Text extraction (col1, col2)
        text = page.extract_text()
        if not text:
            errors.append(f"{pdf_name} Page {page_index + 1}: No text found.")
            return ["NOT FOUND", "NOT FOUND", "NOT FOUND"]

        # col1 (bundle number)
        match_col1 = re.search(r"\bB[\s\-X\d]+of\s\d+\b", text)
        col1 = match_col1.group(0).strip() if match_col1 else f"B ? of ? (Page {page_index + 1})"

        # col2 (5-digit number)
        match_col2 = re.search(r'\b\d{5}\b', text)
        col2 = match_col2.group(0) if match_col2 else "NOT FOUND"

        # OCR for col3
        image = page.to_image(resolution=300).original
        rotated_image = image.rotate(270, expand=True)
        custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789-'
        ocr_text = pytesseract.image_to_string(rotated_image, config=custom_config)
        match_col3 = re.search(r'\b\d-\d{2}-\d{2}\b', ocr_text)
        col3 = match_col3.group(0) if match_col3 else "NOT FOUND"

        print(f"DEBUG Page {page_index + 1}: col1={col1}, col2={col2}, col3={col3}")
        return [col1, col2, col3]

    except Exception as e:
        errors.append(f"{pdf_name} Page {page_index + 1} Exception: {str(e)}")
        return ["ERROR", "ERROR", "ERROR"]


# --- Process PDF file ---
def process_pdf(file_path):
    records = []
    errors = []
    pdf_name = os.path.basename(file_path)

    with pdfplumber.open(file_path) as pdf:
        for i, page in enumerate(pdf.pages):
            col1, col2, col3 = extract_info_from_page(page, i, errors, pdf_name)

            # Try fallback if col3 failed
            if col3 == "NOT FOUND":
                col3_fallback = extract_col3_with_fallback(page, i)
                if col3_fallback != "NOT FOUND":
                    col3 = col3_fallback
                else:
                    errors.append(f"{pdf_name} Page {i + 1}: col3 still NOT FOUND after fallback")

            records.append({
                "col1": col1,
                "col2": col2,
                "col3": col3
            })

            time.sleep(1)  # Optional small delay per page

    # Save Excel
    output_file = os.path.join(OUTPUT_FOLDER, pdf_name.replace(".pdf", ".xlsx"))
    pd.DataFrame(records).to_excel(output_file, index=False)

    # Save error log
    if errors:
        log_file = os.path.join(LOG_FOLDER, pdf_name.replace(".pdf", ".log"))
        with open(log_file, "w") as f:
            for err in errors:
                f.write(err + "\n")

    print(f"✅ Processed: {pdf_name}")


# --- Main entry point ---
def main():
    pdf_files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print("⚠️ No PDF files found in input_pdfs folder.")
        return

    for pdf_file in pdf_files:
        process_pdf(os.path.join(INPUT_FOLDER, pdf_file))

    print(f"🎉 All PDFs processed. Excel files saved in '{OUTPUT_FOLDER}'.")


if __name__ == "__main__":
    main()
