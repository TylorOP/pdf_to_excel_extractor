# PDF to Excel Extractor

This project extracts data from PDFs and saves it into Excel files.

---

## Requirements
- Python 3.10+ installed
- Git (optional, if cloning from GitHub)
- `pip` for installing Python packages

---

## Steps to Run

### 1. Clone the repository
Open Terminal (or Command Prompt on Windows) and run:

```bash```
git clone git@github.com:TylorOP/pdf_to_excel_extractor.git  
cd pdf_to_excel_extractor  


### 2. Create a virtual environment  (Optional but recommended)
```bash```
python3 -m venv venv


source venv/bin/activate       # macOS/Linux  
venv\Scripts\activate       # On Windows  

### 3. Install dependencies
All required Python packages can be installed with:

```bash```
pip install -r requirements.txt

### 4.Install Tesseract OCR engine:

`pytesseract` needs the Tesseract OCR program installed on your machine. Follow your OS below:

#### **macOS**
1. Make sure [Homebrew](https://brew.sh/) is installed.    
2. Open Terminal and run:

brew install tesseract

Test installation:
Next run this :

tesseract --version


You should see a version number like tesseract 5.x.x.

Windows: Download from Tesseract OCR and add to PATH

Download the Tesseract installer from Tesseract at UB Mannheim 
Link--> " https://github.com/UB-Mannheim/tesseract/wiki "

Click the latest exe file under “Windows Installer” (for e.g. tesseract-ocr-w64-setup-5.5.0.20241111.exe ). 

Run the installer and follow the instructions. Keep the default folder (usually C:\Program Files\Tesseract-OCR).

Add Tesseract to your system PATH:
Press Win + R, type sysdm.cpl → Enter
Go to Advanced → Environment Variables → Path → Edit → New
Add C:\Program Files\Tesseract-OCR → Click OK

Test installation in Command Prompt:
tesseract --version
You should see a version number printed.

### 5. Place your PDFs in the input_pdfs/ folder.

### 6. Run the main script

```bash```
python final_code.py

The script will process PDFs and save Excel files in output_excels/.
Logs will be saved in the logs/ folder.

### 7. Check output
Processed Excel files appear in output_excels/.
Check logs/ for any errors or skipped pages.


Notes
Make sure Tesseract OCR is installed if pytesseract is used:
macOS (Homebrew): brew install tesseract
Windows: download installer from Tesseract OCR and add to PATH
