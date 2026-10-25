# PDF-OCR-to-EXCEL-Converter
                                                 
# PDF OCR to Excel Converter with PAC Interne Automation

This application extracts tables and information from scanned PDF files (using OCR), generates styled Excel files, and provides automation for generating "PAC Interne" Word documents.

## Features

- Drag & drop PDF upload via web interface
- Automatic OCR extraction (French/English)
- Generates Excel files with extracted data (one or multiple PDFs)
- Adds a "PAC Interne" button in Excel for Word document automation (with VBA instructions)
- Generates PAC Word documents directly from Excel or via the web app

## Requirements

- Windows OS (for Word automation)
- Python 3.8+
- [pip](https://pip.pypa.io/en/stable/installation/)
- Microsoft Excel (for macro-enabled features)
- Microsoft Word (for PAC document automation)
- [poppler-utils](https://github.com/oschwartz10612/poppler-windows/releases/) (for PDF conversion)
- (Optional) [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) if you want to use other OCR engines

## Installation

1. **Clone or Download the Repository**

   ```
   git clone <your-repo-url>
   cd <your-repo-folder>
   ```

2. **Install Python Dependencies**

   ```
   pip install -r requirements.txt
   ```

3. **Install Poppler**

   - Download from: https://github.com/oschwartz10612/poppler-windows/releases/
   - Extract and add the `bin` folder to your Windows PATH.

4. **(Optional) Install Tesseract OCR**

   - Download from: https://github.com/tesseract-ocr/tesseract
   - Add the install directory to your PATH if you want to use it.

## Usage

1. **Start the Application**

   Double-click `run_app.bat` or run:

   ```
   python app.py
   ```

   The app will start on [http://localhost:5001](http://localhost:5001).

2. **Open the Web Interface**

   - Go to [http://localhost:5001](http://localhost:5001) in your browser.

3. **Upload PDFs**

   - Drag and drop your scanned PDF files or use the "Choisir des fichiers" button.
   - Wait for extraction (can take a few seconds per file).

4. **Review & Edit Data**

   - The extracted data will appear in editable tables.
   - You can correct or complete any information before generating Excel.

5. **Generate Excel**

   - Click the "Générer Excel" button.
   - Download the generated Excel file.

6. **PAC Interne Button & Word Automation**

   - The Excel file will have a "PAC Interne" button in each sheet.
   - Click the button to generate a PAC Word document (requires the Flask app to be running).
   - Alternatively, follow the VBA instructions in the hidden sheet or the generated `_VBA_Instructions.txt` file to enable offline PAC generation.

7. **Troubleshooting**

   - If you see errors about missing `poppler` or `tesseract`, ensure they are installed and in your PATH.
   - For Word automation, Microsoft Word must be installed and accessible.

---

For any issues, please contact
