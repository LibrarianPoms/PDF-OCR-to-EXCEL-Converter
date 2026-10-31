import easyocr
import fitz  # PyMuPDF
import json
import os

# Initialize EasyOCR reader
reader = easyocr.Reader(['fr', 'en'])

# Open PDF
pdf_path = "scan_hanifr_2025-07-10-14-27-01.pdf"
pdf_document = fitz.open(pdf_path)

print(f"Number of pages: {len(pdf_document)}")

all_text = []
for page_num in range(len(pdf_document)):
    print(f"\n--- Page {page_num+1} ---")
    
    # Get the page
    page = pdf_document[page_num]
    
    # Convert to image
    mat = fitz.Matrix(300/72, 300/72)  # 300 DPI
    pix = page.get_pixmap(matrix=mat)
    
    # Save image temporarily
    temp_image = f'temp_page_{page_num}.png'
    pix.save(temp_image)
    
    # Extract text using EasyOCR
    result = reader.readtext(temp_image)
    
    # Print all detected text
    for detection in result:
        bbox, text, confidence = detection
        print(f"Text: {text} | Confidence: {confidence:.2f}")
        all_text.append({
            'text': text,
            'confidence': confidence,
            'page': page_num
        })
    
    # Clean up
    os.remove(temp_image)

pdf_document.close()

# Save all text to file for analysis
with open('extracted_text.json', 'w', encoding='utf-8') as f:
    json.dump(all_text, f, ensure_ascii=False, indent=2)

print(f"\nTotal text elements extracted: {len(all_text)}")
print("Extracted text saved to extracted_text.json")
