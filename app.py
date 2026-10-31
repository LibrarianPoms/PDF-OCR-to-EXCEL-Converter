from flask import Flask, render_template, request, send_file, jsonify
import os
import easyocr
import pandas as pd
import fitz  # PyMuPDF
import re
from datetime import datetime
import tempfile
import shutil
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_COLOR_INDEX
import subprocess

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create upload folder if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('temp', exist_ok=True)

# Initialize EasyOCR reader (French and English languages)
reader = easyocr.Reader(['fr', 'en'])

def cleanup_artifacts(text):
    """Remove unwanted artifacts from text."""
    known_artifacts = ['0ynn29', '1085fsi', '42G98', '{888890', '21(89', '3098op', '@e;', 'moqEPI', 
                      'Soô@', '4@', '18f20o58', 'moqEPI']
    clean_words = []
    for word in text.split():
        # Skip known artifacts
        if word in known_artifacts:
            continue
        # Skip very short numbers (1-2 digits) unless they're part of measurements
        if re.match(r'^\d{1,2}$', word) and word not in ['3m', '9m']:
            continue
        # Skip mixed alphanumeric codes that are clearly artifacts
        if word in ['0ynn29', '1085fsi', '42G98', '{888890']:
            continue
        # Keep meaningful words even if they might look like codes
        clean_words.append(word)
    return ' '.join(clean_words).strip()

def extract_text_from_pdf(pdf_path):
    """Convert PDF to images and extract text using EasyOCR"""
    extracted_data = {
        'fournisseur': '',
        'numero_commande': '',
        'numero_contrat': '',
        'date_livraison': '',
        'articles': []
    }
    
    try:
        # Open PDF with PyMuPDF
        pdf_document = fitz.open(pdf_path)
        
        all_text = []
        for page_num in range(len(pdf_document)):
            # Get the page
            page = pdf_document[page_num]
            
            # Convert to image
            mat = fitz.Matrix(300/72, 300/72)  # 300 DPI
            pix = page.get_pixmap(matrix=mat)
            
            # Save image temporarily
            temp_image = os.path.join('temp', f'page_{page_num}.png')
            pix.save(temp_image)
            
            # Extract text using EasyOCR
            result = reader.readtext(temp_image)
            
            # Combine all text with positions
            for detection in result:
                bbox, text, confidence = detection
                all_text.append({
                    'text': text,
                    'bbox': bbox,
                    'confidence': confidence,
                    'page': page_num
                })
            
            # Clean up temp image
            os.remove(temp_image)
        
        pdf_document.close()
        
        # Parse extracted text
        extracted_data = parse_extracted_text(all_text)
        
    except Exception as e:
        print(f"Error extracting text: {str(e)}")
    
    return extracted_data

def is_meaningful_text(text):
    """Determine if a text string is meaningful content vs OCR artifacts"""
    if not text or len(text.strip()) == 0:
        return False
    
    text = text.strip()
    
    # Aggressive filtering of known OCR artifacts
    artifacts = [
        'Soô@', '4@', '0ynn29', '1085fsi', '42G98', '{888890', '21(89', '3098op', '@e;',
        'moqEPI', 'soaHarxesom', 'S.A', 'IE', 'RC', 'CNSS', 'ICE', 'H.T_', 'H.T', 'MAD',
        'Tél', 'Fax', 'Page', 'Pos', 'Uté'
    ]
    if text in artifacts:
        return False
    
    # Allow pure numbers if they are part of meaningful designations
    if re.match(r'^\d+$', text) and len(text) > 3:  # Ensure at least a reasonable length
        return True
    
    # Skip space-separated numbers like "26", "04 0"
    if re.match(r'^\d+\s+\d*$', text):
        return False
    
    # Skip obvious OCR artifacts - mixed alphanumeric codes
    if re.match(r'^[0-9]+[a-zA-Z]+[0-9]*$', text):  # Like "0ynn29", "1085fsi", "42G98"
        return False
    
    # Skip codes that mix letters and numbers without clear word structure
    if re.match(r'^[a-zA-Z0-9]+$', text) and len(text) < 10:
        # Check if it has a reasonable letter/number distribution
        letters = sum(1 for c in text if c.isalpha())
        numbers = sum(1 for c in text if c.isdigit())
        if letters > 0 and numbers > 0 and letters < 3:  # Short letter sequences with numbers
            return False
    
    # Skip bracketed artifacts
    if re.match(r'^[{\[].*[}\]]$', text) or re.match(r'^[{\[][0-9]+$', text) or re.match(r'^[0-9]+[}\]]$', text):
        return False
    
    # Skip symbol-heavy artifacts
    
    # Skip parentheses with numbers only
    if re.match(r'^[(][0-9\s,]+[)]$', text):  # Like "(6 , 10"
        return False
    
    # Skip lines starting with symbols
    if re.match(r'^[@;]+.*$', text):
        return False
    
    # Skip price-like codes
    if re.match(r'^[0-9]+[(][0-9]+$', text):  # Like "21(89"
        return False
    
    # Keep meaningful French words and technical terms
    meaningful_patterns = [
        r'^[a-zA-ZÀ-ÿ\s]+$',  # Pure French words/phrases
        r'\d+\s*m\s*à\s*\d+\s*m',   # Technical measurements like "3m à 9m"
        r'^[A-Z][a-zA-ZÀ-ÿ]+',  # Capitalized words
        r'^[a-zA-ZÀ-ÿ]+$',  # Any alphabetic content
        r'^\d+[a-zA-Z]+$',  # Numbers followed by letters (like "3m")
        r'^[a-zA-ZÀ-ÿ]+\d+$',  # Letters followed by numbers
    ]
    
    for pattern in meaningful_patterns:
        if re.match(pattern, text):
            return True
    
    # Allow anything that contains meaningful text even if it has some special chars
    if any(c.isalpha() for c in text):
        return True
    
    return False

def parse_extracted_text(text_data):
    """Parse the extracted text to find specific fields"""
    extracted = {
        'fournisseur': '',
        'numero_commande': '',
        'numero_contrat': '',
        'date_livraison': '',
        'articles': []
    }
    
    # Convert to simple text list for easier searching
    text_lines = [item['text'] for item in text_data]
    
    # Debug: Print all extracted text
    print("DEBUG: All extracted text:")
    for i, text in enumerate(text_lines):
        print(f"{i}: {text}")
    
    # Extract Fournisseur - try QIT Solutions first, then fallback to general extraction
    # Check for special case where 'Fournisseur' is in line 1 and the name is in lines 7-9
    if len(text_lines) > 9 and text_lines[1].lower() == 'fournisseur':
        # Try to combine lines 7, 8, and 9 (or just 7 and 8 if 9 is not part of the name)
        line7 = text_lines[7].strip() if len(text_lines) > 7 else ''
        line8 = text_lines[8].strip() if len(text_lines) > 8 else ''
        line9 = text_lines[9].strip() if len(text_lines) > 9 else ''
        
        # Check if these lines form a company name (like "BTS com SARL")
        if line7 and line8:
            # Check if line9 looks like part of company name (common suffixes)
            if line9.upper() in ['SARL', 'SA', 'SAS', 'EURL', 'SNC', 'SCA', 'SCS'] or 'SARL' in line9:
                extracted['fournisseur'] = f"{line7} {line8} {line9}"
                print(f"DEBUG: Found fournisseur (lines 7-8-9): {extracted['fournisseur']}")
            else:
                # Just use lines 7 and 8
                extracted['fournisseur'] = f"{line7} {line8}"
                print(f"DEBUG: Found fournisseur (lines 7-8): {extracted['fournisseur']}")
    
    # If not found yet, use the original search method
    if not extracted['fournisseur']:
        for i, text in enumerate(text_lines):
            if 'fournisseur' in text.lower():
                # Look for company name in the lines following "Fournisseur"
                fournisseur_lines = []
                for j in range(i+1, min(i+10, len(text_lines))):
                    current_text = text_lines[j].strip()
                    # Skip empty lines and common non-company words
                    if (current_text and len(current_text) > 1 and 
                        current_text.lower() not in ['ordre', 'sur', 'contrat', 'n', 'n°', 'commande', 'date', '/'] and
                        not re.match(r'^\d+$', current_text) and  # Skip pure numbers
                        not re.match(r'^\d{1,2}[./]\d{1,2}[./]\d{4}$', current_text)):  # Skip dates
                        
                        # Check if this might be part of company name
                        # Include common company suffixes and general text that could be company names
                        if (current_text.upper() in ['SARL', 'SA', 'SAS', 'EURL', 'SNC', 'SCA', 'SCS', 'LTD', 'LLC', 'INC'] or
                            'SARL' in current_text or 'Solutions' in current_text or
                            not current_text[0].islower()):  # Company names usually start with capital
                            fournisseur_lines.append(current_text)
                        else:
                            # If we already have some lines and hit a lowercase word, might be end of company name
                            if fournisseur_lines and current_text[0].islower():
                                break
                            # Otherwise include it
                            fournisseur_lines.append(current_text)
                    else:
                        # Stop if we hit a clear boundary
                        if fournisseur_lines:  # Only stop if we already found something
                            break
                
                if fournisseur_lines:
                    extracted['fournisseur'] = ' '.join(fournisseur_lines)
                    print(f"DEBUG: Found fournisseur: {' '.join(fournisseur_lines)}")
                    break
    
    # Extract Numero de commande - look for numbers near "commande"
    for i, text in enumerate(text_lines):
        if 'commande' in text.lower() and 'contrat' not in text.lower():
            # Look in same line and nearby lines for various number patterns
            for j in range(max(0, i-2), min(i+8, len(text_lines))):
                # Try different number patterns (8-12 digits)
                matches = re.findall(r'\b(\d{8,12})\b', text_lines[j])
                for match in matches:
                    if len(match) >= 8:  # At least 8 digits
                        extracted['numero_commande'] = match
                        print(f"DEBUG: Found numero_commande: {match} near line {i}: {text}")
                        break
                if extracted['numero_commande']:
                    break
            if extracted['numero_commande']:
                break
    
    # If not found, try broader search for "N°" or "No" patterns
    if not extracted['numero_commande']:
        for i, text in enumerate(text_lines):
            if re.search(r'n[°o]?\s*commande', text.lower()) or re.search(r'commande\s*n[°o]?', text.lower()):
                # Look around this line for numbers
                for j in range(max(0, i-1), min(i+5, len(text_lines))):
                    matches = re.findall(r'\b(\d{8,12})\b', text_lines[j])
                    for match in matches:
                        extracted['numero_commande'] = match
                        print(f"DEBUG: Found numero_commande (pattern 2): {match}")
                        break
                if extracted['numero_commande']:
                    break
    
    # Extract Numero de contrat - look for number near "Contrat"
    # First try to find "Contrat:" pattern (could be split across lines)
    for i, text in enumerate(text_lines):
        if 'contrat:' in text.lower() or ('contrat' in text.lower() and ':' in text):
            # Look in the same line and next few lines for number
            for j in range(i, min(i+5, len(text_lines))):
                match = re.search(r'\b(\d{10})\b', text_lines[j])
                if match and match.group(1) != extracted['numero_commande']:
                    extracted['numero_contrat'] = match.group(1)
                    print(f"DEBUG: Found numero_contrat: {match.group(1)} near line {i}: {text}")
                    break
            if extracted['numero_contrat']:
                break
    
    # If not found, look for "N" followed by "Contrat:" pattern (split across lines)
    if not extracted['numero_contrat']:
        for i, text in enumerate(text_lines):
            if text.strip().upper() in ['N', 'N°', 'N°'] and i + 1 < len(text_lines):
                next_line = text_lines[i + 1].strip().lower()
                if 'contrat' in next_line:
                    # Found "N" followed by "Contrat:", look for number in next few lines
                    for j in range(i + 1, min(i + 5, len(text_lines))):
                        match = re.search(r'\b(\d{10})\b', text_lines[j])
                        if match and match.group(1) != extracted['numero_commande']:
                            extracted['numero_contrat'] = match.group(1)
                            break
                    if extracted['numero_contrat']:
                        break
    
    # If still not found, try broader search for "contrat" keyword
    if not extracted['numero_contrat']:
        for i, text in enumerate(text_lines):
            if 'contrat' in text.lower():
                # Look in nearby lines
                for j in range(max(0, i-1), min(i+5, len(text_lines))):
                    match = re.search(r'\b(\d{10})\b', text_lines[j])
                    if match and match.group(1) != extracted['numero_commande']:
                        extracted['numero_contrat'] = match.group(1)
                        break
                if extracted['numero_contrat']:
                    break
    
    # Extract Date de livraison - look for date patterns
    date_patterns = [
        r'\b(\d{1,2}\s*[./]\s*\d{1,2}\s*[./]\s*\d{4})\b',  # DD/MM/YYYY or DD.MM.YYYY
        r'\b(\d{4}\s*[./]\s*\d{1,2}\s*[./]\s*\d{1,2})\b',  # YYYY/MM/DD or YYYY.MM.DD
        r'\b(\d{1,2}\s*[-]\s*\d{1,2}\s*[-]\s*\d{4})\b',   # DD-MM-YYYY
    ]
    
    for i, text in enumerate(text_lines):
        if 'livraison' in text.lower() or 'delivery' in text.lower():
            # Look for date in nearby lines
            for j in range(max(0, i-2), min(i+8, len(text_lines))):
                for pattern in date_patterns:
                    match = re.search(pattern, text_lines[j])
                    if match:
                        extracted['date_livraison'] = match.group(1).replace(' ', '')
                        print(f"DEBUG: Found date_livraison: {match.group(1)}")
                        break
                if extracted['date_livraison']:
                    break
            if extracted['date_livraison']:
                break
    
    # If not found near "livraison", try general date search
    if not extracted['date_livraison']:
        for i, text in enumerate(text_lines):
            for pattern in date_patterns:
                match = re.search(pattern, text)
                if match:
                    # Make sure it's not just any random date, should be reasonable
                    date_str = match.group(1)
                    if '202' in date_str:  # Likely a valid recent date
                        extracted['date_livraison'] = date_str.replace(' ', '')
                        print(f"DEBUG: Found date_livraison (general search): {date_str}")
                        break
            if extracted['date_livraison']:
                break
    
    # Extract Articles table - look for specific article patterns
    articles = []
    
    # Find table start - look for various table indicators
    table_start = -1
    table_indicators = ['pos', 'position', 'item', 'article', 'désignation', 'description', 'quantité', 'quantity']
    
    for i, text in enumerate(text_lines):
        text_lower = text.lower().strip()
        if text_lower in table_indicators:
            table_start = i
            print(f"DEBUG: Found table start at line {i}: {text}")
            break
    
    # Alternative: look for pattern like "010", "020" etc. which indicates start of items
    if table_start == -1:
        for i, text in enumerate(text_lines):
            if re.match(r'^0[1-9]0$', text.strip()):  # Matches 010, 020, 030, etc.
                table_start = max(0, i-5)  # Start a few lines before the first item
                print(f"DEBUG: Found table start by item pattern at line {i}: {text}")
                break
    
    if table_start >= 0:
        # Define known unit types
        known_units = ['UNI', 'FOR']
        
        # Dynamically look for position numbers (e.g., 010, 020, etc.) using a regex pattern
        # Position numbers typically follow patterns like 010, 020, 030... or could be any 3-digit starting with 0
        pos_pattern = re.compile(r'^0\d{2}$')  # Matches 3-digit numbers starting with 0
        found_positions = []
        
        # Scan for all position numbers in the document
        # Only look in a reasonable range after table start to avoid footer numbers
        for i in range(table_start, min(table_start + 100, len(text_lines))):
            line = text_lines[i].strip()
            
            # Direct match for properly OCR'd position numbers
            if pos_pattern.match(line):
                found_positions.append(line)
            # Handle OCR errors where numbers might be split (e.g., "04 0" instead of "040")
            elif re.match(r'^0\d\s+0$', line):  # Matches patterns like "01 0", "02 0", etc.
                # Reconstruct the position number by removing spaces
                fixed_pos = line.replace(' ', '')
                if pos_pattern.match(fixed_pos):
                    found_positions.append(fixed_pos)
                    print(f"DEBUG: Fixed OCR split number '{line}' to '{fixed_pos}'")
        
        # Remove duplicates and sort
        found_positions = sorted(list(set(found_positions)))
        print(f"DEBUG: Looking for position numbers: {found_positions}")
        
        for pos in found_positions:
            print(f"DEBUG: Parsing designation for position: {pos}")  # Added
            # Find this position in the text
            for i in range(table_start, len(text_lines)):
                line = text_lines[i].strip()
                
                # Check for exact match or OCR split version
                if line == pos or (re.match(r'^0\d\s+0$', line) and line.replace(' ', '') == pos):
                    article = {
                        'pos_article': pos,  # Combine pos and article in same field
                        'designation': '',
                        'quantite': '',
                        'unite': '',
                        'prix_unitaire': '',
                        'prix_total': ''
                    }
                    
                    # Build designation from the lines following the position number
                    raw_parts = []
                    j = i + 1
                    quantity_found = False
                    
                    # First, collect raw text between position numbers
                    while j < len(text_lines) and j < i + 20:
                        current_line = text_lines[j].strip()
                        
                        # Stop if we hit another position number
                        if current_line in found_positions:
                            break
                        
                        # Also check for OCR split position numbers
                        if re.match(r'^0\d\s+0$', current_line) and current_line.replace(' ', '') in found_positions:
                            break
                            
                        # Stop collecting if we hit footer text
                        if current_line.lower() in ['montant', 'h.t', 'mad'] or 'montant' in current_line.lower():
                            break
                        
                        # Add to raw parts
                        if current_line:
                            raw_parts.append(current_line)
                        
                        j += 1
                    
                    print(f"DEBUG: Raw parts for position {pos}: {raw_parts}")
                    
# Now post-process to improve accuracy after initial designation extraction
                    designation_parts = []
                    skip_indices = set()
                    
                    # First pass: find quantity and unit
                    for idx, part in enumerate(raw_parts):
                        if idx in skip_indices:
                            continue
                            
                        # Check for quantity + unit pattern
                        if re.match(r'^\d+$', part) and not quantity_found:
                            # Check if next part is a unit
                            if idx + 1 < len(raw_parts) and raw_parts[idx + 1].upper() in known_units:
                                article['quantite'] = part
                                article['unite'] = raw_parts[idx + 1].upper()
                                quantity_found = True
                                skip_indices.add(idx)
                                skip_indices.add(idx + 1)
                        # Check if this is a unit and previous was quantity
                        elif part.upper() in known_units and not quantity_found and idx > 0:
                            if re.match(r'^\d+$', raw_parts[idx - 1]):
                                article['quantite'] = raw_parts[idx - 1]
                                article['unite'] = part.upper()
                                quantity_found = True
                                skip_indices.add(idx)
                                skip_indices.add(idx - 1)
                    
                    # Second pass: collect designation parts
                    for idx, part in enumerate(raw_parts):
                        if idx not in skip_indices:
                            designation_parts.append(part)
                    
                    # Smart designation extraction with line grouping
                    if designation_parts:
                        # Group consecutive lines that likely belong together
                        grouped_lines = []
                        current_group = []
                        
                        for i, part in enumerate(designation_parts):
                            part_clean = part.strip()
                            
                            # Skip empty, quantities, units
                            if not part_clean or part_clean == article['quantite'] or part_clean.upper() in ['UNI', 'FOR']:
                                if current_group:
                                    grouped_lines.append(' '.join(current_group))
                                    current_group = []
                                continue
                            
                            # Check if this line should be grouped with previous
                            # Group if: starts with lowercase, is short, or continues a phrase
                            should_group = False
                            if current_group:
                                # Check various conditions for grouping
                                if part_clean and len(part_clean) > 0:
                                    if (part_clean[0].islower() or  # Continuation of sentence
                                        len(part_clean) < 40 or      # Increased threshold for fragments
                                        part_clean.lower().startswith(('de ', 'à ', 'en ', 'pour ', 'avec ', 'et ', 'du ', 'des ', 'le ', 'la ', 'les ')) or  # French articles/prepositions
                                        part_clean.endswith(',') or  # Comma-ended lines likely continue
                                        re.match(r'^[a-z]', part_clean) or  # Starts with lowercase
                                        (len(current_group[-1]) < 20 and not current_group[-1].endswith('.'))  # Previous line was short and didn't end with period
                                    ):
                                        should_group = True
                                    # Special case: if current line looks like a continuation (e.g., "Elec" after "ConstrcElec")
                                    elif len(part_clean) < 10 and part_clean[0].isupper() and not part_clean.isupper():
                                        # Check if it could be part of a compound word
                                        last_word = current_group[-1].split()[-1] if current_group else ''
                                        if last_word and not last_word.endswith('.'):
                                            should_group = True
                            
                            if should_group:
                                current_group.append(part_clean)
                            else:
                                # Start new group
                                if current_group:
                                    grouped_lines.append(' '.join(current_group))
                                current_group = [part_clean]
                        
                        # Don't forget the last group
                        if current_group:
                            grouped_lines.append(' '.join(current_group))
                        
                        print(f"DEBUG: Grouped lines for position {pos}: {grouped_lines}")
                        
                        # Now filter the grouped lines
                        meaningful_parts = []
                        for group in grouped_lines:
                            if is_meaningful_text(group):
                                meaningful_parts.append(group)
                            else:
                                # If the group doesn't pass as a whole, try individual words
                                words = group.split()
                                filtered_words = [w for w in words if is_meaningful_text(w) and len(w) > 2]
                                if filtered_words:
                                    meaningful_parts.append(' '.join(filtered_words))
                        
                        # Join meaningful parts
                        if meaningful_parts:
                            complete_designation = ' '.join(meaningful_parts).strip()
                            complete_designation = re.sub(r'\s+', ' ', complete_designation)
                            
                            # Comprehensive cleanup - stronger artifact filtering
                            cleaned_designation = cleanup_artifacts(complete_designation)
                            
                            article['designation'] = cleaned_designation.strip()
                    
                    # If we still don't have quantity, try alternative methods
                    if not article['quantite']:
                        # Look more broadly for number + unit pattern
                        for k in range(i + 1, min(i + 20, len(text_lines))):
                            if text_lines[k].strip().upper() in known_units:
                                # Check surrounding lines for numbers
                                for l in range(max(0, k-3), k):
                                    if re.match(r'^\d+$', text_lines[l].strip()):
                                        article['quantite'] = text_lines[l].strip()
                                        article['unite'] = text_lines[k].strip().upper()
                                        break
                                if article['quantite']:
                                    break
                    
                    # Only add the article if it has a designation or quantity
                    if article['designation'] or article['quantite']:
                        articles.append(article)
                    
                    break
    
    # If automatic parsing failed, don't add any fallback articles
    # Let the user manually enter the data instead of providing fake content
    if not articles:
        print("DEBUG: No articles found - leaving empty for user to fill")
    
    extracted['articles'] = articles
    print(f"DEBUG: Total articles found: {len(articles)}")
    return extracted

def add_pac_interne_button(excel_path):
    """Add PAC Interne button with VBA instructions (avoid file corruption)"""
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        from openpyxl.comments import Comment
        import os
        
        print(f"Adding PAC Interne button to: {excel_path}")
        
        # Keep as .xlsx to avoid corruption - we'll provide VBA instructions separately
        wb = load_workbook(excel_path)
        
        # Add a hidden sheet to store VBA code instructions
        if 'VBA_Instructions' not in wb.sheetnames:
            vba_sheet = wb.create_sheet('VBA_Instructions')
            vba_sheet.sheet_state = 'hidden'
            
            # Store the complete VBA macro code
            vba_macro_code = '''
Sub GeneratePACDocument()
    ' VBA Macro to generate PAC Interne document
    ' This macro reads data from the current sheet and fills a Word template
    
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim templatePath As String
    Dim fournisseur As String
    Dim numeroCommande As String
    Dim numeroContrat As String
    Dim dateLivraison As String
    Dim articlesData As String
    Dim currentSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Get current sheet reference
    Set currentSheet = ActiveSheet
    
    ' Read data from the current sheet
    fournisseur = Trim(CStr(currentSheet.Range("B1").Value))
    numeroCommande = Trim(CStr(currentSheet.Range("B2").Value))
    numeroContrat = Trim(CStr(currentSheet.Range("B3").Value))
    dateLivraison = Trim(CStr(currentSheet.Range("B4").Value))
    
    ' Build articles list from the table (starting from row 7)
    articlesData = ""
    lastRow = currentSheet.Cells(currentSheet.Rows.Count, "A").End(xlUp).Row
    
    For i = 7 To lastRow
        If Trim(CStr(currentSheet.Cells(i, 1).Value)) <> "" And _
           Trim(CStr(currentSheet.Cells(i, 1).Value)) <> "TOTAL:" Then
            articlesData = articlesData & "• " & _
                          Trim(CStr(currentSheet.Cells(i, 1).Value)) & " - " & _
                          Trim(CStr(currentSheet.Cells(i, 2).Value)) & " (Qté: " & _
                          Trim(CStr(currentSheet.Cells(i, 3).Value)) & " " & _
                          Trim(CStr(currentSheet.Cells(i, 4).Value)) & ")" & vbCrLf
        End If
    Next i
    
    ' Set template path (adjust this path as needed)
    templatePath = ThisWorkbook.Path & "\\Exemple Pac Interne (2).doc"
    
    ' Check if template exists
    If Dir(templatePath) = "" Then
        MsgBox "Template Word non trouvé: " & templatePath & vbCrLf & vbCrLf & _
               "Veuillez placer le fichier 'Exemple Pac Interne (2).doc' dans le même dossier que ce fichier Excel.", _
               vbExclamation, "Fichier template manquant"
        Exit Sub
    End If
    
    ' Create Word application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    
    ' Open the template
    Set wordDoc = wordApp.Documents.Open(templatePath)
    
    ' Replace placeholders in the Word document
    ' These should match the yellow highlighted fields in your template
    
    With wordDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        
        ' Replace fournisseur placeholders
        .Text = "[FOURNISSEUR]"
        .Replacement.Text = fournisseur
        .Execute Replace:=2  ' Replace all
        
        ' Replace numero commande placeholders
        .Text = "[NUMERO_COMMANDE]"
        .Replacement.Text = numeroCommande
        .Execute Replace:=2
        
        ' Replace numero contrat placeholders
        .Text = "[NUMERO_CONTRAT]"
        .Replacement.Text = numeroContrat
        .Execute Replace:=2
        
        ' Replace date placeholders
        .Text = "[DATE_LIVRAISON]"
        .Replacement.Text = dateLivraison
        .Execute Replace:=2
        
        ' Replace articles list
        .Text = "[LISTE_ARTICLES]"
        .Replacement.Text = articlesData
        .Execute Replace:=2
        
        ' Additional common placeholders
        .Text = "[DATE_AUJOURDHUI]"
        .Replacement.Text = Format(Date, "dd/mm/yyyy")
        .Execute Replace:=2
        
        .Text = "[ENTREPRISE]"
        .Replacement.Text = "MEDI TELECOM"
        .Execute Replace:=2
    End With
    
    ' Success message
    MsgBox "Document PAC Interne généré avec succès!" & vbCrLf & vbCrLf & _
           "Fournisseur: " & fournisseur & vbCrLf & _
           "N° Commande: " & numeroCommande & vbCrLf & vbCrLf & _
           "Le document Word est maintenant ouvert et prêt à être utilisé.", _
           vbInformation, "PAC Interne - Succès"
    
    ' Cleanup
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors de la génération du document PAC:" & vbCrLf & _
           "Erreur " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Vérifiez que:" & vbCrLf & _
           "1. Microsoft Word est installé" & vbCrLf & _
           "2. Le template existe dans le bon dossier" & vbCrLf & _
           "3. Les macros sont activées", _
           vbCritical, "Erreur PAC Interne"
    
    ' Cleanup in case of error
    If Not wordDoc Is Nothing Then
        wordDoc.Close False
        Set wordDoc = Nothing
    End If
    If Not wordApp Is Nothing Then
        wordApp.Quit False
        Set wordApp = Nothing
    End If
End Sub

Sub PACInterneButton_Click()
    ' This is the macro that will be called when the button is clicked
    Call GeneratePACDocument
End Sub
'''
            
            # Store the VBA code in the hidden sheet
            vba_sheet['A1'] = 'VBA_MACRO_CODE_PAC_INTERNE'
            vba_sheet['A2'] = 'Copy this VBA code to a Module in Excel:'
            vba_sheet['A3'] = vba_macro_code
            
            # Add instructions for using the macro
            instructions = '''
INSTRUCTIONS POUR UTILISER LE MACRO PAC INTERNE:

1. ACTIVER LES MACROS:
   - Ouvrez ce fichier Excel
   - Cliquez sur "Activer le contenu" si demandé
   - Allez dans Fichier > Options > Centre de confiance > Paramètres du centre de confiance
   - Sélectionnez "Activer toutes les macros"

2. AJOUTER LE CODE VBA:
   - Appuyez sur Alt+F11 pour ouvrir l'éditeur VBA
   - Clic droit sur votre classeur dans la fenêtre de gauche
   - Choisissez Insertion > Module
   - Copiez et collez le code VBA depuis la cellule A3 de cette feuille
   - Fermez l'éditeur VBA (Alt+Q)

3. LIER LE BOUTON AU MACRO:
   - Clic droit sur le bouton "PAC Interne" (cellule H1)
   - Choisissez "Affecter une macro"
   - Sélectionnez "PACInterneButton_Click"
   - Cliquez OK

4. UTILISATION:
   - Cliquez sur le bouton "PAC Interne"
   - Le document Word s'ouvrira automatiquement avec les données remplies

REMARQUES:
- Le template Word doit être dans le même dossier que ce fichier Excel
- Les placeholders dans Word: [FOURNISSEUR], [NUMERO_COMMANDE], [NUMERO_CONTRAT], [DATE_LIVRAISON], [LISTE_ARTICLES]
'''
            
            vba_sheet['E1'] = 'INSTRUCTIONS_UTILISATION'
            vba_sheet['E2'] = instructions
        
        # Add buttons to each worksheet (except Main and hidden sheets)
        for sheet_name in wb.sheetnames:
            if sheet_name not in ['Main', 'VBA_Instructions']:
                sheet = wb[sheet_name]
                
                # Create a clickable button at H1 with hyperlink
                button_cell = sheet['H1']
                button_cell.value = '🔄 PAC Interne'
                
                # Add hyperlink to the Flask endpoint
                sheet_name_encoded = sheet_name.replace(' ', '%20')  # URL encode sheet name
                button_cell.hyperlink = f"http://localhost:5001/generate-pac/{sheet_name_encoded}"
                
                # Style it to look like an active button with hyperlink
                button_cell.font = Font(bold=True, color='FFFFFF', size=12, underline='single')
                button_cell.fill = PatternFill(start_color='FF6B35', end_color='FF6B35', fill_type='solid')  # Orange for active button
                button_cell.alignment = Alignment(horizontal='center', vertical='center')
                button_cell.border = Border(
                    left=Side(style='thick', color='CC5500'),
                    right=Side(style='thick', color='CC5500'),
                    top=Side(style='thick', color='CC5500'),
                    bottom=Side(style='thick', color='CC5500')
                )
                
                # Get data from the sheet for the comment
                fournisseur = str(sheet['B1'].value or 'N/A')
                numero_commande = str(sheet['B2'].value or 'N/A')
                numero_contrat = str(sheet['B3'].value or 'N/A')
                date_livraison = str(sheet['B4'].value or 'N/A')
                
                # Enhanced comment with hyperlink and macro instructions
                comment_text = (
                    "⚡ BOUTON PAC INTERNE AUTOMATIQUE ⚡\n\n"
                    f"📊 DONNÉES DÉTECTÉES:\n"
                    f"Fournisseur: {fournisseur}\n"
                    f"N° Commande: {numero_commande}\n"
                    f"N° Contrat: {numero_contrat}\n"
                    f"Date Livraison: {date_livraison}\n\n"
                    "🔗 UTILISATION DIRECTE (NOUVEAU!):\n"
                    "• Cliquez sur ce bouton pour générer le PAC\n"
                    "• Le serveur Flask doit être en cours d'exécution\n"
                    "• Le document sera généré automatiquement\n\n"
                    "🔧 ALTERNATIVE VBA (si pas de serveur):\n"
                    "1. Activez les macros Excel\n"
                    "2. Copiez le code VBA (feuille cachée)\n"
                    "3. Assignez la macro à ce bouton\n\n"
                    "✨ FONCTIONNALITÉ:\n"
                    "• Génère le document PAC Interne\n"
                    "• Remplit automatiquement les données\n"
                    "• Format Word (.docx) ou HTML"
                )
                
                comment = Comment(comment_text, 'PAC Automation System')
                comment.width = 350
                comment.height = 250
                button_cell.comment = comment
                
                # Make the button more prominent
                sheet.row_dimensions[1].height = 40
                sheet.column_dimensions['H'].width = 22
                
                # Add visual effects
                shadow_cells = ['I1', 'H2', 'I2']
                for shadow_cell in shadow_cells:
                    if shadow_cell == 'I1':
                        sheet[shadow_cell].fill = PatternFill(start_color='FFCCAA', end_color='FFCCAA', fill_type='solid')
                    elif shadow_cell == 'H2':
                        sheet[shadow_cell].fill = PatternFill(start_color='FFDDCC', end_color='FFDDCC', fill_type='solid')
                    elif shadow_cell == 'I2':
                        sheet[shadow_cell].fill = PatternFill(start_color='FFEEEE', end_color='FFEEEE', fill_type='solid')
        
        # Save as regular .xlsx to avoid corruption - provide VBA instructions separately
        wb.save(excel_path)
        wb.close()
        
        print("PAC Interne button with VBA instructions added successfully!")
        print("Next steps: Open Excel, enable macros, and copy VBA code from the hidden sheet.")
        
        # Create a separate .txt file with VBA instructions for easy access
        try:
            vba_instructions_file = excel_path.replace('.xlsx', '_VBA_Instructions.txt')
            with open(vba_instructions_file, 'w', encoding='utf-8') as f:
                f.write("=== VBA CODE FOR PAC INTERNE AUTOMATION ===\n\n")
                f.write(vba_macro_code)
                f.write("\n\n=== SETUP INSTRUCTIONS ===\n\n")
                f.write(instructions)
            print(f"VBA instructions also saved to: {vba_instructions_file}")
        except Exception as e:
            print(f"Could not create VBA instructions file: {str(e)}")
        
        return excel_path
        
    except Exception as e:
        print(f"Error adding PAC button with macro: {str(e)}")
        # Fallback to simple button
        try:
            return create_simple_pac_button(excel_path)
        except:
            return excel_path

def create_simple_pac_button(excel_path):
    """Fallback method to create a simple PAC button"""
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        from openpyxl.comments import Comment
        
        wb = load_workbook(excel_path)
        
        # Add buttons to each worksheet (except Main)
        for sheet_name in wb.sheetnames:
            if sheet_name not in ['Main', 'Instructions_PAC']:
                sheet = wb[sheet_name]
                
                # Create a visual button at H1
                button_cell = sheet['H1']
                button_cell.value = '📋 PAC Interne'
                
                # Style it to look like a button
                button_cell.font = Font(bold=True, color='FFFFFF', size=12)
                button_cell.fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
                button_cell.alignment = Alignment(horizontal='center', vertical='center')
                button_cell.border = Border(
                    left=Side(style='thick', color='2F4F4F'),
                    right=Side(style='thick', color='2F4F4F'),
                    top=Side(style='thick', color='2F4F4F'),
                    bottom=Side(style='thick', color='2F4F4F')
                )
                
                # Add functionality through comment and hyperlink
                comment_text = (
                    "🔹 BOUTON PAC INTERNE 🔹\n\n"
                    "Fournisseur: " + str(sheet['B1'].value or '') + "\n"
                    "N° Commande: " + str(sheet['B2'].value or '') + "\n\n"
                    "Pour utiliser:\n"
                    "• Activez les macros dans Excel\n"
                    "• Ou copiez les données manuellement\n\n"
                    "Double-cliquez sur cette cellule pour plus d'options."
                )
                
                comment = Comment(comment_text, 'PAC System')
                button_cell.comment = comment
                
                # Make the cell larger
                sheet.row_dimensions[1].height = 30
                sheet.column_dimensions['H'].width = 18
        
        # Save the workbook
        wb.save(excel_path)
        wb.close()
        
        print("Simple PAC button created successfully")
        return excel_path
        
    except Exception as e:
        print(f"Error creating simple PAC button: {str(e)}")
        return excel_path

def create_excel_multiple_sheets(pdf_data_list, filename='extracted_multiple.xlsx'):
    """Create Excel file with a main summary sheet and separate sheets for each PDF"""
    output_path = os.path.join('temp', filename)
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        
        # Prepare data for main sheet
        main_sheet_data = []
        
        for idx, pdf_data in enumerate(pdf_data_list, 1):
            # Create sheet name from numero de commande or use filename as fallback
            data = pdf_data['data']
            numero_commande = data.get('numero_commande', '')
            fournisseur = data.get('fournisseur', '')
            numero_contrat = data.get('numero_contrat', '')
            
            # Create sheet name from numero de commande or use filename as fallback
            if numero_commande:
                sheet_name = numero_commande[:31]  # Excel sheet name limit
            else:
                sheet_name = pdf_data.get('filename', f'PDF_{idx}').replace('.pdf', '')[:31]

            # Calculate total sum for "Budget engagé"
            total_sum = 0
            if data['articles']:
                articles_df = pd.DataFrame(data['articles'])
                
                # Calculate prix_total automatically for each row
                for index, row in articles_df.iterrows():
                    try:
                        quantite = float(row.get('quantite', 0)) if row.get('quantite') else 0
                        prix_unitaire = float(row.get('prix_unitaire', 0)) if row.get('prix_unitaire') else 0
                        articles_df.at[index, 'prix_total'] = f"{quantite * prix_unitaire:.2f}"
                        total_sum += quantite * prix_unitaire
                    except (ValueError, TypeError):
                        articles_df.at[index, 'prix_total'] = "0.00"
                
                # Set column names
                if 'pos_article' in articles_df.columns:
                    articles_df.columns = ['Pos Article', 'Désignation', 'Quantité', 'Unité', 'Prix unitaire (DH)', 'Total (DH)']
                else:
                    # Fallback for old structure
                    if 'pos' in articles_df.columns:
                        articles_df['pos_article'] = articles_df['pos']
                        articles_df = articles_df.drop(['pos'], axis=1, errors='ignore')
                        if 'article' in articles_df.columns:
                            articles_df = articles_df.drop(['article'], axis=1, errors='ignore')
                    articles_df.columns = ['Pos Article', 'Désignation', 'Quantité', 'Unité', 'Prix unitaire (DH)', 'Total (DH)']
            else:
                articles_df = pd.DataFrame(columns=['Pos Article', 'Désignation', 'Quantité', 'Unité', 'Prix unitaire (DH)', 'Total (DH)'])
            
            # Append data for main sheet
            main_sheet_data.append([
                fournisseur, numero_contrat, '', total_sum, '', '', ''  # Placeholders for Budget initial and Budget réceptionné
            ])

            # Write articles table starting from row 6
            articles_df.to_excel(writer, sheet_name=sheet_name, startrow=5, index=False)
            
            # Get the worksheet to add info at the top
            worksheet = writer.sheets[sheet_name]
            
            # Set column widths for better readability
            column_widths = {
                'A': 15,  # Pos Article
                'B': 50,  # Désignation (wider for better text display)
                'C': 12,  # Quantité
                'D': 10,  # Unité
                'E': 18,  # Prix unitaire
                'F': 18   # Total
            }
            
            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width
            
            # Add general information with enhanced styling
            info_font = Font(bold=True, size=11, color='1F4E79')  # Dark blue for labels
            value_font = Font(size=11, color='333333')  # Dark gray for values
            
            worksheet['A1'] = 'Fournisseur:'
            worksheet['A1'].font = info_font
            worksheet['B1'] = data.get('fournisseur', '')
            worksheet['B1'].font = value_font
            
            worksheet['A2'] = 'Numéro de commande:'
            worksheet['A2'].font = info_font
            worksheet['B2'] = data.get('numero_commande', '')
            worksheet['B2'].font = value_font
            
            worksheet['A3'] = 'Numéro de contrat:'
            worksheet['A3'].font = info_font
            worksheet['B3'] = data.get('numero_contrat', '')
            worksheet['B3'].font = value_font
            
            worksheet['A4'] = 'Date de livraison:'
            worksheet['A4'].font = info_font
            worksheet['B4'] = data.get('date_livraison', '')
            worksheet['B4'].font = value_font
            
            # Add light background to info section
            info_fill = PatternFill(start_color='F8F9FA', end_color='F8F9FA', fill_type='solid')
            for row in range(1, 5):
                for col in range(1, 7):  # A to F columns
                    worksheet.cell(row=row, column=col).fill = info_fill
            
            # Style the table headers with green theme (different from main sheet)
            header_fill = PatternFill(start_color='70AD47', end_color='70AD47', fill_type='solid')  # Green header
            header_font = Font(color='FFFFFF', bold=True, size=11)
            header_border = Border(
                left=Side(style='thin', color='FFFFFF'),
                right=Side(style='thin', color='FFFFFF'),
                top=Side(style='thin', color='FFFFFF'),
                bottom=Side(style='thin', color='FFFFFF')
            )
            
            for col in range(1, 7):  # A to F columns
                cell = worksheet.cell(row=6, column=col)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = header_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Style data rows with alternating colors and borders
            light_fill = PatternFill(start_color='F2F8F1', end_color='F2F8F1', fill_type='solid')  # Light green
            data_border = Border(
                left=Side(style='thin', color='D0D0D0'),
                right=Side(style='thin', color='D0D0D0'),
                top=Side(style='thin', color='D0D0D0'),
                bottom=Side(style='thin', color='D0D0D0')
            )
            data_font = Font(size=10)
            
            # Apply styling to data rows
            for row in range(7, 7 + len(articles_df)):  # Start from row 7 (after header)
                for col in range(1, 7):  # A to F columns
                    cell = worksheet.cell(row=row, column=col)
                    cell.border = data_border
                    cell.font = data_font
                    
                    # Alternate row colors
                    if row % 2 == 0:
                        cell.fill = light_fill
                    
                    # Center alignment for some columns
                    if col in [1, 3, 4]:  # Pos Article, Quantité, Unité
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    elif col in [5, 6]:  # Prix columns
                        cell.alignment = Alignment(horizontal='right', vertical='center')
                        # Format currency
                        if cell.value and str(cell.value).replace('.', '').replace(',', '').isdigit():
                            cell.number_format = '#,##0.00 "DH"'
                    else:  # Désignation
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            # Enhanced total row styling
            last_row = 6 + len(articles_df) + 1
            
            # Add "TOTAL" label with enhanced styling
            total_label_cell = worksheet.cell(row=last_row, column=5, value="TOTAL:")
            total_label_cell.font = Font(bold=True, size=12, color='70AD47')
            total_label_cell.alignment = Alignment(horizontal='right', vertical='center')
            
            # Add total sum with enhanced styling
            total_value_cell = worksheet.cell(row=last_row, column=6, value=f"{total_sum:.2f} DH")
            total_value_cell.font = Font(bold=True, size=12, color='70AD47')
            total_value_cell.alignment = Alignment(horizontal='right', vertical='center')
            total_value_cell.number_format = '#,##0.00 "DH"'
            
            # Add enhanced border to total row
            total_border = Border(
                top=Side(style='thick', color='70AD47'),
                bottom=Side(style='thin', color='70AD47')
            )
            for col in range(5, 7):
                worksheet.cell(row=last_row, column=col).border = total_border
            
            # Add some spacing
            worksheet.row_dimensions[6].height = 25  # Header row height
    
        # Add the main summary sheet with enhanced formatting
        main_df = pd.DataFrame(main_sheet_data, columns=[
            'Fournisseur', 'Numero de contrat', 'Budget initial', 'Budget engagé',
            'Budget réceptionné', 'Reliquat Eng', 'Reliquat recep'
        ])
        
        # Calculate the calculated columns for display
        for index, row in main_df.iterrows():
            try:
                budget_initial = float(row.get('Budget initial', 0)) if row.get('Budget initial') else 0
                budget_engage = float(row.get('Budget engagé', 0)) if row.get('Budget engagé') else 0
                budget_receptionne = float(row.get('Budget réceptionné', 0)) if row.get('Budget réceptionné') else 0
                
                # Calculate Reliquat Eng = Budget initial - Budget engagé
                main_df.at[index, 'Reliquat Eng'] = budget_initial - budget_engage if budget_initial > 0 else ''
                
                # Calculate Reliquat recep = Budget initial - Budget réceptionné
                main_df.at[index, 'Reliquat recep'] = budget_initial - budget_receptionne if budget_initial > 0 and budget_receptionne > 0 else ''
            except (ValueError, TypeError):
                pass
        
        main_df.to_excel(writer, sheet_name='Main', index=False)
        
        # Get the Main worksheet for styling
        main_worksheet = writer.sheets['Main']
        
        # Set column widths for better readability
        column_widths = {
            'A': 25,  # Fournisseur
            'B': 20,  # Numero de contrat
            'C': 18,  # Budget initial
            'D': 18,  # Budget engagé
            'E': 20,  # Budget réceptionné
            'F': 18,  # Reliquat Eng
            'G': 18   # Reliquat recep
        }
        
        for col, width in column_widths.items():
            main_worksheet.column_dimensions[col].width = width
        
        # Style the header row with colors
        header_fill = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')  # Blue header
        header_font = Font(color='FFFFFF', bold=True, size=12)
        header_border = Border(
            left=Side(style='thin', color='FFFFFF'),
            right=Side(style='thin', color='FFFFFF'),
            top=Side(style='thin', color='FFFFFF'),
            bottom=Side(style='thin', color='FFFFFF')
        )
        
        # Apply header styling
        for col in range(1, 8):  # A to G columns
            cell = main_worksheet.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = header_border
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Style data rows with alternating colors
        light_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')  # Light gray
        data_border = Border(
            left=Side(style='thin', color='CCCCCC'),
            right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),
            bottom=Side(style='thin', color='CCCCCC')
        )
        data_font = Font(size=11)
        
        # Apply styling to data rows
        for row in range(2, len(main_df) + 2):  # Start from row 2 (after header)
            for col in range(1, 8):  # A to G columns
                cell = main_worksheet.cell(row=row, column=col)
                cell.border = data_border
                cell.font = data_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # Alternate row colors
                if row % 2 == 0:
                    cell.fill = light_fill
                
                # Format currency columns
                if col in [3, 4, 5, 6, 7]:  # Budget columns
                    if cell.value and str(cell.value).replace('.', '').replace('-', '').isdigit():
                        cell.number_format = '#,##0.00 "DH"'
        
        # Add a title row above the table
        main_worksheet.insert_rows(1)
        title_cell = main_worksheet.cell(row=1, column=1, value="TABLEAU DE BORD - RÉSUMÉ DES BUDGETS")
        title_cell.font = Font(bold=True, size=16, color='2E75B6')
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Merge title across all columns
        main_worksheet.merge_cells('A1:G1')
        
        # Add some spacing
        main_worksheet.row_dimensions[1].height = 30
        main_worksheet.row_dimensions[2].height = 25
        
    # Note: VBA macro addition is commented out to prevent file corruption
    # Users can manually add the PAC interne macro using the instructions
    return output_path

def create_excel_single_sheet(data, filename='extracted_data.xlsx'):
    """Create Excel file from extracted data (single sheet)"""
    output_path = os.path.join('temp', filename)
    
    # Get numero de commande for worksheet name
    numero_commande = data.get('numero_commande', '')
    worksheet_name = numero_commande if numero_commande else 'Data'
    
    # Create articles dataframe
    if data['articles']:
        articles_df = pd.DataFrame(data['articles'])
        
        # Calculate prix_total automatically for each row
        for index, row in articles_df.iterrows():
            try:
                quantite = float(row.get('quantite', 0)) if row.get('quantite') else 0
                prix_unitaire = float(row.get('prix_unitaire', 0)) if row.get('prix_unitaire') else 0
                articles_df.at[index, 'prix_total'] = f"{quantite * prix_unitaire:.2f}"
            except (ValueError, TypeError):
                articles_df.at[index, 'prix_total'] = "0.00"
        
        # Set column names based on actual data structure
        if 'pos_article' in articles_df.columns:
            articles_df.columns = ['Pos Article', 'Désignation', 'Quantité', 'Unité', 'Prix unitaire (DH)', 'Total (DH)']
        else:
            # Fallback for old structure - combine pos and article if separate
            if 'pos' in articles_df.columns:
                articles_df['pos_article'] = articles_df['pos']
                articles_df = articles_df.drop(['pos'], axis=1, errors='ignore')
                if 'article' in articles_df.columns:
                    articles_df = articles_df.drop(['article'], axis=1, errors='ignore')
            articles_df.columns = ['Pos Article', 'Désignation', 'Quantité', 'Unité', 'Prix unitaire (DH)', 'Total (DH)']
    else:
        articles_df = pd.DataFrame(columns=['Pos Article', 'Désignation', 'Quantité', 'Unité', 'Prix unitaire (DH)', 'Total (DH)'])
    
    # Write to Excel - single sheet with info at top
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Write articles table starting from row 6 to leave space for info at top
        articles_df.to_excel(writer, sheet_name=worksheet_name, startrow=5, index=False)
        
        # Get the worksheet to add info at the top
        worksheet = writer.sheets[worksheet_name]
        
        # Add general information in the upper corner
        worksheet['A1'] = 'Fournisseur:'
        worksheet['B1'] = data.get('fournisseur', '')
        worksheet['A2'] = 'Numéro de commande:'
        worksheet['B2'] = data.get('numero_commande', '')
        worksheet['A3'] = 'Numéro de contrat:'
        worksheet['B3'] = data.get('numero_contrat', '')
        worksheet['A4'] = 'Date de livraison:'
        worksheet['B4'] = data.get('date_livraison', '')
        
        # Style the header cells
        from openpyxl.styles import Font, PatternFill, Border, Side
        
        # Make info labels bold
        for row in range(1, 5):
            worksheet[f'A{row}'].font = Font(bold=True)
        
        # Style the table headers
        header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        header_font = Font(color='FFFFFF', bold=True)
        
        for col in range(1, 7):  # A to F columns
            cell = worksheet.cell(row=6, column=col)
            cell.fill = header_fill
            cell.font = header_font
        
        # Add total sum row at the bottom
        last_row = 6 + len(articles_df) + 1  # 6 for start row + data rows + 1 for spacing
        
        # Calculate total sum
        total_sum = 0
        for index, row in articles_df.iterrows():
            try:
                prix_total = float(row.get('Total (DH)', 0)) if row.get('Total (DH)') else 0
                total_sum += prix_total
            except (ValueError, TypeError):
                pass
        
        # Add "TOTAL" label in column E (Prix unitaire column)
        worksheet.cell(row=last_row, column=5, value="TOTAL:")
        worksheet.cell(row=last_row, column=5).font = Font(bold=True)
        
        # Add total sum in column F (Total column)
        worksheet.cell(row=last_row, column=6, value=f"{total_sum:.2f} DH")
        worksheet.cell(row=last_row, column=6).font = Font(bold=True)
        
        # Add border to total row for better visibility
        thin_border = Border(top=Side(style='thin'))
        
        for col in range(5, 7):  # Columns E and F
            worksheet.cell(row=last_row, column=col).border = thin_border
    
    # Add PAC Interne button automatically
    try:
        output_path = add_pac_interne_button(output_path)
    except Exception as e:
        print(f"Warning: Could not add PAC button: {str(e)}")
    
    return output_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'files' not in request.files:
        return jsonify({'error': 'No files selected'}), 400
    
    files = request.files.getlist('files')
    if not files or all(file.filename == '' for file in files):
        return jsonify({'error': 'No files selected'}), 400
    
    results = []
    
    for file in files:
        if file and file.filename.endswith('.pdf'):
            try:
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                
                print(f"Processing file: {filename}")
                
                # Extract data from PDF
                extracted_data = extract_text_from_pdf(filepath)
                
                # Clean up uploaded PDF
                os.remove(filepath)
                
                results.append({
                    'filename': filename,
                    'success': True,
                    'data': extracted_data
                })
                
                print(f"Successfully processed: {filename}")
                
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")
                results.append({
                    'filename': filename,
                    'success': False,
                    'error': str(e)
                })
        else:
            results.append({
                'filename': file.filename,
                'success': False,
                'error': 'Invalid file type - only PDF files are supported'
            })
    
    # Return the results - frontend will handle display logic
    successful_results = [r for r in results if r['success']]
    
    if len(successful_results) == 0:
        return jsonify({
            'success': False,
            'error': 'No files were processed successfully'
        }), 400
    
    return jsonify({
        'success': True,
        'processed': len(successful_results),
        'failed': len([r for r in results if not r['success']]),
        'results': successful_results,
        'multiple_files': len(successful_results) > 1
    })

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    """Generate Excel file from user-modified data"""
    try:
        data = request.get_json()
        
        # Check if this is multiple PDFs data
        if 'pdf_data_list' in data:
            # Multiple PDFs - create separate sheets
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            excel_filename = f'combined_pdfs_{timestamp}.xlsx'
            excel_path = create_excel_multiple_sheets(data['pdf_data_list'], excel_filename)
            # Add PAC Interne button automatically
            try:
                excel_path = add_pac_interne_button(excel_path)
                excel_filename = os.path.basename(excel_path)
            except Exception as e:
                print(f"Warning: Could not add PAC button: {str(e)}")
        else:
            # Single PDF - use original method
            numero_commande = data.get('numero_commande', '')
            if numero_commande:
                excel_filename = f'{numero_commande}.xlsx'
            else:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                excel_filename = f'extracted_{timestamp}.xlsx'
            
            excel_path = create_excel_single_sheet(data, excel_filename)
        
        return jsonify({
            'success': True,
            'excel_file': excel_filename
        })
    except Exception as e:
        print(f"Error generating Excel: {str(e)}")
        return jsonify({'error': 'Failed to generate Excel file'}), 500

@app.route('/generate-pac/<sheet_name>')
def generate_pac_document(sheet_name):
    """Generate PAC document from Excel sheet data"""
    try:
        # Find the most recent Excel file containing this sheet
        excel_files = [f for f in os.listdir('temp') if f.endswith('.xlsx') and os.path.isfile(os.path.join('temp', f))]
        excel_files.sort(key=lambda x: os.path.getmtime(os.path.join('temp', x)), reverse=True)
        
        excel_path = None
        wb = None
        
        # Find the Excel file containing the requested sheet
        for excel_file in excel_files:
            try:
                from openpyxl import load_workbook
                wb_temp = load_workbook(os.path.join('temp', excel_file), read_only=True)
                if sheet_name in wb_temp.sheetnames:
                    excel_path = os.path.join('temp', excel_file)
                    wb = load_workbook(excel_path)
                    break
                wb_temp.close()
            except:
                continue
        
        if not excel_path or not wb:
            return jsonify({'error': 'Sheet not found'}), 404
        
        # Get the sheet
        sheet = wb[sheet_name]
        
        # Extract data from the sheet
        fournisseur = str(sheet['B1'].value or '')
        numero_commande = str(sheet['B2'].value or '')
        numero_contrat = str(sheet['B3'].value or '')
        date_livraison = str(sheet['B4'].value or '')
        
        # Extract articles
        articles = []
        for row in range(7, sheet.max_row + 1):
            if sheet.cell(row, 1).value and str(sheet.cell(row, 1).value) != 'TOTAL:':
                articles.append({
                    'pos': str(sheet.cell(row, 1).value or ''),
                    'designation': str(sheet.cell(row, 2).value or ''),
                    'quantite': str(sheet.cell(row, 3).value or ''),
                    'unite': str(sheet.cell(row, 4).value or '')
                })
        
        wb.close()
        
        # Check if Word template exists
        template_path = os.path.join(os.path.dirname(excel_path), 'Exemple Pac Interne (2).doc')
        if not os.path.exists(template_path):
            template_path = 'Exemple Pac Interne (2).doc'
            if not os.path.exists(template_path):
                # Create a simple Word document if template doesn't exist
                return generate_simple_pac_document(fournisseur, numero_commande, numero_contrat, date_livraison, articles)
        
        # Try to use Word automation
        try:
            import win32com.client
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = True
            doc = word.Documents.Open(os.path.abspath(template_path))
            
            # Replace placeholders
            find_and_replace(doc, '[FOURNISSEUR]', fournisseur)
            find_and_replace(doc, '[NUMERO_COMMANDE]', numero_commande)
            find_and_replace(doc, '[NUMERO_CONTRAT]', numero_contrat)
            find_and_replace(doc, '[DATE_LIVRAISON]', date_livraison)
            find_and_replace(doc, '[DATE_AUJOURDHUI]', datetime.now().strftime('%d/%m/%Y'))
            find_and_replace(doc, '[ENTREPRISE]', 'MEDI TELECOM')
            
            # Build articles list
            articles_text = ''
            for art in articles:
                articles_text += f"• {art['pos']} - {art['designation']} (Qté: {art['quantite']} {art['unite']})\n"
            find_and_replace(doc, '[LISTE_ARTICLES]', articles_text)
            
            return jsonify({'success': True, 'message': 'Document PAC opened in Word'})
            
        except Exception as e:
            # Fallback to simple document generation
            return generate_simple_pac_document(fournisseur, numero_commande, numero_contrat, date_livraison, articles)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def find_and_replace(doc, find_text, replace_text):
    """Helper function to find and replace in Word document"""
    try:
        find = doc.Content.Find
        find.ClearFormatting()
        find.Text = find_text
        find.Replacement.ClearFormatting()
        find.Replacement.Text = replace_text
        find.Execute(Replace=2)  # wdReplaceAll
    except:
        pass

def generate_simple_pac_document(fournisseur, numero_commande, numero_contrat, date_livraison, articles):
    """Generate a simple PAC document using python-docx"""
    try:
        from docx import Document
        from docx.shared import Pt, RGBColor, Inches, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
        from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        
        # Create new document
        doc = Document()
        
        # Set page margins
        section = doc.sections[0]
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        
        # Add header table for logo and title
        header_table = doc.add_table(rows=1, cols=3)
        header_table.style = 'Table Grid'
        header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Left cell - Orange logo placeholder
        left_cell = header_table.rows[0].cells[0]
        left_cell.width = Inches(1.5)
        logo_para = left_cell.paragraphs[0]
        logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo_run = logo_para.add_run('orange')
        logo_run.font.name = 'Arial'
        logo_run.font.size = Pt(24)
        logo_run.font.bold = True
        logo_run.font.color.rgb = RGBColor(255, 140, 0)  # Orange color
        
        # Set orange background for logo cell
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), 'FF8C00')
        left_cell._element.get_or_add_tcPr().append(shading_elm)
        
        # Middle cell - Title
        middle_cell = header_table.rows[0].cells[1]
        middle_cell.width = Inches(3.5)
        title_para = middle_cell.paragraphs[0]
        title_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        title_run = title_para.add_run('Procès-Verbal de Réception\nEn Interne')
        title_run.font.name = 'Arial'
        title_run.font.size = Pt(16)
        title_run.font.bold = True
        title_run.font.color.rgb = RGBColor(255, 255, 255)  # White text
        
        # Set dark background for title cell
        shading_elm2 = OxmlElement('w:shd')
        shading_elm2.set(qn('w:fill'), '2F2F2F')
        middle_cell._element.get_or_add_tcPr().append(shading_elm2)
        
        # Right cell - Direction info
        right_cell = header_table.rows[0].cells[2]
        right_cell.width = Inches(2.5)
        dir_para = right_cell.paragraphs[0]
        dir_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        dir_run = dir_para.add_run('Direction DID\nDirection Centrale DCT')
        dir_run.font.name = 'Arial'
        dir_run.font.size = Pt(10)
        dir_run.font.bold = True
        dir_run.font.color.rgb = RGBColor(255, 255, 255)  # White text
        
        # Set dark background for direction cell
        shading_elm3 = OxmlElement('w:shd')
        shading_elm3.set(qn('w:fill'), '2F2F2F')
        right_cell._element.get_or_add_tcPr().append(shading_elm3)
        
        # Add some space
        doc.add_paragraph()
        
        # Add supplier and contract info in two columns
        info_table = doc.add_table(rows=5, cols=2)
        info_table.style = None
        
        # Left column
        info_table.cell(0, 0).text = 'Fournisseur :'
        fournisseur_para = info_table.cell(1, 0).paragraphs[0]
        fournisseur_run = fournisseur_para.add_run(fournisseur)
        fournisseur_run.font.bold = True
        fournisseur_run.font.color.rgb = RGBColor(255, 255, 0)  # Yellow highlight effect
        
        info_table.cell(2, 0).text = 'Objet :'
        
        info_table.cell(3, 0).text = 'Référence PV'
        
        date_cell = info_table.cell(4, 0)
        date_cell.text = 'Date PV : '
        date_run = date_cell.paragraphs[0].add_run(datetime.now().strftime('%d/%m/%Y'))
        date_run.font.bold = True
        date_run.font.color.rgb = RGBColor(255, 255, 0)  # Yellow highlight effect
        
        # Right column
        info_table.cell(0, 1).text = 'Facture'
        
        contrat_cell = info_table.cell(1, 1)
        contrat_cell.text = 'Réf. Contrat Cadre SAP : '
        contrat_run = contrat_cell.paragraphs[0].add_run(numero_contrat)
        contrat_run.font.bold = True
        contrat_run.font.color.rgb = RGBColor(255, 255, 0)  # Yellow highlight effect
        
        commande_cell = info_table.cell(2, 1)
        commande_cell.text = 'Bon de commande: '
        commande_run = commande_cell.paragraphs[0].add_run(numero_commande)
        commande_run.font.bold = True
        commande_run.font.color.rgb = RGBColor(255, 255, 0)  # Yellow highlight effect
        
        info_table.cell(3, 1).text = 'Période Du'
        
        info_table.cell(4, 1).text = 'Au'
        
        # Style the info table
        for row in info_table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.style = doc.styles['Normal']
                    for run in paragraph.runs:
                        if not run.font.bold:
                            run.font.name = 'Arial'
                            run.font.size = Pt(11)
        
        doc.add_paragraph()
        
        # Add contract reference line
        ref_para = doc.add_paragraph('Réf. Contrat')
        ref_para.style = doc.styles['Normal']
        ref_para.runs[0].font.name = 'Arial'
        ref_para.runs[0].font.size = Pt(11)
        
        doc.add_paragraph()
        
        # Add main text
        main_text = (
            f"Les services/équipements relatives aux contrat/BC/Facture/Période cités ci-dessus ont été bien réalisés selon les "
            f"termes du bon de commande et/ou dispositions contractuelles convenus entre Méditelecom et {fournisseur}. "
            f"Ce PV a été dressé aux fins de preuve et de confirmation de la bonne réalisation/livraison/installation des dites "
            f"prestations/équipements sous réserves des remarques ci-dessous"
        )
        main_para = doc.add_paragraph(main_text)
        main_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        main_para.runs[0].font.name = 'Arial'
        main_para.runs[0].font.size = Pt(11)
        
        # Add section title
        disp_para = doc.add_paragraph()
        disp_run = disp_para.add_run('Dispositions Contractuelles Sujettes de Validation')
        disp_run.font.name = 'Arial'
        disp_run.font.size = Pt(12)
        disp_run.font.bold = True
        disp_run.font.underline = True
        
        doc.add_paragraph()
        
        # Create articles table
        if articles:
            articles_table = doc.add_table(rows=len(articles) + 1, cols=3)
            articles_table.style = 'Table Grid'
            
            # Header row
            header_cells = articles_table.rows[0].cells
            header_cells[0].text = 'PoS'
            header_cells[1].text = 'Items'
            header_cells[2].text = 'Qté'
            
            # Style header row with gray background
            for cell in header_cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell.paragraphs[0].runs[0].font.bold = True
                cell.paragraphs[0].runs[0].font.name = 'Arial'
                cell.paragraphs[0].runs[0].font.size = Pt(11)
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(qn('w:fill'), 'D3D3D3')
                cell._element.get_or_add_tcPr().append(shading_elm)
            
            # Add article rows
            for idx, article in enumerate(articles):
                row = articles_table.rows[idx + 1]
                
                # Position with yellow background
                pos_cell = row.cells[0]
                pos_cell.text = article['pos']
                pos_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                pos_cell.paragraphs[0].runs[0].font.bold = True
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(qn('w:fill'), 'FFFF00')
                pos_cell._element.get_or_add_tcPr().append(shading_elm)
                
                # Item description
                row.cells[1].text = article['designation']
                row.cells[1].paragraphs[0].runs[0].font.name = 'Arial'
                row.cells[1].paragraphs[0].runs[0].font.size = Pt(11)
                
                # Quantity with yellow background
                qty_cell = row.cells[2]
                qty_cell.text = article['quantite']
                qty_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                qty_cell.paragraphs[0].runs[0].font.bold = True
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(qn('w:fill'), 'FFFF00')
                qty_cell._element.get_or_add_tcPr().append(shading_elm)
        
        # Add space before signature section
        for _ in range(3):
            doc.add_paragraph()
        
        # Add payment question
        payment_para = doc.add_paragraph()
        payment_para.add_run('Est-ce que ce document constitue un « Bon à Payer » ?     ').font.italic = True
        payment_para.add_run('Non').font.bold = True
        payment_para.runs[0].font.name = 'Arial'
        payment_para.runs[0].font.size = Pt(11)
        payment_para.runs[1].font.name = 'Arial'
        payment_para.runs[1].font.size = Pt(11)
        
        doc.add_paragraph()
        
        # Add signature table
        sig_table = doc.add_table(rows=5, cols=3)
        sig_table.style = 'Table Grid'
        sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Header row - Orange
        sig_table.cell(0, 1).text = 'Orange'
        sig_table.cell(0, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        sig_table.cell(0, 1).paragraphs[0].runs[0].font.bold = True
        
        # Signature details
        sig_table.cell(1, 0).text = 'Nom du signataire'
        sig_table.cell(1, 1).text = 'Abdelhay Derraz'
        
        sig_table.cell(2, 0).text = 'Qualité du signataire'
        sig_table.cell(2, 1).text = 'Senior Manager'
        
        sig_table.cell(3, 0).text = 'Date'
        sig_table.cell(3, 1).text = datetime.now().strftime('%d/%m/%Y')
        sig_table.cell(3, 1).paragraphs[0].runs[0].font.bold = True
        sig_table.cell(3, 1).paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 0)
        
        sig_table.cell(4, 0).text = 'Cachet et signature'
        
        # Style signature table
        for row in sig_table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.style = doc.styles['Normal']
                    for run in paragraph.runs:
                        run.font.name = 'Arial'
                        run.font.size = Pt(11)
        
        # Save the document
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        doc_filename = f'PAC_Interne_{numero_commande}_{timestamp}.docx'
        doc_path = os.path.join('temp', doc_filename)
        doc.save(doc_path)
        
        # Return the file
        return send_file(doc_path, as_attachment=True, download_name=doc_filename)
        
    except ImportError:
        # If python-docx is not installed, create HTML version
        html_content = f"""
        <html>
        <head>
            <title>PAC Interne - {numero_commande}</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 40px; }}
                h1 {{ color: #2E75B6; }}
                .info {{ margin: 20px 0; }}
                .articles {{ margin: 20px 0; }}
                .signature {{ margin-top: 50px; }}
            </style>
        </head>
        <body>
            <h1 style="text-align: center;">PROCÈS-VERBAL DE CONTRÔLE INTERNE</h1>
            <p>Date: {datetime.now().strftime('%d/%m/%Y')}</p>
            
            <div class="info">
                <h2>Informations Générales</h2>
                <p><strong>Fournisseur:</strong> {fournisseur}</p>
                <p><strong>Numéro de commande:</strong> {numero_commande}</p>
                <p><strong>Numéro de contrat:</strong> {numero_contrat}</p>
                <p><strong>Date de livraison:</strong> {date_livraison}</p>
            </div>
            
            <div class="articles">
                <h2>Articles Contrôlés</h2>
                <ul>
        """
        
        for art in articles:
            html_content += f"<li>{art['pos']} - {art['designation']} (Quantité: {art['quantite']} {art['unite']})</li>\n"
        
        html_content += """
                </ul>
            </div>
            
            <div class="signature">
                <h2>Validation</h2>
                <p>Contrôlé par: _______________________</p>
                <p>Date: _______________________</p>
                <p>Signature: _______________________</p>
            </div>
        </body>
        </html>
        """
        
        # Save HTML file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        html_filename = f'PAC_Interne_{numero_commande}_{timestamp}.html'
        html_path = os.path.join('temp', html_filename)
        
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return send_file(html_path, as_attachment=True, download_name=html_filename)

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join('temp', filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True, download_name=filename)
    return jsonify({'error': 'File not found'}), 404

if __name__ == '__main__':
    app.run(debug=True, port=5001)
