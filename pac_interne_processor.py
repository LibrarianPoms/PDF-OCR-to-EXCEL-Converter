import os
import pandas as pd
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor
import win32com.client
from datetime import datetime
import shutil

class PACInterneProcessor:
    def __init__(self, excel_file, word_template):
        self.excel_file = excel_file
        self.word_template = word_template
        self.excel_data = None
        
    def read_excel_data(self):
        """Read data from Excel file"""
        try:
            # Read Excel file
            self.excel_data = pd.read_excel(self.excel_file)
            print(f"Loaded {len(self.excel_data)} rows from Excel")
            print(f"Columns: {list(self.excel_data.columns)}")
            return True
        except Exception as e:
            print(f"Error reading Excel: {str(e)}")
            return False
    
    def convert_doc_to_docx(self, doc_path):
        """Convert .doc to .docx using Word COM"""
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            # Open the .doc file
            doc = word.Documents.Open(os.path.abspath(doc_path))
            
            # Save as .docx
            docx_path = doc_path.replace('.doc', '.docx')
            doc.SaveAs2(os.path.abspath(docx_path), FileFormat=16)  # 16 = wdFormatXMLDocument
            
            # Close
            doc.Close()
            word.Quit()
            
            print(f"Converted {doc_path} to {docx_path}")
            return docx_path
        except Exception as e:
            print(f"Error converting doc to docx: {str(e)}")
            return None
    
    def find_highlighted_text(self, paragraph):
        """Find yellow highlighted text in a paragraph"""
        highlighted_runs = []
        for run in paragraph.runs:
            # Check if the run has yellow highlighting
            if run.font.highlight_color == WD_COLOR_INDEX.YELLOW:
                highlighted_runs.append(run)
        return highlighted_runs
    
    def process_word_document(self, row_data, output_filename):
        """Process Word document and replace highlighted text with data from Excel row"""
        try:
            # Convert .doc to .docx if needed
            if self.word_template.endswith('.doc'):
                docx_path = self.convert_doc_to_docx(self.word_template)
                if not docx_path:
                    return False
            else:
                docx_path = self.word_template
            
            # Open the document
            doc = Document(docx_path)
            
            # Map Excel columns to replacement values
            replacements = {
                'NUMERO_COMMANDE': str(row_data.get('Numéro de commande', '')),
                'FOURNISSEUR': str(row_data.get('Fournisseur', '')),
                'DATE_LIVRAISON': str(row_data.get('Date de livraison', '')),
                'NUMERO_CONTRAT': str(row_data.get('Numéro de contrat', '')),
                'DESIGNATION': str(row_data.get('Désignation', '')),
                'QUANTITE': str(row_data.get('Quantité', '')),
                'PRIX_UNITAIRE': str(row_data.get('Prix unitaire', '')),
                'MONTANT_TOTAL': str(row_data.get('Montant total', '')),
            }
            
            # Process all paragraphs
            for paragraph in doc.paragraphs:
                for run in paragraph.runs:
                    # Check if run is highlighted in yellow
                    if run.font.highlight_color == WD_COLOR_INDEX.YELLOW:
                        original_text = run.text
                        # Try to match with our replacements
                        for key, value in replacements.items():
                            if key in original_text.upper():
                                run.text = value
                                # Remove highlighting after replacement
                                run.font.highlight_color = None
                                print(f"Replaced {original_text} with {value}")
                                break
            
            # Also check tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                if run.font.highlight_color == WD_COLOR_INDEX.YELLOW:
                                    original_text = run.text
                                    for key, value in replacements.items():
                                        if key in original_text.upper():
                                            run.text = value
                                            run.font.highlight_color = None
                                            print(f"Replaced in table: {original_text} with {value}")
                                            break
            
            # Save the document
            doc.save(output_filename)
            print(f"Saved processed document as {output_filename}")
            return True
            
        except Exception as e:
            print(f"Error processing Word document: {str(e)}")
            return False
    
    def add_pac_button_to_excel(self):
        """Add PAC Interne button to Excel sheets with numero de commande"""
        try:
            # Use Excel COM to add button with macro
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            
            # Open workbook
            wb = excel.Workbooks.Open(os.path.abspath(self.excel_file))
            
            # Create output filename
            output_file = self.excel_file.replace('.xlsx', '_with_PAC_button.xlsm').replace('.xlsm', '_with_PAC_button.xlsm')
            
            # Process each sheet
            for sheet in wb.Sheets:
                # Check if sheet has "Numéro de commande" column
                has_numero_commande = False
                try:
                    # Check first row for headers
                    for col in range(1, sheet.UsedRange.Columns.Count + 1):
                        cell_value = sheet.Cells(1, col).Value
                        if cell_value and 'numéro de commande' in str(cell_value).lower():
                            has_numero_commande = True
                            break
                except:
                    pass
                
                if has_numero_commande:
                    # Add button to sheet
                    button = sheet.Buttons().Add(Left=10, Top=10, Width=120, Height=30)
                    button.Caption = "PAC Interne"
                    button.Name = f"btnPACInterne_{sheet.Name}"
                    
                    # Add macro code
                    macro_code = f'''
Sub {button.Name}_Click()
    Dim currentRow As Long
    Dim numCommande As String
    
    ' Get the active row
    currentRow = ActiveCell.Row
    
    ' Find numero de commande column
    Dim numCommandeCol As Integer
    For i = 1 To ActiveSheet.UsedRange.Columns.Count
        If InStr(1, LCase(Cells(1, i).Value), "numéro de commande") > 0 Then
            numCommandeCol = i
            Exit For
        End If
    Next i
    
    ' Get numero de commande value
    If numCommandeCol > 0 And currentRow > 1 Then
        numCommande = Cells(currentRow, numCommandeCol).Value
        MsgBox "Traitement PAC Interne pour commande: " & numCommande, vbInformation
        
        ' Here you can add code to process the PAC Interne document
        ' For now, just show the message
    Else
        MsgBox "Veuillez sélectionner une ligne avec un numéro de commande", vbExclamation
    End If
End Sub
'''
                    # Add the macro to the workbook
                    wb.VBProject.VBComponents.Add(1).CodeModule.AddFromString(macro_code)
                    button.OnAction = f"{button.Name}_Click"
                    
                    print(f"Added PAC Interne button to sheet: {sheet.Name}")
            
            # Save as macro-enabled workbook
            wb.SaveAs(os.path.abspath(output_file), FileFormat=52)  # 52 = xlOpenXMLWorkbookMacroEnabled
            wb.Close()
            excel.Quit()
            
            print(f"Saved Excel with PAC buttons as: {output_file}")
            return output_file
            
        except Exception as e:
            print(f"Error adding PAC button to Excel: {str(e)}")
            if 'excel' in locals():
                excel.Quit()
            return None
    
    def process_all_commands(self):
        """Process PAC Interne documents for all commands in Excel"""
        if not self.read_excel_data():
            return False
        
        # Create output directory
        output_dir = "PAC_Interne_Output"
        os.makedirs(output_dir, exist_ok=True)
        
        # Process each row
        for index, row in self.excel_data.iterrows():
            numero_commande = row.get('Numéro de commande', '')
            if numero_commande:
                output_filename = os.path.join(output_dir, f"PAC_Interne_{numero_commande}.docx")
                print(f"\nProcessing command: {numero_commande}")
                self.process_word_document(row, output_filename)
        
        # Add PAC button to Excel
        self.add_pac_button_to_excel()
        
        return True


def main():
    # Configuration
    excel_file = "combined_pdfs_with_PAC_button.xlsx"  # Your Excel file
    word_template = "Exemple Pac Interne (2).doc"  # Your Word template
    
    # Create processor
    processor = PACInterneProcessor(excel_file, word_template)
    
    # Process all commands
    if processor.process_all_commands():
        print("\nProcessing completed successfully!")
        print("Check the 'PAC_Interne_Output' folder for generated documents")
        print("Excel file with PAC buttons has been created")
    else:
        print("\nProcessing failed!")


if __name__ == "__main__":
    main()
