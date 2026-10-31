import os
import pandas as pd
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import win32com.client
from datetime import datetime
import pythoncom

class PACInterneIntegrated:
    def __init__(self):
        self.excel_file = None
        self.word_template = "Exemple Pac Interne (2).doc"
        
    def create_excel_with_pac_macro(self, input_excel):
        """Create Excel file with integrated PAC Interne macro"""
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Open workbook
            wb = excel.Workbooks.Open(os.path.abspath(input_excel))
            
            # Create output filename
            output_file = input_excel.replace('.xlsx', '_PAC_Integrated.xlsm').replace('.xlsm', '_PAC_Integrated.xlsm')
            
            # Add the main PAC processing module
            vb_module = wb.VBProject.VBComponents.Add(1)
            vb_module.Name = "PACInterneModule"
            
            module_code = '''
Option Explicit

Sub ProcessPACInterne(numCommande As String, rowData As Range)
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim docPath As String
    Dim outputPath As String
    Dim para As Object
    Dim run As Object
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' Set paths
    docPath = ThisWorkbook.Path & "\\Exemple Pac Interne (2).doc"
    outputPath = ThisWorkbook.Path & "\\PAC_Interne_Output\\"
    
    ' Create output directory if it doesn't exist
    If Dir(outputPath, vbDirectory) = "" Then
        MkDir outputPath
    End If
    
    ' Create Word application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    ' Open the template document
    Set wordDoc = wordApp.Documents.Open(docPath)
    
    ' Get column headers
    Dim headers As Range
    Set headers = rowData.Parent.Range("1:1")
    
    ' Process replacements
    Dim j As Integer
    For Each para In wordDoc.Paragraphs
        For j = 1 To para.Range.Runs.Count
            Set run = para.Range.Runs(j)
            ' Check if text is highlighted in yellow
            If run.HighlightColorIndex = 7 Then ' 7 = wdYellow
                ' Replace based on column mapping
                Dim cellValue As String
                Select Case UCase(Trim(run.Text))
                    Case "NUMERO_COMMANDE", "NUMÉRO DE COMMANDE"
                        cellValue = GetCellValue(rowData, headers, "Numéro de commande")
                    Case "FOURNISSEUR"
                        cellValue = GetCellValue(rowData, headers, "Fournisseur")
                    Case "DATE_LIVRAISON", "DATE DE LIVRAISON"
                        cellValue = GetCellValue(rowData, headers, "Date de livraison")
                    Case "NUMERO_CONTRAT", "NUMÉRO DE CONTRAT"
                        cellValue = GetCellValue(rowData, headers, "Numéro de contrat")
                    Case "DESIGNATION", "DÉSIGNATION"
                        cellValue = GetCellValue(rowData, headers, "Désignation")
                    Case "QUANTITE", "QUANTITÉ"
                        cellValue = GetCellValue(rowData, headers, "Quantité")
                    Case "PRIX_UNITAIRE", "PRIX UNITAIRE"
                        cellValue = GetCellValue(rowData, headers, "Prix unitaire")
                    Case "MONTANT_TOTAL", "MONTANT TOTAL"
                        cellValue = GetCellValue(rowData, headers, "Montant total")
                End Select
                
                If cellValue <> "" Then
                    run.Text = cellValue
                    run.HighlightColorIndex = 0 ' Remove highlight
                End If
            End If
        Next j
    Next para
    
    ' Also process tables
    Dim tbl As Object
    Dim row As Object
    Dim cell As Object
    For Each tbl In wordDoc.Tables
        For Each row In tbl.Rows
            For Each cell In row.Cells
                For Each para In cell.Range.Paragraphs
                    For j = 1 To para.Range.Runs.Count
                        Set run = para.Range.Runs(j)
                        If run.HighlightColorIndex = 7 Then
                            ' Same replacement logic as above
                            Dim tableCellValue As String
                            Select Case UCase(Trim(run.Text))
                                Case "NUMERO_COMMANDE", "NUMÉRO DE COMMANDE"
                                    tableCellValue = GetCellValue(rowData, headers, "Numéro de commande")
                                Case "FOURNISSEUR"
                                    tableCellValue = GetCellValue(rowData, headers, "Fournisseur")
                                Case "DATE_LIVRAISON", "DATE DE LIVRAISON"
                                    tableCellValue = GetCellValue(rowData, headers, "Date de livraison")
                                Case "NUMERO_CONTRAT", "NUMÉRO DE CONTRAT"
                                    tableCellValue = GetCellValue(rowData, headers, "Numéro de contrat")
                                Case "DESIGNATION", "DÉSIGNATION"
                                    tableCellValue = GetCellValue(rowData, headers, "Désignation")
                                Case "QUANTITE", "QUANTITÉ"
                                    tableCellValue = GetCellValue(rowData, headers, "Quantité")
                                Case "PRIX_UNITAIRE", "PRIX UNITAIRE"
                                    tableCellValue = GetCellValue(rowData, headers, "Prix unitaire")
                                Case "MONTANT_TOTAL", "MONTANT TOTAL"
                                    tableCellValue = GetCellValue(rowData, headers, "Montant total")
                            End Select
                            
                            If tableCellValue <> "" Then
                                run.Text = tableCellValue
                                run.HighlightColorIndex = 0
                            End If
                        End If
                    Next j
                Next para
            Next cell
        Next row
    Next tbl
    
    ' Save the document
    Dim outputFile As String
    outputFile = outputPath & "PAC_Interne_" & numCommande & ".docx"
    wordDoc.SaveAs2 outputFile, FileFormat:=16 ' 16 = wdFormatXMLDocument
    
    ' Close Word
    wordDoc.Close
    wordApp.Quit
    
    MsgBox "PAC Interne document créé avec succès!" & vbCrLf & vbCrLf & _
           "Fichier: " & outputFile, vbInformation, "Succès"
    
    Exit Sub
    
ErrorHandler:
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    MsgBox "Erreur lors du traitement: " & Err.Description, vbCritical, "Erreur"
End Sub

Function GetCellValue(rowData As Range, headers As Range, columnName As String) As String
    Dim col As Integer
    Dim cell As Range
    
    ' Find column index
    For Each cell In headers.Cells
        If cell.Value <> "" Then
            If InStr(1, LCase(cell.Value), LCase(columnName)) > 0 Then
                col = cell.Column
                Exit For
            End If
        End If
    Next cell
    
    ' Get value from the row
    If col > 0 Then
        GetCellValue = CStr(rowData.Cells(1, col).Value)
    Else
        GetCellValue = ""
    End If
End Function
'''
            vb_module.CodeModule.AddFromString(module_code)
            
            # Now add button click handlers for each sheet
            for sheet in wb.Sheets:
                # Check if sheet has "Numéro de commande" column
                has_numero_commande = False
                numero_commande_col = 0
                
                try:
                    # Check first row for headers
                    for col in range(1, sheet.UsedRange.Columns.Count + 1):
                        cell_value = sheet.Cells(1, col).Value
                        if cell_value and 'numéro de commande' in str(cell_value).lower():
                            has_numero_commande = True
                            numero_commande_col = col
                            break
                except:
                    pass
                
                if has_numero_commande:
                    # Add button to sheet
                    button = sheet.Buttons().Add(Left=10, Top=10, Width=120, Height=30)
                    button.Caption = "PAC Interne"
                    button.Name = f"btnPACInterne_{sheet.Name.replace(' ', '_')}"
                    
                    # Add sheet-specific code
                    sheet_module = wb.VBProject.VBComponents(sheet.Name)
                    button_code = f'''
Private Sub {button.Name}_Click()
    Dim currentRow As Long
    Dim numCommande As String
    Dim rowData As Range
    
    ' Get the active row
    currentRow = ActiveCell.Row
    
    ' Validate selection
    If currentRow <= 1 Then
        MsgBox "Veuillez sélectionner une ligne de données (pas l'en-tête)", vbExclamation
        Exit Sub
    End If
    
    ' Get numero de commande value (column {numero_commande_col})
    numCommande = Cells(currentRow, {numero_commande_col}).Value
    
    If numCommande = "" Then
        MsgBox "Aucun numéro de commande trouvé sur cette ligne", vbExclamation
        Exit Sub
    End If
    
    ' Get the entire row data
    Set rowData = Rows(currentRow)
    
    ' Call the processing function
    Call PACInterneModule.ProcessPACInterne(numCommande, rowData)
End Sub
'''
                    sheet_module.CodeModule.AddFromString(button_code)
                    button.OnAction = f"{sheet.Name}.{button.Name}_Click"
                    
                    print(f"Added PAC Interne button to sheet: {sheet.Name}")
            
            # Save as macro-enabled workbook
            wb.SaveAs(os.path.abspath(output_file), FileFormat=52)  # 52 = xlOpenXMLWorkbookMacroEnabled
            wb.Close()
            excel.Quit()
            
            print(f"\nExcel file with integrated PAC Interne functionality created: {output_file}")
            print("\nInstructions:")
            print("1. Open the Excel file")
            print("2. Enable macros when prompted")
            print("3. Select a row with data")
            print("4. Click the 'PAC Interne' button")
            print("5. The document will be generated automatically")
            
            return output_file
            
        except Exception as e:
            print(f"Error creating integrated Excel: {str(e)}")
            if 'excel' in locals():
                try:
                    excel.Quit()
                except:
                    pass
            return None
        finally:
            pythoncom.CoUninitialize()


def main():
    # Configuration
    input_excel = "combined_pdfs_with_PAC_button.xlsx"  # Your existing Excel file
    
    # Create integrated Excel file
    processor = PACInterneIntegrated()
    output_file = processor.create_excel_with_pac_macro(input_excel)
    
    if output_file:
        print(f"\n✅ Success! Created: {output_file}")
        print("\nThe Excel file now has PAC Interne buttons that will:")
        print("- Read data from the selected row")
        print("- Open the Word template")
        print("- Replace yellow highlighted text with row data")
        print("- Save the document with the numero de commande in the filename")
    else:
        print("\n❌ Failed to create integrated Excel file")


if __name__ == "__main__":
    main()
