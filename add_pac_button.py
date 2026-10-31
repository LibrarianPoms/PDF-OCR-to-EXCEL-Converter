import os
import sys
import win32com.client
import pythoncom

def add_pac_button_to_excel(input_file):
    """Add PAC Interne button to an existing Excel file"""
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Get absolute path
        abs_path = os.path.abspath(input_file)
        if not os.path.exists(abs_path):
            print(f"Error: File not found: {input_file}")
            return None
            
        # Open workbook
        wb = excel.Workbooks.Open(abs_path)
        
        # Create output filename
        base_name = os.path.splitext(input_file)[0]
        output_file = f"{base_name}_with_PAC.xlsm"
        
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
                    Case "MONTANT_TOTAL", "MONTANT TOTAL", "TOTAL"
                        cellValue = GetCellValue(rowData, headers, "Total")
                        If cellValue = "" Then
                            cellValue = GetCellValue(rowData, headers, "Total (DH)")
                        End If
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
                                Case "MONTANT_TOTAL", "MONTANT TOTAL", "TOTAL"
                                    tableCellValue = GetCellValue(rowData, headers, "Total")
                                    If tableCellValue = "" Then
                                        tableCellValue = GetCellValue(rowData, headers, "Total (DH)")
                                    End If
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
        
        # Add button to each sheet
        for sheet in wb.Sheets:
            # Add button at the top of each sheet
            button = sheet.Buttons().Add(Left=500, Top=10, Width=120, Height=30)
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
    
    ' Try to find numero de commande in different possible locations
    numCommande = ""
    
    ' First check if it's in the sheet name
    If numCommande = "" Then
        numCommande = ActiveSheet.Name
    End If
    
    ' Then check if it's in the header area (rows 1-5)
    If numCommande = "" Or Len(numCommande) < 5 Then
        Dim i As Integer, j As Integer
        For i = 1 To 5
            For j = 1 To 10
                If InStr(1, LCase(Cells(i, j).Value), "commande") > 0 And j < 10 Then
                    If Cells(i, j + 1).Value <> "" Then
                        numCommande = CStr(Cells(i, j + 1).Value)
                        Exit For
                    End If
                End If
            Next j
            If numCommande <> "" Then Exit For
        Next i
    End If
    
    ' If still not found, check the current row for a numeric value that could be a command number
    If numCommande = "" Then
        For j = 1 To 10
            If IsNumeric(Cells(currentRow, j).Value) And Len(CStr(Cells(currentRow, j).Value)) > 5 Then
                numCommande = CStr(Cells(currentRow, j).Value)
                Exit For
            End If
        Next j
    End If
    
    If numCommande = "" Then
        MsgBox "Aucun numéro de commande trouvé", vbExclamation
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
        
        print(f"\n✅ Success! Created: {output_file}")
        print("\nThe PAC Interne button has been added to all sheets.")
        print("To use it:")
        print("1. Open the new Excel file")
        print("2. Enable macros when prompted")
        print("3. Select a row with article data")
        print("4. Click the 'PAC Interne' button")
        
        return output_file
        
    except Exception as e:
        print(f"Error adding PAC button: {str(e)}")
        if 'excel' in locals():
            try:
                excel.Quit()
            except:
                pass
        return None
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    # Check if filename was provided as argument
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        # Use the most recent Excel file from the app
        input_file = "combined_pdfs_20250806_150107.xlsx"
    
    result = add_pac_button_to_excel(input_file)
    if result:
        print(f"\nYou can now open: {result}")
    else:
        print("\nFailed to add PAC button to the Excel file")
