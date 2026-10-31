Sub PAC_Interne()
    ' PAC Interne - Procès-verbal d'acceptation interne
    ' This macro creates an acceptance certificate based on the selected order data
    
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim selectedRow As Long
    Dim fournisseur As String
    Dim numCommande As String
    Dim numContrat As String
    Dim dateLivraison As String
    Dim totalAmount As String
    
    ' Get the active worksheet
    Set ws = ActiveSheet
    
    ' Get selected row or use row 2 if nothing selected
    selectedRow = ActiveCell.Row
    If selectedRow < 2 Then selectedRow = 2
    
    ' Get data from the Excel sheet
    On Error Resume Next
    fournisseur = ws.Range("B1").Value
    numCommande = ws.Range("B2").Value
    numContrat = ws.Range("B3").Value
    dateLivraison = ws.Range("B4").Value
    
    ' Find the total amount
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 1 To lastRow
        If InStr(1, ws.Cells(i, 6).Value, "TOTAL", vbTextCompare) > 0 Then
            totalAmount = ws.Cells(i, 7).Value
            Exit For
        End If
    Next i
    
    On Error GoTo 0
    
    ' Create a new worksheet for the PAC
    Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newWs.Name = "PAC_" & Format(Now, "yyyymmdd_hhmmss")
    
    ' Format the PAC sheet
    With newWs
        ' Set column widths
        .Columns("A:H").ColumnWidth = 15
        .Columns("B").ColumnWidth = 30
        
        ' Title
        .Range("A1:H2").Merge
        .Range("A1").Value = "PROCÈS-VERBAL D'ACCEPTATION INTERNE"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16
        .Range("A1").HorizontalAlignment = xlCenter
        
        ' Add spacing
        .Range("A4").Value = "Informations du Bon de Commande:"
        .Range("A4").Font.Bold = True
        .Range("A4").Font.Size = 12
        
        ' Order details
        .Range("A6").Value = "Fournisseur:"
        .Range("B6").Value = fournisseur
        
        .Range("A7").Value = "N° Commande:"
        .Range("B7").Value = numCommande
        
        .Range("A8").Value = "N° Contrat:"
        .Range("B8").Value = numContrat
        
        .Range("A9").Value = "Date de livraison:"
        .Range("B9").Value = dateLivraison
        
        .Range("A10").Value = "Montant Total:"
        .Range("B10").Value = totalAmount
        
        ' Acceptance text
        .Range("A12:H14").Merge
        .Range("A12").Value = "Les services/équipements relatifs au contrat/BC/Facture/Période cités ci-dessus ont été bien réalisés selon les termes du bon de commande et/ou du contrat."
        .Range("A12").WrapText = True
        
        .Range("A16:H18").Merge
        .Range("A16").Value = "Ce PV a été dressé aux fins de preuve et de confirmation de la bonne réalisation/livraison/installation des dites prestations/équipements sous réserve."
        .Range("A16").WrapText = True
        
        ' Signature section
        .Range("A20").Value = "Cachet et signature:"
        .Range("A20").Font.Bold = True
        
        ' Add signature boxes
        .Range("A22").Value = "Pour le Fournisseur:"
        .Range("E22").Value = "Pour MEDI TELECOM:"
        
        ' Add lines for signatures
        .Range("A23:C24").Merge
        .Range("A23").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range("E23:G24").Merge
        .Range("E23").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        ' Bon à payer section
        .Range("A26").Value = "Constitue un « Bon à Payer » ?"
        .Range("B26").Value = "☐ Oui    ☐ Non"
        
        ' Date section
        .Range("A28").Value = "Date:"
        .Range("B28").Value = Format(Date, "dd/mm/yyyy")
        
        ' Format borders
        .Range("A1:H30").Borders.LineStyle = xlContinuous
        .Range("A1:H30").Borders.Weight = xlThin
        
        ' Center align some cells
        .Range("B6:B10").HorizontalAlignment = xlLeft
        
        ' Add some color
        .Range("A1:H2").Interior.Color = RGB(200, 200, 200)
        .Range("A4").Interior.Color = RGB(220, 220, 220)
        .Range("A20").Interior.Color = RGB(220, 220, 220)
        
    End With
    
    ' Activate the new sheet
    newWs.Activate
    
    MsgBox "PAC Interne créé avec succès!", vbInformation, "PAC Interne"
    
End Sub

Sub AddPACButton()
    ' This sub adds a button to run the PAC_Interne macro
    Dim btn As Button
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    
    ' Add button at the top right of the sheet
    Set btn = ws.Buttons.Add(500, 10, 120, 30)
    
    With btn
        .Name = "btnPACInterne"
        .Caption = "Générer PAC Interne"
        .OnAction = "PAC_Interne"
    End With
    
    MsgBox "Bouton PAC Interne ajouté!", vbInformation
End Sub
