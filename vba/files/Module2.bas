Attribute VB_Name = "Module2"





Sub buttonInventory_Click()
    Dim datesValid As Boolean
    
    ' Appeler la fonction de validation des dates et stocker le r�sultat
    datesValid = ValidateDates
    
    ' V�rifier si les dates sont valides avant de continuer
    If Not datesValid Then
        ' Si les dates ne sont pas valides, arr�ter le traitement
        Exit Sub
    End If

    ' Si les dates sont valides, appeler la fonction pour sauvegarder les donn�es
    Call SaveAllData
    
    Call Workbook_Open
End Sub

Sub SaveAllData()
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim tbl As ListObject
    Dim lastRow As ListRow
    Dim srcRow As Long
    Dim srcLastRow As Long
    Dim allZero As Boolean
    Dim i As Integer

    On Error GoTo ErrorHandler

    ' D�finir les feuilles de calcul
    Set ws = ThisWorkbook.Sheets("Inventory")
    Set destWs = ThisWorkbook.Sheets("Inventory Log")

    ' D�finir le tableau dans la feuille de destination
    Set tbl = destWs.ListObjects("Tableau4")

    ' Trouver la derni�re ligne utilis�e dans la feuille source
    srcLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If srcLastRow < 12 Then
        MsgBox "Aucune donn�e � partir de la ligne 12 dans la feuille source.", vbExclamation
        Exit Sub
    End If

    ' V�rifier si toutes les lignes et leurs champs sont �gaux � 0
    allZero = True
    For srcRow = 12 To 49
        If ws.Cells(srcRow, 3).value <> 0 Or _
           ws.Cells(srcRow, 4).value <> 0 Or _
           ws.Cells(srcRow, 5).value <> 0 Or _
           ws.Cells(srcRow, 6).value <> 0 Or _
           ws.Cells(srcRow, 7).value <> 0 Or _
           ws.Cells(srcRow, 8).value <> 0 Then
            allZero = False
            Exit For
        End If
    Next srcRow

    If allZero Then
        MsgBox "La feuille enti�re est vide (� partir de la ligne 12).", vbExclamation
        Exit Sub
    End If

    ' Ajouter des lignes au tableau
    For srcRow = 12 To srcLastRow
        ' V�rifier si les valeurs sont valides et ne contiennent pas "-"
        If ws.Cells(srcRow, 1).value <> "-" And ws.Cells(srcRow, 1).value <> "" And _
           ws.Cells(srcRow, 2).value <> "-" And ws.Cells(srcRow, 2).value <> "" And _
           ws.Range("B7").value <> "-" And ws.Range("B7").value <> "" And _
           ws.Range("E7").value <> "-" And ws.Range("E7").value <> "" And _
           ws.Range("B8").value <> "-" And ws.Range("B8").value <> "" And _
           ws.Range("E8").value <> "-" And ws.Range("E8").value <> "" Then
           
            ' V�rifier si tous les champs (� l'exception de Date, Magasin et LPG) sont � 0
            If ws.Cells(srcRow, 3).value <> 0 Or _
               ws.Cells(srcRow, 4).value <> 0 Or _
               ws.Cells(srcRow, 5).value <> 0 Or _
               ws.Cells(srcRow, 6).value <> 0 Or _
               ws.Cells(srcRow, 7).value <> 0 Or _
               ws.Cells(srcRow, 8).value <> 0 Then
               
                Set lastRow = tbl.ListRows.Add
                
                ' Enregistrer les donn�es dans la ligne ajout�e
                lastRow.Range(1, 1).value = ws.Cells(srcRow, 1).value ' Article
                lastRow.Range(1, 2).value = ws.Cells(srcRow, 2).value ' Emballage
                lastRow.Range(1, 3).value = ws.Range("B7").value ' Date de d�but
                lastRow.Range(1, 4).value = ws.Range("E7").value ' Date de fin
                lastRow.Range(1, 5).value = ws.Range("B8").value ' Magasin
                lastRow.Range(1, 6).value = ws.Range("E8").value ' Nom du magasin LPG
                lastRow.Range(1, 7).value = ws.Cells(srcRow, 4).value ' Inventaire physique d'ouverture
                lastRow.Range(1, 8).value = ws.Cells(srcRow, 5).value ' Mouvements de r�ception
                lastRow.Range(1, 9).value = ws.Cells(srcRow, 6).value ' Mouvements de sortie
                lastRow.Range(1, 10).value = ws.Cells(srcRow, 7).value ' Inventaire physique de cl�ture
                lastRow.Range(1, 11).value = ws.Cells(srcRow, 8).value ' Inventaire th�orique de cl�ture
                lastRow.Range(1, 12).value = ws.Cells(srcRow, 9).value ' �cart
            End If
        End If
    Next srcRow

    MsgBox "Toutes les donn�es ont �t� enregistr�es avec succ�s.", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Une erreur s'est produite : " & Err.description, vbCritical
End Sub
Function ValidateDates() As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Inventory")

    Dim startDate As Date
    Dim endDate As Date
    Dim startDateValue As Variant
    Dim endDateValue As Variant
    
    ' Initialiser la fonction � False par d�faut
    ValidateDates = False

    ' R�cup�rer les valeurs des cellules
    startDateValue = ws.Range("B7").value
    endDateValue = ws.Range("E7").value
    
    ' V�rifier si les valeurs sont des dates valides
    If IsDate(startDateValue) Then
        startDate = CDate(startDateValue)
    Else
        MsgBox "Veuillez entrer une date de d�but valide au format jj/mm/aaaa.", vbExclamation
        Exit Function
    End If
    
    If IsDate(endDateValue) Then
        endDate = CDate(endDateValue)
    Else
        MsgBox "Veuillez entrer une date de fin valide au format jj/mm/aaaa.", vbExclamation
        Exit Function
    End If

    ' V�rifier si la date de fin est post�rieure ou �gale � la date de d�but
    If endDate < startDate Then
        MsgBox "La date de fin doit �tre post�rieure ou �gale � la date de d�but.", vbExclamation
        Exit Function
    End If

    ' Si toutes les v�rifications sont pass�es, d�finir la fonction � True
    ValidateDates = True
End Function
Private Sub Workbook_Open()
    ' D�finir des valeurs par d�faut pour certaines cellules
    With ThisWorkbook.Sheets("Inventory")
        ' D�finir des valeurs sp�cifiques
        .Range("B8").value = " "
        .Range("E8").value = " "
        .Range("E7").value = Date ' Date actuelle
        .Range("B7").value = Date

        ' D�finir des valeurs par d�faut � 0 pour une plage de cellules
        .Range("C10:C49").value = 0 ' Ajustez la plage de cellules selon vos besoins
        .Range("D10:D49").value = 0
        .Range("E10:E49").value = 0
        .Range("F10:F49").value = 0
        .Range("G10:G49").value = 0
        .Range("H10:H49").value = 0
        .Range("I10:I49").value = 0
    End With
End Sub
