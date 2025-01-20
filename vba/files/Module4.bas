Attribute VB_Name = "Module4"
Sub SetDefaultValues()
    Dim wsDaily As Worksheet
    Set wsDaily = ThisWorkbook.Sheets("Daily inventory")

    ' Verify and set default date
    If Not ValidateDates() Then
        wsDaily.Range("B7").value = Date
        wsDaily.Range("B7").NumberFormat = "dd/mm/yyyy"
        MsgBox "La date était invalide. Elle a été définie à aujourd'hui : " & Format(Date, "dd/mm/yyyy"), vbInformation, "Information"
    End If

    ' Verify and set default time
    If Not ValidateTime() Then
        wsDaily.Range("B8").value = Format(Now, "HH:mm")
        wsDaily.Range("B8").NumberFormat = "HH:mm"
    End If
End Sub

Function ValidateDates() As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Daily inventory")
    Dim dateValue As Variant
    Dim dateValueInCell As Date
    
    ' Initialize the function to False by default
    ValidateDates = False

    ' Get the value of cell B7
    dateValue = ws.Range("B7").value
    
    ' Check if the value is a valid date and not in the future
    If IsDate(dateValue) Then
        dateValueInCell = CDate(dateValue)
        If dateValueInCell <= Date Then
            ValidateDates = True
        End If
    End If
End Function
Function ValidateTime() As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Daily inventory")
    Dim timeValue As Variant
    Dim timeValueInCell As Date
    
    ' Initialize the function to False by default
    ValidateTime = False

    ' Get the value of cell B8
    timeValue = ws.Range("B8").value
    
    ' Check if the value is a valid time
    If IsDate(timeValue) Then
        timeValueInCell = CDate(timeValue)
        ValidateTime = True
    End If
End Function
Sub Stocker()
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim tbl As ListObject
    Dim lastRowD As Long
    Dim newRow As ListRow
    Dim i As Long
    Dim cellValue As Variant
    Dim dataToSave As Boolean
    Dim currentTime As String
    Dim heureFormatted As String
    Dim dateValue As Variant
    Dim typeInventaire As Variant
    Dim articleValue As Variant
    Dim emballageValue As Variant
    Dim quantiteValue As Variant
    Dim regAdjustValue As Variant
    Dim marqueValue As Variant

    ' Set the source and destination worksheets
    Set wsSource = ThisWorkbook.Sheets("Daily inventory")
    Set wsDestination = ThisWorkbook.Sheets("Inventory Log")

    ' Set the table in the destination worksheet
    Set tbl = wsDestination.ListObjects("Tableau2")

    ' Find the last used row in column D of the source worksheet
    lastRowD = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row

    ' Ensure there is data to process
    If lastRowD < 12 Then
        MsgBox "Erreur : Pas de données à traiter à partir de la ligne D12.", vbInformation
        Exit Sub
    End If

    ' Initialize the dataToSave flag to False
    dataToSave = False

    ' Get the current time
    currentTime = Format(Now)

    ' Format the time in B8 to "HH:mm"
    heureFormatted = Format(wsSource.Range("B8").value, "HH:mm")

    ' Get other field values
    dateValue = wsSource.Range("B7").value
    typeInventaire = wsSource.Range("E8").value
    regAdjustValue = wsSource.Range("E7").value
    marqueValue = wsSource.Range("E6").value

    ' Check if critical fields are not empty
    If IsEmpty(dateValue) Or IsEmpty(typeInventaire) Or IsEmpty(regAdjustValue) Or IsEmpty(marqueValue) Then
        MsgBox "Erreur : Un ou plusieurs champs critiques sont vides. Veuillez vérifier les valeurs dans les cellules : Date (B7), Type d'inventaire (E8), Régulier/Ajustement (E7), ou Marque (E6).", vbCritical
        Exit Sub
    End If

    ' Loop through the values from D12 to the last used row in column D
    For i = 12 To lastRowD
        ' Get the value in column D
        cellValue = wsSource.Cells(i, "D").value
        
        ' Get other field values for the current row
        articleValue = wsSource.Cells(i, "A").value
        emballageValue = wsSource.Cells(i, "B").value
        quantiteValue = wsSource.Cells(i, "D").value

        ' Check if the value in column D is not 0 and other fields are not empty
        If cellValue <> 0 And Not IsEmpty(articleValue) And Not IsEmpty(emballageValue) And Not IsEmpty(quantiteValue) Then
            ' Set the dataToSave flag to True
            dataToSave = True
            
            ' Add a new row to the table
            Set newRow = tbl.ListRows.Add
            
            ' Populate the new row with data
            With newRow
                .Range(1, 1).value = dateValue ' Date from B7
                .Range(1, 2).value = heureFormatted ' Heure from B8
                .Range(1, 3).value = typeInventaire ' Type d'inventaire from E8
                .Range(1, 4).value = articleValue ' Article from A12 downwards
                .Range(1, 5).value = emballageValue ' Emballage from B12 downwards
                .Range(1, 6).value = quantiteValue ' Quantité from D12 downwards
                .Range(1, 7).value = regAdjustValue ' Régulier/Ajustement from E7
                .Range(1, 8).value = "hamid" ' Expéditeur (empty as per the provided mapping)
                .Range(1, 9).value = marqueValue ' Marque
                .Range(1, 10).value = currentTime ' Saisie (current time)
            End With
        End If
    Next i

    ' Check if no data was saved
    If Not dataToSave Then
        MsgBox "Erreur : Il n'y a pas de données à enregistrer. Toutes les valeurs de D12 à la dernière ligne sont égales à 0 ou des champs nécessaires sont vides.", vbCritical
        Exit Sub
    End If

    MsgBox "Les valeurs ont été enregistrées avec succès dans Tableau1."
End Sub


Function DateExiste(ws As Worksheet, dateInventaire As Date) As Boolean
    Dim lastRow As Long
    Dim i As Long
    Dim dateTrouvee As Boolean
    
    dateTrouvee = False
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If IsDate(ws.Cells(i, 1).value) Then
            If CDate(ws.Cells(i, 1).value) = dateInventaire Then
                dateTrouvee = True
                Exit For
            End If
        End If
    Next i
    
    DateExiste = dateTrouvee
End Function
Sub VerifierEtMettreAJourInventaire()
    Dim wsDaily As Worksheet
    Dim wsRecord As Worksheet
    Dim dateInventaire As Date
    
    ' Set the worksheets
    Set wsDaily = ThisWorkbook.Sheets("Daily inventory")
    Set wsRecord = ThisWorkbook.Sheets("Inventory record")
    
    ' Get the date
    dateInventaire = CDate(wsDaily.Cells(7, 2).value)
    
    ' Check if the date already exists in the "Inventory record" sheet
    If DateExiste(wsRecord, dateInventaire) Then
        wsDaily.Range("E7").value = "Ajustement"
    Else
        wsDaily.Range("E7").value = "Régulier"
    End If
End Sub
Sub ReinitialiserValeurs()
    Dim wsDaily As Worksheet
    Set wsDaily = ThisWorkbook.Sheets("Daily inventory")
    
    ' Reset date to today
    wsDaily.Range("B7").value = Date
    wsDaily.Range("B7").NumberFormat = "dd/mm/yyyy"
    
    ' Reset time to current time
    wsDaily.Range("B8").value = Format(Now, "HH:mm")
    wsDaily.Range("B8").NumberFormat = "HH:mm"
    
    ' Clear the inventory type field
    wsDaily.Range("E8").ClearContents
    
    ' Clear the quantity fields from row 12 to the last row
    Dim lastRowDaily As Long
    lastRowDaily = wsDaily.Cells(wsDaily.Rows.Count, "A").End(xlUp).Row
    wsDaily.Range("D12:D" & lastRowDaily).value = 0
End Sub


