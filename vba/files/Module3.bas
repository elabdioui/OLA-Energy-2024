Attribute VB_Name = "Module3"
Sub AddNewRecord()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim templateRow As Long
    Dim tableStartRow As Long
    Dim tableEndRow As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Movement Sheet")

    ' Define the start and end rows of the "Details du mouvement" table
    tableStartRow = 55
    tableEndRow = 75  ' Adjusted to the last row of your "Details du mouvement" table

    ' Find the last filled row within the table range in column B
    lastRow = ws.Cells(tableEndRow, "B").End(xlUp).Row

    ' Ensure the last row is within the "Details du mouvement" table range
    If lastRow < tableStartRow Or lastRow >= tableEndRow Then
        MsgBox "Not enough rows to determine a template or the table is full.", vbExclamation
        Exit Sub
    End If

    ' Determine the template row, which is the row two rows above the last filled row within the table range
    If lastRow >= tableStartRow + 1 Then
        templateRow = lastRow - 1
    Else
        MsgBox "Not enough rows to determine a template.", vbExclamation
        Exit Sub
    End If

    ' Insert a new row after the last filled row within the table range
    ws.Rows(lastRow + 1).Insert Shift:=xlDown

    ' Copy the template row (two rows above the last filled row) to the new row within the table range
    ws.Rows(templateRow).Copy
    ws.Rows(lastRow + 1).PasteSpecial Paste:=xlPasteFormats
    ws.Rows(lastRow + 1).PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False

    ' Clear values in the new row to maintain template consistency
    ws.Rows(lastRow + 1).SpecialCells(xlCellTypeConstants).ClearContents

    ' Number the new row starting from 3
    ws.Cells(lastRow + 1, 2).value = lastRow - tableStartRow + 2
End Sub
Sub InsertDataIntoDestination(wsSource As Worksheet, wsDestination As Worksheet, startRow As Long, endRow As Long)
    Dim lastRowSource As Long
    Dim lastRowDestination As Long
    Dim i As Long
    Dim nomComplet As String
    Dim articlesCount As Long
    Dim fa As Variant
    Dim typeMouvement As Variant
    Dim ficheMouvement As Variant
    Dim source As Variant
    Dim destination As Variant
    Dim article As Variant
    Dim quantite As Variant
    Dim marque As Variant
    Dim emballage As Variant

    ' Find the last used row in the source sheet's details table
    lastRowSource = wsSource.Cells(endRow, "B").End(xlUp).Row

    ' Initialize the last row destination, starting from line 2 in the destination sheet's table
    lastRowDestination = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Row
    If lastRowDestination < 2 Then
        lastRowDestination = 2
    ElseIf wsDestination.Cells(lastRowDestination, 1).value <> "" Then
        lastRowDestination = lastRowDestination + 1
    End If

    ' Loop through the source data and copy it to the destination sheet
    For i = startRow To lastRowSource
        ' Assign values to variables
        fa = wsSource.Range("H3").value
        typeMouvement = wsSource.Range("H4").value
        ficheMouvement = wsSource.Range("H5").value
        source = wsSource.Range("D18").value
        destination = wsSource.Range("H18").value
        article = wsSource.Cells(i, 3).value
        quantite = wsSource.Cells(i, 5).value
        marque = wsSource.Cells(i, 6).value
        emballage = wsSource.Cells(i, 4).value

        ' Check if required fields are not empty
        If (IsEmpty(fa) Or fa = "") Or _
           (IsEmpty(typeMouvement) Or typeMouvement = "") Or _
           (IsEmpty(ficheMouvement) Or ficheMouvement = "") Or _
           (IsEmpty(source) Or source = "") Or _
           (IsEmpty(destination) Or destination = "") Or _
           (IsEmpty(article) Or article = "") Or _
           (IsEmpty(quantite) Or quantite = "") Or _
           (IsEmpty(marque) Or marque = "") Or _
           (IsEmpty(emballage) Or emballage = "") Then

            ' Display error message and exit the sub
            MsgBox "Erreur : Un ou plusieurs champs requis sont vides. Veuillez vérifier les données et réessayer.", vbCritical
            Exit Sub
        End If

        ' Validate that quantite is a long integer
        If Not IsLong(quantite) Then
            MsgBox "Erreur : La quantité à la ligne " & i & " doit être un entier long.", vbCritical
            Exit Sub
        End If

        ' Calculate concatenated names for each row
        nomComplet = ""
        If wsSource.Range("D35").value <> "" Then
            nomComplet = Trim(wsSource.Range("D35").value)
        End If
        If wsSource.Range("D36").value <> "" Then
            If nomComplet <> "" Then
                nomComplet = nomComplet & " / " & Trim(wsSource.Range("D36").value)
            Else
                nomComplet = Trim(wsSource.Range("D36").value)
            End If
        End If

        ' Copy data to the destination sheet
        wsDestination.Cells(lastRowDestination, 1).value = fa ' FA
        wsDestination.Cells(lastRowDestination, 2).value = ficheMouvement ' Fiche de mouvement
        wsDestination.Cells(lastRowDestination, 3).value = typeMouvement ' Type de Mouvement
        wsDestination.Cells(lastRowDestination, 4).value = wsSource.Range("H6").value ' Date mouvement
        wsDestination.Cells(lastRowDestination, 5).value = source ' Source
        wsDestination.Cells(lastRowDestination, 6).value = destination ' Destination

        ' Copy repetitive data for each row
        wsDestination.Cells(lastRowDestination, 7).value = wsSource.Cells(i, 2).value ' Numero
        wsDestination.Cells(lastRowDestination, 8).value = article ' Article
        wsDestination.Cells(lastRowDestination, 9).value = emballage ' Emballage
        wsDestination.Cells(lastRowDestination, 10).value = quantite ' Quantite
        wsDestination.Cells(lastRowDestination, 11).value = marque ' Marque
        wsDestination.Cells(lastRowDestination, 12).value = wsSource.Cells(i, 7).value ' Commentaire

        ' Add additional fields for each article
        wsDestination.Cells(lastRowDestination, 13).value = wsSource.Range("B26").value ' Gestionnaire de stock
        wsDestination.Cells(lastRowDestination, 14).value = wsSource.Range("E26").value ' Agent de gardiennage
        wsDestination.Cells(lastRowDestination, 15).value = wsSource.Range("H26").value ' Transporteur

        ' Add concatenated recipient names
        wsDestination.Cells(lastRowDestination, 16).value = nomComplet ' Nom complet

        ' Increment the articles count
        articlesCount = articlesCount + 1

        ' Move to the next row in the destination sheet
        lastRowDestination = lastRowDestination + 1
    Next i

    MsgBox "Les données ont été transférées avec succès!"
End Sub

Function IsLong(value As Variant) As Boolean
    On Error Resume Next
    IsLong = CLng(value) = value
    On Error GoTo 0
End Function



Sub SaveMovementData()
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim startRow As Long
    Dim endRow As Long

    ' Define source and destination sheets
    Set wsSource = ThisWorkbook.Sheets("Movement Sheet")
    Set wsDestination = ThisWorkbook.Sheets("Mouvement Log")
    
    ' Validate dates
    If Not ValidateDates() Then
        Exit Sub
    End If
    
    ' Define start and end rows of the "Details du mouvement" table
    startRow = 55
    endRow = 75

    ' Insert data into destination sheet
    InsertDataIntoDestination wsSource, wsDestination, startRow, endRow

    ' Reset data in the source sheet
    ResetSourceData wsSource, startRow, endRow
End Sub
Sub ResetSourceData(wsSource As Worksheet, startRow As Long, endRow As Long)
    Dim i As Long

    ' Reset specific fields in the source sheet
    wsSource.Range("H3").value = ""  ' FA
    wsSource.Range("H4").value = ""  ' Fiche de mouvement
    wsSource.Range("H5").value = ""  ' Type de Mouvement
    wsSource.Range("H6").value = Date  ' Date actuelle
    wsSource.Range("D18").value = ""  ' Source
    wsSource.Range("H18").value = ""  ' Destination

    ' Reset repetitive data for each row
    For i = startRow To endRow
        wsSource.Cells(i, 3).value = ""  ' Article
        wsSource.Cells(i, 4).value = ""  ' Emballage
        wsSource.Cells(i, 5).value = ""  ' Quantite
        wsSource.Cells(i, 6).value = ""  ' Marque
        wsSource.Cells(i, 7).value = ""  ' Commentaire
    Next i
    ' Reset additional fields
    wsSource.Range("B26").value = ""  ' Gestionnaire de stock
    wsSource.Range("E26").value = ""  ' Agent de gardiennage
    wsSource.Range("H26").value = ""  ' Transporteur

    ' Reset concatenated names and merged cells
    wsSource.Range("D35").value = ""
    wsSource.Range("E35").value = ""
    wsSource.Range("F35:G35").value = ""
    
    wsSource.Range("D36").value = ""
    wsSource.Range("E36").value = ""
    wsSource.Range("F36:G36").value = ""
End Sub
Function ValidateDates() As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Movement Sheet")

    Dim startDate As Date
    Dim startDateValue As Variant
    Dim dateFormat As String

    ' Define the expected date format
    dateFormat = "dd/mm/yyyy"

    ' Initialize the function to False by default
    ValidateDates = False

    ' Retrieve the value from the merged cells
    startDateValue = ws.Range("H6").MergeArea.Cells(1, 1).value

    ' Check if the value is empty
    If IsEmpty(startDateValue) Then
        MsgBox "Erreur : Un ou plusieurs champs requis sont vides. Veuillez vérifier les données et réessayer.", vbCritical
        Exit Function
    End If

    ' Check if the value is a valid date
    If IsDate(startDateValue) Then
        startDate = CDate(startDateValue)

        ' Check if the date format matches the expected format
        If Format(startDate, dateFormat) = Format(startDateValue, dateFormat) Then
            ' If all checks pass, set the function to True
            ValidateDates = True
        Else
            MsgBox "La date n'est pas au format correct. Veuillez utiliser le format jj/mm/aaaa.", vbExclamation
        End If
    Else
        MsgBox "Veuillez entrer une date valide.", vbExclamation
    End If
End Function

