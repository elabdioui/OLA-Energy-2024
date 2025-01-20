VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim movementCounters As Object ' Scripting.Dictionary

Sub InitializeCounters()
    ' Initialize the Dictionary
    Set movementCounters = CreateObject("Scripting.Dictionary")
    
    ' Define the movement types and counters
    movementCounters.Add "TRFEMB", 0
    movementCounters.Add "EXPEMB", 0
    movementCounters.Add "LIVPRD", 0
    movementCounters.Add "RCPPRD", 0
    movementCounters.Add "RCPEMB", 0
    movementCounters.Add "RCPACC", 0
    movementCounters.Add "EXPACC", 0
    movementCounters.Add "TRFACC", 0
End Sub
Private Sub Worksheet_Activate()
    ' Create the list of articles for C55 and C56
    Dim articlesC55 As Variant
    Dim articlesC56 As Variant
    Dim rowNum As Long
    Dim categories As Variant
    
    categories = Array("StationDeService", "ClientB2B", "MagasinGPL", "DépôtGPL", "Fournisseur", "Prestataire")
    
    
    articles = Array("Bouteille pleine", "Bouteille vide en service", "Bouteille vide en consignation", _
                        "Bouteille fuyarde", "Bouteille à ré éprouver", "Bouteille à entretenir", "Bouteille à détruire", _
                        "Casier en bon état", "Casier à entretenir", "Casier à réformer", _
                        "Robinet en bon état", "Robinet défectueux", _
                        "Boîte à clapet en bon état", "Boîte à clapet défectueux", _
                        "Ecrou", "Joint d'écrou", _
                        "Socle", "Ance")

    
    ' Set up the data validation list for cell C55
    With Range("C55").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=Join(articles, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' Set up the data validation list for cell C56
    With Range("C56").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=Join(articles, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    
    ' Set up the data validation list for cell D16 using categories
    With Range("D16").Validation
        .Delete ' Clear any existing validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=Join(categories, ",")
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' Update data validation in column D for rows 55 to 75
    For rowNum = 55 To 75
        UpdateDColumn rowNum
    Next rowNum
End Sub


Sub IncrementCounter(shortName As String)
    ' Ensure that movementCounters is initialized
    If movementCounters Is Nothing Then
        InitializeCounters
    End If
    
    ' Increment the counter
    If movementCounters.Exists(shortName) Then
        movementCounters(shortName) = movementCounters(shortName) + 1
    End If
End Sub

Sub SetValidationList()
    Dim ws As Worksheet
    Dim validationArray As Variant
    
    ' Define the two-dimensional array for mapping
    Dim nameMapping As Variant
    nameMapping = Array( _
        Array("Transfert d'emballage", "TRFEMB"), _
        Array("Expédition d'emballage", "EXPEMB"), _
        Array("Livraison produit", "LIVPRD"), _
        Array("Réception produit", "RCPPRD"), _
        Array("Réception d'emballage", "RCPEMB"), _
        Array("Réception d'accesoire", "RCPACC"), _
        Array("Expédition d'accesoire", "EXPACC"), _
        Array("µ", "TRFACC") _
    )
    
    ' Set the worksheet where you want to apply the validation
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Movement Sheet") ' Change to your actual sheet name
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "The specified sheet does not exist.", vbExclamation
        Exit Sub
    End If
    
    ' Define the validation array
    validationArray = Application.Transpose(Application.Index(nameMapping, 0, 1))
    
    ' Apply data validation to cell H4
    With ws.Range("H4").Validation
        .Delete ' Clear any existing validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="" & Join(validationArray, ",") & ""
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim nameMapping As Variant
    Dim i As Integer
    Dim selectedName As String
    Dim shortName As String
    Dim counter As Integer
    Dim rowNum As Long

    ' Check if the change happened in column C (within the range of 55 to 75)
    If Not Intersect(Target, Me.Range("C55:C75")) Is Nothing Then
        ' Loop through all rows in the intersected range
        For Each cell In Target
            If cell.Column = 3 And cell.Row >= 55 And cell.Row <= 75 Then
                rowNum = cell.Row
                ' Call UpdateDColumn for the changed row
                UpdateDColumn rowNum
            End If
        Next cell
    End If

    ' Define the two-dimensional array for mapping
    nameMapping = Array( _
        Array("Transfert d'emballage", "TRFEMB"), _
        Array("Expédition d'emballage", "EXPEMB"), _
        Array("Livraison produit", "LIVPRD"), _
        Array("Réception produit", "RCPPRD"), _
        Array("Réception d'emballage", "RCPEMB"), _
        Array("Réception d'accesoire", "RCPACC"), _
        Array("Expédition d'accesoire", "EXPACC"), _
        Array("µ", "TRFACC") _
    )
    
    ' Check if the change happened in cell H4
    If Not Intersect(Target, Me.Range("H4")) Is Nothing Then
        selectedName = Trim(Me.Range("H4").value)
        
        ' Loop through the array to find the corresponding short name
        For i = LBound(nameMapping) To UBound(nameMapping)
            If Trim(nameMapping(i)(0)) = selectedName Then
                shortName = nameMapping(i)(1)
                IncrementCounter shortName
                counter = movementCounters(shortName)
                Me.Range("H5").value = shortName & counter ' Update H5 with type and counter
                Exit For
            End If
        Next i
    End If

    ' Check if the change is in cell D16
    If Not Intersect(Target, Me.Range("D16")) Is Nothing Then
        ' Clear D17
        Me.Range("D17").value = ""
        ' Call the function to update D17 based on D16
        UpdateD17BasedOnD16
    End If
    
End Sub
Sub UpdateDColumn(rowNum As Long)
    Dim ws As Worksheet
    Dim selectedValue As String
    Dim emballageOptions As Variant
    Dim validationList As String
    Dim previousValue As String

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Movement Sheet")

    ' Ensure rowNum is within the valid range (55 to 75)
    If rowNum < 55 Or rowNum > 75 Then Exit Sub

    ' Get the previous value from column D
    previousValue = ws.Cells(rowNum, "D").value

    ' Get the selected value from the corresponding cell in column C
    selectedValue = ws.Cells(rowNum, "C").value

    ' Reset the value in column D only if the selected value in column C has changed
    If selectedValue <> previousValue Then
        ws.Cells(rowNum, "D").ClearContents
    End If

    ' Determine the specific type of emballage options based on the selected value in column C
    
    Select Case selectedValue
    
        Case "Bouteille pleine", "Bouteille vide en service", "Bouteille vide en consignation", _
             "Bouteille fuyarde", "Bouteille à ré éprouver", "Bouteille à entretenir", "Bouteille à détruire"
            emballageOptions = Array("Bouteille 6KG", "Bouteille 12.5KG", "Bouteille 35KG")
            
        Case "Casier en bon état", "Casier à entretenir", "Casier à réformer"
            emballageOptions = Array("Casier 6KG", "Casier 12.5KG", "Casier 35KG")
            
        Case "Robinet en bon état", "Robinet défectueux"
            emballageOptions = Array("Robinet")
         If UBound(emballageOptions) = 0 Then ws.Cells(rowNum, "D").value = emballageOptions(0)
            
        Case "Boîte à clapet en bon état", "Boîte à clapet défectueux"
            emballageOptions = Array("Boîte à clapet")
         If UBound(emballageOptions) = 0 Then ws.Cells(rowNum, "D").value = emballageOptions(0)
            
        Case "Ecrou"
            emballageOptions = Array("Ecrou")
         If UBound(emballageOptions) = 0 Then ws.Cells(rowNum, "D").value = emballageOptions(0)
            
        Case "Joint d'écrou"
            emballageOptions = Array("Joint d'écrou")
         If UBound(emballageOptions) = 0 Then ws.Cells(rowNum, "D").value = emballageOptions(0)
            
        Case "Socle"
            emballageOptions = Array("Socle")
         If UBound(emballageOptions) = 0 Then ws.Cells(rowNum, "D").value = emballageOptions(0)
            
        Case "Ance"
            emballageOptions = Array("Ance")
         If UBound(emballageOptions) = 0 Then ws.Cells(rowNum, "D").value = emballageOptions(0)
            
        Case Else
            emballageOptions = Array("")
            
    End Select


    ' Join the array into a comma-separated string
    If UBound(emballageOptions) >= 0 Then
        validationList = Join(emballageOptions, ",")
    Else
        validationList = ""
    End If


    ' Ensure the validation list is not empty
    If Len(validationList) > 0 Then
        ' Set up the data validation list for the cell in column D for the specific row
        With ws.Cells(rowNum, "D").Validation
            .Delete ' Clear any existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="" & validationList & ""
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Else
        ' Debug information if the list is empty
        Debug.Print "Validation list is empty for selected value: " & selectedValue
    End If
End Sub
Sub UpdateD17BasedOnD16()
    Dim ws As Worksheet
    Dim selectedType As String
    Dim optionsList As Variant
    Dim validationList As String
    Dim previousValue As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Movement Sheet")

    ' Get the selected value from D16
    selectedType = ws.Cells(16, "D").value
    
    ' Determine the specific options based on the selected type in D16
    Select Case selectedType
        Case "MagasinGPL"
            optionsList = Array("Magasin GPL BAFOUSSAM Code N°425", "Magasin GPL BASSA Code N°415", _
                                "Magasin GPL BERTOUA Code N°435", "Magasin GPL GAROUA Code N°445", _
                                "Magasin GPL YAOUNDE Code N°465", "Consignation Interne", "Déclassification Interne")
        Case "DépôtGPL"
            optionsList = Array("BOCOM GAZ DOUALA Dépôt N°498", "BOCOM GAZ YAOUNDE Dépôt N°499", _
                                "SCDP GAZ BAFOUSSAM Dépôt N°428", "SCDP GAZ BERTOUA Dépôt N°478", _
                                "SCDP GAZ DOUALA BONABERI Dépôt N°418", "SCDP GAZ MAROUA Dépôt N°448", _
                                "SCDP GAZ Ngaoundéré Dépôt N°458", "SCDP GAZ YAOUNDE Dépôt N°468")
        Case "ClientB2B"
            optionsList = Array("ABEREC LTD", "ADAX PETROLEUM", "AFRICODIS", "Akam And Bobvalla Company Limited (A&B Co Ltd)", _
                                "ALIMENTATION NONO", "BALUCO", "BISCUITERIE SAMUEL FOYOU", "Blessing Petroleum SA", _
                                "CCRC", "CHOCOCAM", "Chris & Co. Lubrigood", "CROSS MOTO Sarl", "DAGHELA GUSTAVE", _
                                "DESTIN DU NOUN", "ERNAD SARL", "ETA- AFRIQUE SARL", "ETS  M.D. GAZ DISTRIBUTION", _
                                "ETS ABOUBAKAR OUSMANOU", "ETS ALIOU IBRAHIM", "ETS ASFESO", "ETS ATCHOUDA TCHEBEBIA ECLADOR", _
                                "ETS BANDJA MENAGERE", "ETS BELLYNN", "ETS BENIGAZ", "ETS BOUCHEU", "ETS BRAHIM SALI", _
                                "ETS BUTA GAZ.", "ETS DEULAH", "ETS DJOUFFO ET GUY.", "ETS ESPACE PLANETE ELECTRONIQUE BAFOUSSAM", _
                                "ETS FAKAT", "ETS FEMMES DYNAMIQUES", "ETS FOGUIENG MELO DUPLEX", "ETS FOKOU GRACE A GOD", _
                                "ETS FRUITIER GAZ", "ETS K.D GAZ", "ETS K.M. GAZ", "ETS KONDEM GABRIEL", "ETS LA CONFIANCE", _
                                "ETS LA MAISON DU GAZ", "ETS LIPAL", "ETS M2A service Mohamadou Awalou", "ETS MATHURINO.", _
                                "ETS MFEPIT SALIFOU", "ETS MODA", "ETS MOHAMADOU BACHIROU", "ETS MUFOR", "ETS NANA DAIROU", _
                                "ETS NGAMI GAZ", "ETS PASCALO AND DONALD", "ETS PUMFA.", "ETS SOCADIS", "ETS SONFACK TATANFACK HERVE", _
                                "ETS SONOBA", "ETS TADZUGUE", "ETS TAYO", "ETS TOUSSIDES & FILS", "ETS VINA G", "ETS YIMENE ALAIN CLOVIS", _
                                "ETS ZEMMO", "FBS", "FUTHER MARKET CAMEROUN SA", "GERANCE VICTORIA", "GESC SARL", "GLOBAL GAZ SERVICES YDE", _
                                "GLOBAL VOYAGES Sarl", "GROUPE TOFEUTOUO SARL.", "HOME DISTRIBUTION SERVICES", "HOTEL SAVANA", _
                                "JOCKY S/S GAZ YAOUNDE", "KAJERI ROBINSON NKAMTA", "KAMGUE TOGUE EPSE TOUOLAC", "KENNE TSAGUE CARLOS CHARENT", _
                                "KENTSACA DISTRIBUTEUR SARL", "LA BOUTIQUE DU GAZ", "LE GRAND MOULIN", "LONTSIE RIGOBERT", "MENGEL France", _
                                "MIAH SARL", "MOUMI NUFI", "Neptune Oil SA", "NJOPMO FIDELE", "PHARMEDI CO", "PRESTATAIRE", "PROMETAL", _
                                "PROMO GAZ ENTREPRISE", "PROPALM BOIS", "REHOBOT", "RETRAM TRADING", "SAFACAM", "SANDI - GAZ SARL", "SCTL", _
                                "SEBC", "SOCABANE S", "SOCAPURSEL", "SOROUBAT", "SOCIETE HELIOS SARL", "SOCIETE SOTRADI SARL", "SODISCO", _
                                "SOTRACO SARL", "STE FOKOUA SARL", "STE STRACO", "T.T.M S Sarl", "TCHUIKEP Epse NKAMGNIA Julienne (LA PETITE MENAGERE)", _
                                "UBC", "UNIVERSAL GAS LIMITED")
        Case "StationDeService"
            optionsList = Array("AHALA", "AXE LOURD 1", "AXE LOURD 2", "BAFANG", "BAFIA", "BANDJOUN", "BANYO", "BASSA", "BASTOS", "BEACH", _
                                "BEKA", "BELABO", "BEPANDA", "BERTOUA", "BIYEM-ASSI", "BONABERI", "BONAKOUAMOUANG", "BONAMOUSSADI", _
                                "BONENDALE", "BONIS BERTOUA GD", "BRAZZAVILLE", "BUEA", "CANTON BELL", "COMMERCIAL AVENUE", "DJOUNGOLO", _
                                "DSCHANG", "EBOLOWA", "EDEA", "EMANA", "ESSOMBA", "ESSOS", "GARE ROUTIERE", "GAROUA BOULAI", "KAELE", _
                                "KONDENGUI", "KOUME BERTOUA", "KUMBA", "LECLERC", "LIMBE", "LYCEE", "MADAGASCAR", "MANJO", "MAROUA", _
                                "MAROUA Djarengol", "MBALMAYO", "MBANKOUOP", "MBOUDA", "MEIGANGA", "MESSASSI", "MISSOLE", "MUTENGENE", _
                                "MVAN MAGZI", "MVOG-ADA", "MVOLYE", "NASSARAO", "NDIENGSO", "NDOKOTI", "NEW - MESSA", "NEW DEIDO", "NGEME LIMBE", _
                                "NJO NJO", "NKOLBISSON", "NKWEN", "NOUVEL AEROPORT", "NYALLA", "OBALA", "OMNISPORT", "PK 11 BASSA", "ROND POINT EXPRESS", _
                                "ROUND ABOUT", "ROUTE FOUMBOT", "ROUTE GARE", "SANGMELIMA", "SANTA BARBARA", "Station test1", "TROPICANA", "UNIVERSITE", _
                                "UNIVERSITE DANG", "YELWA GD")
        Case "Fournisseur"
            optionsList = Array("AMTROL Portugal", "EVAS TURQUIE", "MULTIPLAST-CAM", "SMPC THAILAND")
        Case "Prestataire"
            optionsList = Array("CAMEROON ALERT SERVICE", "ENERGYZ", "HYDRAC", "PETRO-SERVICE", "SGS", "ZALI SARL")
        Case Else
            optionsList = Array("")
    End Select
    
    ' Join the array into a comma-separated string
    validationList = Join(optionsList, ",")
    
    ' Ensure the validation list is not empty
    If Len(validationList) > 0 Then
        ' Set up the data validation list for the cell in column D for the specific row
        With ws.Cells(17, "D").Validation
            .Delete ' Clear any existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="" & validationList & ""
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Else
        ' Debug information if the list is empty
        Debug.Print "Validation list is empty for selected value: " & selectedValue
    End If
End Sub






