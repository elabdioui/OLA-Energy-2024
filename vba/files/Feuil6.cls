VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Call ReinitialiserValeurs
End Sub
Private Sub CommandButton1_Click()
    Dim wsDaily As Worksheet
    Set wsDaily = ThisWorkbook.Sheets("Daily inventory")

    ' Verify and set default date if necessary
    If Not ValidateDates() Then
        wsDaily.Range("B7").value = Date
        wsDaily.Range("B7").NumberFormat = "dd/mm/yyyy"
        MsgBox "La date était invalide ou dans le futur. Elle a été définie à aujourd'hui : " & Format(Date, "dd/mm/yyyy"), vbInformation, "Information"
    End If

    ' Stocker les données et réinitialiser les champs
    Call Stocker

    Call ReinitialiserValeurs
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Vérifier si la cellule modifiée est B7
    If Not Intersect(Target, Me.Range("B7")) Is Nothing Then
        ' Appeler la fonction pour vérifier et mettre à jour l'inventaire
        VerifierEtMettreAJourInventaire
    End If
    If Not Intersect(Target, Me.Range("B7")) Is Nothing Then
        ValidateDates
    End If
End Sub
