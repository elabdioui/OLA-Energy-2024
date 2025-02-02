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
        MsgBox "La date �tait invalide ou dans le futur. Elle a �t� d�finie � aujourd'hui : " & Format(Date, "dd/mm/yyyy"), vbInformation, "Information"
    End If

    ' Stocker les donn�es et r�initialiser les champs
    Call Stocker

    Call ReinitialiserValeurs
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    ' V�rifier si la cellule modifi�e est B7
    If Not Intersect(Target, Me.Range("B7")) Is Nothing Then
        ' Appeler la fonction pour v�rifier et mettre � jour l'inventaire
        VerifierEtMettreAJourInventaire
    End If
    If Not Intersect(Target, Me.Range("B7")) Is Nothing Then
        ValidateDates
    End If
End Sub
