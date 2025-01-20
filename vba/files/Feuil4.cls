VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CreateDropdownLists()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Inventory")

    ' Définir les listes d'options pour les magasins et LPG
    Dim magasins As String
    Dim lpg As String
    
    magasins = "Magasin 1,Magasin 2,Magasin 3"
    lpg = "LPG 1,LPG 2,LPG 3"
    
    ' Créer une liste déroulante pour les magasins dans la cellule C7
    With ws.Range("B8").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=magasins
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' Créer une liste déroulante pour les LPG dans la cellule D7
    With ws.Range("E8").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=lpg
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    
End Sub




