VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "LoginForm"
   ClientHeight    =   7200
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5640
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ConnectButton_Click()
    Dim mail As String
    Dim password As String

    mail = userbox.Text
    password = passbox.Text

    Dim isValid As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Variables linked to user connected
    Dim logged_email As String
    Dim logged_password As String
    Dim logged_name As String
    Dim logged_prename As String
    Dim logged_post As String
    Dim logged_filliale As String

    isValid = False

    ' Open the external workbook
    Dim externalWorkbook As Workbook
    Set externalWorkbook = Workbooks.Open("C:\Users\pc\Desktop\vba\data.xlsx")

    ' Set the worksheet to the desired sheet in the external workbook
    Set ws = externalWorkbook.Sheets("Feuil1") ' Change "Feuil1" to the appropriate sheet name

    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Debug: Show last row number
    Debug.Print "Last Row: " & lastRow

    ' Loop through the mail and password ranges to check for valid credentials
    For i = 2 To lastRow
        ' Debug: Show current cell values being checked
        Debug.Print "Checking Mail: " & ws.Range("A" & i).value & " with Password: " & ws.Range("B" & i).value
        
        If mail = ws.Range("A" & i).value And password = ws.Range("B" & i).value Then
            isValid = True
            logged_email = ws.Range("A" & i).value
            logged_password = ws.Range("B" & i).value
            logged_name = ws.Range("C" & i).value
            logged_prename = ws.Range("D" & i).value
            logged_post = ws.Range("E" & i).value
            logged_filliale = ws.Range("F" & i).value
            Exit For
        End If
    Next i

    ' Close the external workbook without saving
    externalWorkbook.Close False

    ' Check if both mail and password are valid
    If isValid Then
        MsgBox "Authentification réussie. Bienvenue " & logged_name, vbInformation, "Bienvenue"
        Me.Tag = "Authenticated"
        
        ' Set the value of the rectangle form (profile_text)
        Dim fullname_text As Object
        Set fullname_text = Sheets("Menu").Shapes("fullname_text")
        fullname_text.TextFrame.Characters.Text = logged_name & " " & logged_prename
        
        Dim profile_text As Object
        Set profile_text = Sheets("Menu").Shapes("profile_text")
        profile_text.TextFrame.Characters.Text = logged_name
        
        Dim type_text As Object
        Set type_text = Sheets("Menu").Shapes("type_text")
        type_text.TextFrame.Characters.Text = logged_post
        
        Dim magasin_text As Object
        Set magasin_text = Sheets("Menu").Shapes("magasin_text")
        magasin_text.TextFrame.Characters.Text = logged_filliale
        
        ' Activate the Menu sheet
        Sheets("Menu").Activate

        ' Unhide all worksheets, including the mouvement log
        For Each ws In ThisWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next ws
        
        Me.Hide
    Else
        MsgBox "Nom d'utilisateur ou mot de passe incorrect", vbCritical, "Impossible de se connecter"
    End If
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

