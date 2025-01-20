VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    ' Hide all worksheets initially
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVeryHidden
    Next ws
    On Error GoTo 0
    
    ' Show the login form
    LoginForm.Show vbModal
    
    ' Check if authentication is successful
    If LoginForm.Tag = "Authenticated" Then
        ' Unhide the worksheets if authenticated
        For Each ws In ThisWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next ws
    Else
        ' Close the workbook if authentication fails
        ThisWorkbook.Close False
    End If
End Sub

