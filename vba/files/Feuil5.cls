VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ExportSheetAsPDF()
    Dim ws As Worksheet
    Dim pdfPath As String

    ' Set the worksheet you want to export
    Set ws = ThisWorkbook.Sheets("Mouvement Log") ' Change "Sheet1" to the name of your sheet
    ' Set the path where you want to save the PDF
    pdfPath = "C:\Users\hp\Desktop\" & ws.Name & ".pdf"

    ' Set the print area to cover the used range
    ws.PageSetup.PrintArea = ws.UsedRange.Address

    ' Fit the sheet to one page
    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    ' Export the sheet as a PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    ' Notify the user
    MsgBox "Worksheet has been exported as PDF: " & pdfPath
End Sub

