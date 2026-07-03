Attribute VB_Name = "Module26"
Sub ExportNamePdf()
    Dim wbTarget As Workbook
    Set wbTarget = Workbooks.Open("C:\Users\lenovo\Desktop\ダウンロード\nameListTemp.xlsx")
    
    Dim ws As Worksheet
    Set ws = wbTarget.Sheets(1)
    
    Dim csvPath As String
    csvPath = ws.Range("B1").Value
    Dim baseName As String
    baseName = ws.Range("C1").Value
    
    ws.Columns("B:XFD").Delete
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    ws.Range("A1:A" & lastRow).Font.Size = 48
    ws.Columns("A").AutoFit
    
    Dim pdfPath As String
    pdfPath = Left(csvPath, InStrRev(csvPath, "\")) & baseName & "-name.pdf"
    
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard
    
    wbTarget.Close False
End Sub

