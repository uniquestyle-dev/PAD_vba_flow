Attribute VB_Name = "Module28"
Sub CountOrdersByProduct(ByVal csvPath As String)

    On Error GoTo ErrHandler

    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim code As String
    Dim bigCount As Long, smallCount As Long
    Dim fso As Object, ts As Object

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wb = Workbooks.Open(csvPath)
    Set ws = wb.Sheets(1)

    lastRow = ws.Cells(ws.Rows.Count, 54).End(xlUp).row

    bigCount = 0
    smallCount = 0

    For i = 2 To lastRow
        code = LCase(Trim(CStr(ws.Cells(i, 54).Value)))

        If InStr(code, "3m") > 0 Or InStr(code, "12m") > 0 Then
            bigCount = bigCount + 1
        ElseIf InStr(code, "1m") > 0 Or InStr(code, "spot") > 0 Then
            smallCount = smallCount + 1
        End If
    Next i

    wb.Close False

    ' 結果をテキストファイルに出力
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile("C:\Users\lenovo\Desktop\ダウンロード\salacia_count.txt", True)
    ts.WriteLine smallCount
    ts.WriteLine bigCount
    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "エラー: " & Err.Description, vbCritical

End Sub
