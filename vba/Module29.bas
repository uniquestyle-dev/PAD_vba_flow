Attribute VB_Name = "Module29"
Sub CountBookOrders(ByVal csvPath As String)
    
    On Error GoTo ErrHandler
    
    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim sku As String
    Dim carboffCount As Long, salasiruCount As Long
    Dim medicoCount As Long, aburaCount As Long
    Dim fso As Object, ts As Object
    Dim skuCol As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' CSVを正しくパースして開く
    Workbooks.OpenText fileName:=csvPath, _
        Origin:=65001, _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        Comma:=True
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets(1)
    
    ' ヘッダー行からLineitem sku列を自動検出
    skuCol = 0
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If LCase(Trim(CStr(ws.Cells(1, c).Value))) = "lineitem sku" Then
            skuCol = c
            Exit For
        End If
    Next c
    
    If skuCol = 0 Then
        MsgBox "エラー: Lineitem sku 列が見つかりません", vbCritical
        wb.Close False
        GoTo Cleanup
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, skuCol).End(xlUp).row
    
    carboffCount = 0
    salasiruCount = 0
    medicoCount = 0
    aburaCount = 0
    
    For i = 2 To lastRow
        sku = LCase(Trim(CStr(ws.Cells(i, skuCol).Value)))
        
        If InStr(sku, "01b") > 0 Or InStr(sku, "07b") > 0 Or InStr(sku, "cb") > 0 Then
            carboffCount = carboffCount + 1
        ElseIf InStr(sku, "ss") > 0 Then
            salasiruCount = salasiruCount + 1
        ElseIf InStr(sku, "mc") > 0 Then
            medicoCount = medicoCount + 1
        ElseIf InStr(sku, "abura") > 0 Then
            aburaCount = aburaCount + 1
        End If
    Next i
    
    wb.Close False
    
    ' --- 既存の books_cod_count.txt を読み込み（先ほどのマクロの出力。無ければ全0）---
    Dim countPath As String
    countPath = "C:\Users\lenovo\Desktop\ダウンロード\books_cod_count.txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim existing(1 To 5) As Long
    Dim k As Long
    For k = 1 To 5: existing(k) = 0: Next k
    
    If fso.FileExists(countPath) Then
        Dim tsr As Object, lineStr As String
        Set tsr = fso.OpenTextFile(countPath, 1)   ' 1 = ForReading
        k = 0
        Do While Not tsr.AtEndOfStream And k < 5
            lineStr = Trim(tsr.ReadLine)
            k = k + 1
            If IsNumeric(lineStr) Then existing(k) = CLng(Val(lineStr))
        Loop
        tsr.Close
        Set tsr = Nothing
    End If
    
    ' --- 合算して5行で書き戻し（1?4行目=今回の集計を加算 / 5行目=既存の代引き件数を保持）---
    Set ts = fso.CreateTextFile(countPath, True)
    ts.WriteLine existing(1) + carboffCount
    ts.WriteLine existing(2) + salasiruCount
    ts.WriteLine existing(3) + medicoCount
    ts.WriteLine existing(4) + aburaCount
    ts.WriteLine existing(5)
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    ' --- インポートした元CSVを別名コピー保存（orders_export_cod.csv）---
    Dim fsoCopy As Object
    Set fsoCopy = CreateObject("Scripting.FileSystemObject")
    fsoCopy.CopyFile csvPath, _
        "C:\Users\lenovo\Desktop\ダウンロード\orders_export_cod.csv", True
    Set fsoCopy = Nothing
Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "エラー: " & Err.Description, vbCritical
    
End Sub


