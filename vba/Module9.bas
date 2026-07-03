Attribute VB_Name = "Module9"
Sub サラシア数量_y()
    '
    ' サラシア数量 Macro
    ' Keyboard Shortcut: Ctrl+Shift+I
    '
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range

    ' ============================================
    ' 最新CSVファイルの自動取得＆オープン
    ' ============================================
    Dim folderPath As String
    Dim fileName As String
    Dim latestFile As String
    Dim latestDate As Date
    
    folderPath = "C:\Users\lenovo\Desktop\ダウンロード\"
    fileName = Dir(folderPath & "order_20*.csv")
    
    If fileName = "" Then
        MsgBox "order_20で始まるCSVファイルが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    Do While fileName <> ""
        If FileDateTime(folderPath & fileName) > latestDate Then
            latestDate = FileDateTime(folderPath & fileName)
            latestFile = fileName
        End If
        fileName = Dir()
    Loop
    
    Workbooks.Open folderPath & latestFile
    
    ' アクティブシートを選択し、名前を「1」に変更
    ActiveSheet.Select
    ActiveSheet.Name = "1"
    
    ' --- 不要な列の削除処理 ---
    Worksheets("1").Columns("A:A").Delete Shift:=xlToLeft
    Columns("B:F").Delete Shift:=xlToLeft
    Columns("D:D").Delete Shift:=xlToLeft
    
    ' アクティブシートを設定
    Set ws = ActiveSheet
    
    ' --- 最終行を取得（C列を基準） ---
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' --- C列のテキスト置換 ---
    For Each cell In ws.Range("C2:C" & lastRow)
        If cell.Value Like "*15％引き*" Or cell.Value Like "*20％引き*" Then
            cell.Value = "大"
        ElseIf cell.Value Like "*10％引き*" Or cell.Value Like "*割引なしの単品購入*" Then
            cell.Value = "小"
        End If
    Next cell
    
    ' --- A列とB列の2行目以下の背景色を白に設定 ---
    ws.Range("A2:B" & lastRow).Interior.Color = RGB(255, 255, 255)
    
    ' --- B列の購入数が2以上の場合、背景色を蛍光黄色に設定 ---
    For Each cell In ws.Range("B2:B" & lastRow)
        If IsNumeric(cell.Value) And cell.Value >= 2 Then
            cell.Interior.Color = RGB(255, 255, 0)
        End If
    Next cell

    ' --- C列の内容に応じた背景色の変更 ---
    For Each cell In ws.Range("C2:C" & lastRow)
        If cell.Value = "大" Then
            cell.Interior.Color = RGB(226, 239, 218)
        ElseIf cell.Value = "小" Then
            cell.Interior.Color = RGB(255, 255, 255)
        End If
    Next cell

    ' --- B列の左右に罫線を追加 ---
    With ws.Range("B2:B" & lastRow).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(226, 239, 218)
    End With
    With ws.Range("B2:B" & lastRow).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(226, 239, 218)
    End With

    ' ======================================================
    ' D列・E列の処理
    ' ======================================================
    ws.Columns("D:D").Copy Destination:=ws.Columns("E:E")
    
    ' ----- D列の修正処理（案内状★） -----
    ws.Range("D1").Value = "案内状★"
    
    Dim lastRowD As Long
    Dim cellD As Range
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).row
    For Each cellD In ws.Range("D2:D" & lastRowD)
        If cellD.Value = "コンビニ後払い-★" Then
            cellD.Value = "★"
        Else
            cellD.ClearContents
        End If
    Next cellD
    ws.Range("D2:D" & lastRowD).Interior.Color = RGB(255, 255, 255)
    With ws.Range("D2:D" & lastRowD).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(226, 239, 218)
    End With
    
    ' ----- E列（支払い方法）の処理 -----
    ws.Range("E1").Value = "支払い方法"
    With ws.Range("E1")
        .Interior.Color = RGB(112, 173, 71)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    Dim lastRowE As Long
    Dim cellE As Range, originalVal As String
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).row
    For Each cellE In ws.Range("E2:E" & lastRowE)
        originalVal = cellE.Value
        If InStr(originalVal, "-★") > 0 Then
            originalVal = Replace(originalVal, "-★", "")
        End If
        Select Case originalVal
            Case "コンビニ後払い", "コンビニ支払い-同梱"
                cellE.Value = "コンビニ"
            Case "代引き", "コンビニ"
                ' そのまま
            Case Else
                cellE.ClearContents
        End Select
    Next cellE
    ws.Range("E2:E" & lastRowE).Interior.Color = RGB(255, 255, 255)
    
    ' ----- E列の罫線 -----
    With ws.Range("E1:E" & lastRowE)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeLeft).Color = RGB(169, 208, 142)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeRight).Color = RGB(169, 208, 142)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeTop).Color = RGB(169, 208, 142)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeBottom).Color = RGB(169, 208, 142)
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideHorizontal).Color = RGB(169, 208, 142)
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideVertical).Color = RGB(169, 208, 142)
    End With
    
    ' --- 列幅の自動調整 ---
    ws.Rows("2:" & lastRow).Hidden = True
    ws.Columns("C:E").AutoFit
    ws.Rows("2:" & lastRow).Hidden = False
    ws.Columns("A:A").AutoFit
    
    ' --- 重複ヘッダー行の削除 ---
    Dim lastCol As Long, lastDataRow As Long
    Dim isHeader As Boolean
    Dim j As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastDataRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    For i = lastDataRow To 2 Step -1
        isHeader = True
        For j = 1 To lastCol
            If ws.Cells(i, j).Value <> ws.Cells(1, j).Value Then
                isHeader = False
                Exit For
            End If
        Next j
        If isHeader Then
            ws.Rows(i).Delete
        End If
    Next i
    
    Dim delRow As Long
    For delRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row To 2 Step -1
        If ws.Cells(delRow, "C").Value = ws.Cells(1, "C").Value Then
            ws.Rows(delRow).Delete
        End If
    Next delRow
    

    
End Sub
