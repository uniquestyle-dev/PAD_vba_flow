Attribute VB_Name = "Module33"
Sub TransferTrackingNumberSalacia(ByVal orderPath As String, ByVal daibikiPath As String, ByVal daibikiPath2 As String)
    On Error GoTo ErrHandler
    Dim wbOrder As Workbook, wbDaibiki As Workbook
    Dim wsOrder As Worksheet, wsDaibiki As Worksheet
    Dim lastRowDaibiki As Long, lastRowOrder As Long
    Dim i As Long, j As Long
    Dim targetName As String, orderName As String, trackingNum As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wbOrder = Workbooks.Open(orderPath)
    Set wsOrder = wbOrder.Sheets(1)
    
    ' --- ガード: 処理済み/想定外レイアウトの再処理を防ぐ ---
    If wsOrder.Cells(1, 10).Value <> "配送伝票番号" Then
        wbOrder.Close False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        MsgBox "入力が想定レイアウト(10列目=配送伝票番号)ではありません。" & vbCrLf & _
               "処理済みファイルの再実行の可能性があるため中止しました。", vbExclamation
        Exit Sub
    End If
    
    lastRowOrder = wsOrder.Cells(wsOrder.Rows.Count, 1).End(xlUp).row
    
    ' --- 1つ目: merged_tracking_book.csv ---
    Workbooks.OpenText fileName:=daibikiPath, _
        DataType:=xlDelimited, Comma:=True, _
        FieldInfo:=Array(Array(1, 2), Array(2, 1), Array(3, 1), Array(4, 1), _
                         Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
                         Array(9, 1), Array(10, 1), Array(11, 2), Array(12, 1), _
                         Array(13, 1), Array(14, 1))
    Set wbDaibiki = ActiveWorkbook
    Set wsDaibiki = wbDaibiki.Sheets(1)
    lastRowDaibiki = wsDaibiki.Cells(wsDaibiki.Rows.Count, 1).End(xlUp).row
    
    For i = 2 To lastRowDaibiki
        targetName = Replace(Replace(Trim(CStr(wsDaibiki.Cells(i, 1).Value)), " ", ""), "　", "")
        trackingNum = Trim(CStr(wsDaibiki.Cells(i, 11).Value))
        If targetName <> "" And trackingNum <> "" Then
            For j = 2 To lastRowOrder
                If Trim(CStr(wsOrder.Cells(j, 10).Value)) <> "" Then GoTo NextOrder1
                orderName = Replace(Replace(Trim(CStr(wsOrder.Cells(j, 2).Value)) & _
                            Trim(CStr(wsOrder.Cells(j, 3).Value)), " ", ""), "　", "")
                If orderName = targetName Then
                    wsOrder.Cells(j, 10).NumberFormat = "@"
                    wsOrder.Cells(j, 10).Value = trackingNum
                    Exit For
                End If
NextOrder1:
            Next j
        End If
    Next i
    wbDaibiki.Close False
    
    ' --- 2つ目: merged_tracking.csv ---
    If Len(daibikiPath2) > 0 Then
        If Dir(daibikiPath2) <> "" Then
            Workbooks.OpenText fileName:=daibikiPath2, _
                DataType:=xlDelimited, Comma:=True, _
                FieldInfo:=Array(Array(1, 2), Array(2, 1), Array(3, 1), Array(4, 1), _
                             Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
                             Array(9, 1), Array(10, 1), Array(11, 2), Array(12, 1), _
                             Array(13, 1), Array(14, 1))
            Set wbDaibiki = ActiveWorkbook
            Set wsDaibiki = wbDaibiki.Sheets(1)
            lastRowDaibiki = wsDaibiki.Cells(wsDaibiki.Rows.Count, 1).End(xlUp).row
            
            For i = 2 To lastRowDaibiki
                targetName = Replace(Replace(Trim(CStr(wsDaibiki.Cells(i, 1).Value)), " ", ""), "　", "")
                trackingNum = Trim(CStr(wsDaibiki.Cells(i, 11).Value))
                If targetName <> "" And trackingNum <> "" Then
                    For j = 2 To lastRowOrder
                        If Trim(CStr(wsOrder.Cells(j, 10).Value)) <> "" Then GoTo NextOrder2
                        orderName = Replace(Replace(Trim(CStr(wsOrder.Cells(j, 2).Value)) & _
                                    Trim(CStr(wsOrder.Cells(j, 3).Value)), " ", ""), "　", "")
                        If orderName = targetName Then
                            wsOrder.Cells(j, 10).NumberFormat = "@"
                            wsOrder.Cells(j, 10).Value = trackingNum
                            Exit For
                        End If
NextOrder2:
                    Next j
                End If
            Next i
            wbDaibiki.Close False
        End If
    End If
    
    ' --- 列削除と保存 ---
    wsOrder.Columns("B:I").Delete
    wbOrder.Save
    wbOrder.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub

