Attribute VB_Name = "Module36"
Sub TransferTrackingNumberSalacia転記のみ(ByVal orderPath As String, ByVal daibikiPath As String)
    On Error GoTo ErrHandler
    Dim wbOrder As Workbook, wbDaibiki As Workbook
    Dim wsOrder As Worksheet, wsDaibiki As Worksheet
    Dim lastRowDaibiki As Long, lastRowOrder As Long
    Dim i As Long, j As Long
    Dim targetName As String, orderName As String, trackingNum As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wbOrder = Workbooks.Open(orderPath)
    
    Workbooks.OpenText fileName:=daibikiPath, _
        DataType:=xlDelimited, _
        Comma:=True, _
        FieldInfo:=Array(Array(1, 2), Array(2, 1), Array(3, 1), Array(4, 1), _
                         Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
                         Array(9, 1), Array(10, 1), Array(11, 2), Array(12, 1), _
                         Array(13, 1), Array(14, 1))
    Set wbDaibiki = ActiveWorkbook
    
    Set wsOrder = wbOrder.Sheets(1)
    Set wsDaibiki = wbDaibiki.Sheets(1)
    lastRowDaibiki = wsDaibiki.Cells(wsDaibiki.Rows.Count, 1).End(xlUp).row
    lastRowOrder = wsOrder.Cells(wsOrder.Rows.Count, 1).End(xlUp).row
    
    For i = 2 To lastRowDaibiki
        targetName = Replace(Replace(Trim(CStr(wsDaibiki.Cells(i, 1).Value)), " ", ""), "　", "")
        trackingNum = Trim(CStr(wsDaibiki.Cells(i, 11).Value))
        If targetName <> "" And trackingNum <> "" Then
            For j = 2 To lastRowOrder
                If Trim(CStr(wsOrder.Cells(j, 10).Value)) <> "" Then GoTo NextOrder
                
                orderName = Replace(Replace(Trim(CStr(wsOrder.Cells(j, 2).Value)) & _
                            Trim(CStr(wsOrder.Cells(j, 3).Value)), " ", ""), "　", "")
                If orderName = targetName Then
                    wsOrder.Cells(j, 10).NumberFormat = "@"
                    wsOrder.Cells(j, 10).Value = trackingNum
                    Exit For
                End If
NextOrder:
            Next j
        End If
    Next i
    
    wbOrder.Save
    wbOrder.Close False
    wbDaibiki.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub

