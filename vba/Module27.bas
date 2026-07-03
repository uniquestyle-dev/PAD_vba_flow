Attribute VB_Name = "Module27"
Sub TransferTrackingNumberCOD(ByVal orderPath As String, ByVal daibikiPath As String)
    
    On Error GoTo ErrHandler
    
    Dim wbOrder As Workbook, wbDaibiki As Workbook
    Dim wsOrder As Worksheet, wsDaibiki As Worksheet
    Dim lastRowDaibiki As Long, lastRowOrder As Long
    Dim i As Long, j As Long
    Dim targetID As String, trackingNum As String
    Dim matchCount As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbOrder = Workbooks.Open(orderPath)
    Set wbDaibiki = Workbooks.Open(daibikiPath)
    Set wsOrder = wbOrder.Sheets(1)
    Set wsDaibiki = wbDaibiki.Sheets(1)
    
    lastRowDaibiki = wsDaibiki.Cells(wsDaibiki.Rows.Count, 3).End(xlUp).row
    lastRowOrder = wsOrder.Cells(wsOrder.Rows.Count, 1).End(xlUp).row
    
    matchCount = 0
    For i = 2 To lastRowDaibiki
        targetID = Trim(CStr(wsDaibiki.Cells(i, 3).Value))
        trackingNum = Trim(CStr(wsDaibiki.Cells(i, 2).Value))
        
        If targetID <> "" And trackingNum <> "" Then
            For j = 2 To lastRowOrder
                If Trim(CStr(wsOrder.Cells(j, 1).Value)) = targetID Then
                    wsOrder.Cells(j, 10).Value = trackingNum
                    matchCount = matchCount + 1
                    Exit For
                End If
            Next j
        End If
    Next i
    
    ' B?IóÒçÌèú Å® AóÒ(íçï∂ID)Ç∆BóÒ(îzëóì`ï[î‘çÜ)ÇÃ2óÒÇ…Ç»ÇÈ
    wsOrder.Columns("B:I").Delete
    
    wbOrder.Save
    wbOrder.Close False
    wbDaibiki.Close False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "ÉGÉâÅ[: " & Err.Description & vbCrLf & "èÍèä: " & Erl, vbCritical
    
End Sub

