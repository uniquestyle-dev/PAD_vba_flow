Attribute VB_Name = "Module35"
Sub TransferTrackingNumberStores(ByVal orderPath As String, ByVal daibikiPath As String)
    On Error GoTo ErrHandler
    Dim wbOrder As Workbook, wbDaibiki As Workbook
    Dim wsOrder As Worksheet, wsDaibiki As Worksheet
    Dim lastRowDaibiki As Long, lastRowOrder As Long
    Dim i As Long, j As Long
    Dim targetName As String, orderName As String, trackingNum As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wbOrder = Workbooks.Open(orderPath)
    Set wbDaibiki = Workbooks.Open(daibikiPath)
    Set wsOrder = wbOrder.Sheets(1)
    Set wsDaibiki = wbDaibiki.Sheets(1)
    lastRowDaibiki = wsDaibiki.Cells(wsDaibiki.Rows.Count, 1).End(xlUp).row
    lastRowOrder = wsOrder.Cells(wsOrder.Rows.Count, 1).End(xlUp).row
    
    For i = 2 To lastRowDaibiki
        targetName = Replace(Replace(Trim(CStr(wsDaibiki.Cells(i, 1).Value)), " ", ""), "　", "")
        trackingNum = Trim(CStr(wsDaibiki.Cells(i, 11).Value))
        If targetName <> "" And trackingNum <> "" Then
            For j = 2 To lastRowOrder
                orderName = Replace(Replace(Trim(CStr(wsOrder.Cells(j, 3).Value)) & _
                            Trim(CStr(wsOrder.Cells(j, 4).Value)), " ", ""), "　", "")
                If orderName = targetName And wsOrder.Cells(j, 7).Value = "" Then
                    wsOrder.Cells(j, 7).NumberFormat = "@"
                    wsOrder.Cells(j, 7).Value = trackingNum
                    Exit For
                End If
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

