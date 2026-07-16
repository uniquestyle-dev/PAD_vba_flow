Attribute VB_Name = "Module35"
Sub TransferTrackingNumberStores(ByVal orderPath As String, ByVal daibikiPath As String)
    On Error GoTo ErrHandler
    Dim wbOrder As Workbook, wbDaibiki As Workbook
    Dim wsOrder As Worksheet, wsDaibiki As Worksheet
    Dim lastRowDaibiki As Long, lastRowOrder As Long
    Dim i As Long, j As Long
    Dim targetName As String, orderName As String, trackingNum As String
    Dim fi As Variant
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    fi = Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), _
               Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
               Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), _
               Array(13, 2), Array(14, 2))
    
    Workbooks.OpenText fileName:=orderPath, _
        DataType:=xlDelimited, Comma:=True, FieldInfo:=fi
    Set wbOrder = ActiveWorkbook
    Set wsOrder = wbOrder.Sheets(1)
    
    Workbooks.OpenText fileName:=daibikiPath, _
        DataType:=xlDelimited, Comma:=True, FieldInfo:=fi
    Set wbDaibiki = ActiveWorkbook
    Set wsDaibiki = wbDaibiki.Sheets(1)
    lastRowDaibiki = wsDaibiki.Cells(wsDaibiki.Rows.Count, 1).End(xlUp).Row
    lastRowOrder = wsOrder.Cells(wsOrder.Rows.Count, 1).End(xlUp).Row
    
    Const COL_SHIPPED As Long = 10
    wsOrder.Cells(1, COL_SHIPPED).Value = ChrW(&H767A) & ChrW(&H9001) & ChrW(&H5B8C) & ChrW(&H4E86)

    For i = 2 To lastRowDaibiki
        targetName = Replace(Replace(Trim(CStr(wsDaibiki.Cells(i, 1).Value)), " ", ""), ChrW(&H3000), "")
        trackingNum = Trim(CStr(wsDaibiki.Cells(i, 11).Value))
        If targetName <> "" And trackingNum <> "" Then
            For j = 2 To lastRowOrder
                orderName = Replace(Replace(Trim(CStr(wsOrder.Cells(j, 3).Value)) & _
                            Trim(CStr(wsOrder.Cells(j, 4).Value)), " ", ""), ChrW(&H3000), "")
                If orderName = targetName And wsOrder.Cells(j, 7).Value = "" Then
                    wsOrder.Cells(j, 5).Value = ChrW(&H65E5) & ChrW(&H672C) & ChrW(&H90F5) & ChrW(&H4FBF)
                    wsOrder.Cells(j, 7).NumberFormat = "@"
                    wsOrder.Cells(j, 7).Value = trackingNum
                    Exit For
                End If
            Next j
        End If
    Next i
    
    For j = 2 To lastRowOrder
        If wsOrder.Cells(j, 7).Value <> "" Then
            wsOrder.Cells(j, COL_SHIPPED).Value = 1
        End If
    Next j
    
    wbOrder.SaveAs fileName:=orderPath, FileFormat:=xlCSV, Local:=True
    wbOrder.Close False
    wbDaibiki.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC) & ": " & Err.Description, vbCritical
End Sub
