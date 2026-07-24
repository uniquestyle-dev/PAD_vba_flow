Attribute VB_Name = "Module35"
Sub TransferTrackingNumberStores(ByVal orderPath As String, ByVal daibikiPath As String)
    On Error GoTo ErrHandler
    Dim fNum As Integer
    Dim line As String
    Dim orderLines() As String
    Dim orderCount As Long
    Dim daibikiNames() As String
    Dim daibikiTracking() As String
    Dim daibikiCount As Long
    Dim i As Long, j As Long
    Dim parts As Variant, dParts As Variant
    Dim targetName As String, orderName As String, trackingNum As String
    Dim matchCount As Long
    
    orderCount = 0
    daibikiCount = 0
    
    fNum = FreeFile
    Open orderPath For Input As #fNum
    Do Until EOF(fNum)
        Line Input #fNum, line
        If Len(line) > 0 Then
            orderCount = orderCount + 1
            ReDim Preserve orderLines(1 To orderCount)
            orderLines(orderCount) = line
        End If
    Loop
    Close #fNum
    
    fNum = FreeFile
    Open daibikiPath For Input As #fNum
    Do Until EOF(fNum)
        Line Input #fNum, line
        If Len(line) > 0 Then
            daibikiCount = daibikiCount + 1
            ReDim Preserve daibikiNames(1 To daibikiCount)
            ReDim Preserve daibikiTracking(1 To daibikiCount)
            dParts = SplitCSVLine(line)
            If UBound(dParts) >= 10 Then
                daibikiNames(daibikiCount) = StripQuotes(dParts(0))
                daibikiTracking(daibikiCount) = StripQuotes(dParts(10))
            Else
                daibikiNames(daibikiCount) = ""
                daibikiTracking(daibikiCount) = ""
            End If
        End If
    Loop
    Close #fNum
    
    Dim outLines() As String
    ReDim outLines(1 To orderCount)
    outLines(1) = orderLines(1) & "," & ChrW(&H767A) & ChrW(&H9001) & ChrW(&H5B8C) & ChrW(&H4E86)
    
    matchCount = 0
    For i = 2 To orderCount
        parts = SplitCSVLine(orderLines(i))
        Dim col3 As String, col4 As String, col7 As String
        col3 = "": col4 = "": col7 = ""
        If UBound(parts) >= 2 Then col3 = StripQuotes(parts(2))
        If UBound(parts) >= 3 Then col4 = StripQuotes(parts(3))
        If UBound(parts) >= 6 Then col7 = StripQuotes(parts(6))
        
        orderName = Replace(Replace(Trim(col3) & Trim(col4), " ", ""), ChrW(&H3000), "")
        
        Dim matched As Boolean
        matched = False
        If col7 = "" Then
            For j = 2 To daibikiCount
                targetName = Replace(Replace(Trim(daibikiNames(j)), " ", ""), ChrW(&H3000), "")
                trackingNum = Trim(daibikiTracking(j))
                If targetName <> "" And trackingNum <> "" Then
                    If orderName = targetName Then
                        If UBound(parts) >= 6 Then parts(6) = trackingNum
                        If UBound(parts) >= 4 Then parts(4) = ChrW(&H65E5) & ChrW(&H672C) & ChrW(&H90F5) & ChrW(&H4FBF)
                        matched = True
                        matchCount = matchCount + 1
                        Exit For
                    End If
                End If
            Next j
        End If
        
        Dim rebuilt As String
        rebuilt = Join(parts, ",")
        If matched Or col7 <> "" Then
            rebuilt = rebuilt & ",1"
        Else
            rebuilt = rebuilt & ","
        End If
        outLines(i) = rebuilt
    Next i
    
    fNum = FreeFile
    Open orderPath For Output As #fNum
    For i = 1 To orderCount
        Print #fNum, outLines(i)
    Next i
    Close #fNum
    
    Exit Sub
ErrHandler:
    Dim logFile As String
    Dim logNum As Integer
    logFile = "C:\Users\lenovo\Desktop\" & ChrW(&H30C0) & ChrW(&H30A6) & ChrW(&H30F3) & ChrW(&H30ED) & ChrW(&H30FC) & ChrW(&H30C9) & "\Module35_log.txt"
    logNum = FreeFile
    Open logFile For Output As #logNum
    Print #logNum, "ERROR: " & Err.Description & " (Err.Number=" & Err.Number & ")"
    Close #logNum
End Sub

Private Function SplitCSVLine(ByVal line As String) As Variant
    Dim result() As String
    Dim count As Long
    Dim i As Long
    Dim inQuote As Boolean
    Dim cur As String
    Dim ch As String
    
    count = 0
    inQuote = False
    cur = ""
    
    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        If ch = """" Then
            inQuote = Not inQuote
            cur = cur & ch
        ElseIf ch = "," And Not inQuote Then
            count = count + 1
            ReDim Preserve result(0 To count - 1)
            result(count - 1) = cur
            cur = ""
        Else
            cur = cur & ch
        End If
    Next i
    count = count + 1
    ReDim Preserve result(0 To count - 1)
    result(count - 1) = cur
    
    SplitCSVLine = result
End Function

Private Function StripQuotes(ByVal s As String) As String
    s = Trim(s)
    If Len(s) >= 2 Then
        If Left(s, 1) = """" And Right(s, 1) = """" Then
            s = Mid(s, 2, Len(s) - 2)
        End If
    End If
    StripQuotes = s
End Function