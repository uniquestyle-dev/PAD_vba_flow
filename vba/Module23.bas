Attribute VB_Name = "Module23"
Option Explicit

Public Sub 代引きお問い合わせ番号追記(ByVal ordersFilePath As String)
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folderPath As String
    folderPath = fso.GetParentFolderName(ordersFilePath)
    
    Dim trackingFilePath As String
    trackingFilePath = folderPath & "\" & ChrW(&H4EE3) & ChrW(&H5F15) & ChrW(&H304D) & ChrW(&H3068) & ChrW(&H304A) & ChrW(&H554F) & ChrW(&H3044) & ChrW(&H5408) & ChrW(&H308F) & ChrW(&H305B) & ChrW(&H756A) & ChrW(&H53F7) & ".csv"
    
    Dim outputFilePath As String
    outputFilePath = folderPath & "\111.csv"
    
    If Not fso.FileExists(ordersFilePath) Then
        Err.Raise 10001, , "orders_export.csv が見つかりません: " & ordersFilePath
        Exit Sub
    End If
    If Not fso.FileExists(trackingFilePath) Then
        Err.Raise 10002, , "代引きとお問い合わせ番号.csv が見つかりません: " & trackingFilePath
        Exit Sub
    End If
    
    ' ========================================
    ' STEP 1: 代引きファイルから辞書構築
    ' ========================================
    Dim dictTracking As Object
    Set dictTracking = CreateObject("Scripting.Dictionary")
    
    Dim ts As Object
    Set ts = CreateObject("ADODB.Stream")
    ts.Type = 2
    ts.Charset = "Shift_JIS"
    ts.Open
    ts.LoadFromFile trackingFilePath
    Dim trackingContent As String
    trackingContent = ts.ReadText(-1)
    ts.Close
    Set ts = Nothing
    
    Dim trackingLines() As String
    trackingLines = Split(trackingContent, vbLf)
    
    Dim i As Long
    Dim tLine As String
    Dim tFields() As String
    Dim phoneKey As String
    Dim trackNum As String
    
    For i = 1 To UBound(trackingLines)
        tLine = Trim(Replace(trackingLines(i), vbCr, ""))
        If Len(tLine) > 0 Then
            tFields = ParseCsvLine(tLine)
            If UBound(tFields) >= 1 Then
                phoneKey = NormalizePhone(StripQuotes(tFields(0)))
                trackNum = StripQuotes(tFields(1))
                If Len(phoneKey) > 0 And Len(trackNum) > 0 Then
                    If Not dictTracking.Exists(phoneKey) Then
                        dictTracking.Add phoneKey, trackNum
                    End If
                End If
            End If
        End If
    Next i
    
    ' ========================================
    ' STEP 2: orders_export.csv 読み込み・突合
    ' ========================================
    Dim tsO As Object
    Set tsO = CreateObject("ADODB.Stream")
    tsO.Type = 2
    tsO.Charset = "UTF-8"
    tsO.Open
    tsO.LoadFromFile ordersFilePath
    Dim ordersContent As String
    ordersContent = tsO.ReadText(-1)
    tsO.Close
    Set tsO = Nothing
    
    ' BOM除去
    If Len(ordersContent) > 0 Then
        If AscW(Left(ordersContent, 1)) = 65279 Then
            ordersContent = Mid(ordersContent, 2)
        End If
    End If
    
    Dim ordersLines() As String
    ordersLines = Split(ordersContent, vbLf)
    
    ' ヘッダー解析
    Dim headerFields() As String
    headerFields = ParseCsvLine(Trim(Replace(ordersLines(0), vbCr, "")))
    
    Dim colName As Long:         colName = FindCol(headerFields, "Name")
    Dim colShipPhone As Long:    colShipPhone = FindCol(headerFields, "Shipping Phone")
    Dim colBillPhone As Long:    colBillPhone = FindCol(headerFields, "Billing Phone")
    Dim colPhone As Long:        colPhone = FindCol(headerFields, "Phone")
    Dim colBillingName As Long:  colBillingName = FindCol(headerFields, "Billing Name")
    
    If colName < 0 Or colBillingName < 0 Then
        Err.Raise 10003, , "orders_export.csv: Name / Billing Name 列が見つかりません"
        Exit Sub
    End If
    
    Dim dictProcessed As Object
    Set dictProcessed = CreateObject("Scripting.Dictionary")
    
    Dim outputLines As String
    outputLines = "Order Number,Tracking Number,Tracking Company,Tracking URL,Billing Name" & vbCrLf
    
    Dim matchCount As Long: matchCount = 0
    Dim totalOrders As Long: totalOrders = 0
    Dim oLine As String
    Dim oFields() As String
    Dim orderName As String
    Dim phone As String
    Dim matchedTracking As String
    Dim billingName As String
    
    For i = 1 To UBound(ordersLines)
        oLine = Trim(Replace(ordersLines(i), vbCr, ""))
        If Len(oLine) > 0 Then
            oFields = ParseCsvLine(oLine)
            orderName = StripQuotes(oFields(colName))
            
            If Not dictProcessed.Exists(orderName) Then
                dictProcessed.Add orderName, True
                totalOrders = totalOrders + 1
                
                ' 電話番号取得（Shipping Phone → Billing Phone → Phone）
                phone = ""
                If colShipPhone >= 0 And colShipPhone <= UBound(oFields) Then
                    phone = NormalizePhone(StripQuotes(oFields(colShipPhone)))
                End If
                If Len(phone) = 0 And colBillPhone >= 0 And colBillPhone <= UBound(oFields) Then
                    phone = NormalizePhone(StripQuotes(oFields(colBillPhone)))
                End If
                If Len(phone) = 0 And colPhone >= 0 And colPhone <= UBound(oFields) Then
                    phone = NormalizePhone(StripQuotes(oFields(colPhone)))
                End If
                
                ' 突合
                matchedTracking = ""
                If Len(phone) > 0 And dictTracking.Exists(phone) Then
                    matchedTracking = dictTracking.item(phone)
                    matchCount = matchCount + 1
                End If
                
                billingName = ""
                If colBillingName <= UBound(oFields) Then
                    billingName = StripQuotes(oFields(colBillingName))
                End If
                
                outputLines = outputLines & _
                    orderName & "," & _
                    matchedTracking & "," & _
                    "Japan Post (JA)," & _
                    "," & _
                    billingName & vbCrLf
            End If
        End If
    Next i
    
    ' ========================================
    ' STEP 3: 111.csv 出力（Shift-JIS）
    ' ========================================
    Dim tsOut As Object
    Set tsOut = CreateObject("ADODB.Stream")
    tsOut.Type = 2
    tsOut.Charset = "Shift_JIS"
    tsOut.Open
    tsOut.WriteText outputLines
    tsOut.SaveToFile outputFilePath, 2
    tsOut.Close
    Set tsOut = Nothing
    
    Set dictTracking = Nothing
    Set dictProcessed = Nothing
    Set fso = Nothing

End Sub


' === 電話番号正規化 ===
Private Function NormalizePhone(ByVal ph As String) As String
    Dim result As String, j As Long, code As Long, normalized As String
    result = Trim(ph)
    
    ' 全角数字→半角
    normalized = ""
    For j = 1 To Len(result)
        code = AscW(Mid(result, j, 1))
        If code >= &HFF10 And code <= &HFF19 Then
            normalized = normalized & Chr(code - &HFF10 + 48)
        Else
            normalized = normalized & Mid(result, j, 1)
        End If
    Next j
    result = normalized
    
    result = Replace(result, "-", "")
    result = Replace(result, ChrW(&HFF0D), "")
    result = Replace(result, ChrW(&H30FC), "")
    result = Replace(result, " ", "")
    result = Replace(result, ChrW(&H3000), "")
    result = Replace(result, "(", "")
    result = Replace(result, ")", "")
    result = Replace(result, ChrW(&HFF08), "")
    result = Replace(result, ChrW(&HFF09), "")
    
    If Left(result, 3) = "+81" Then result = "0" & Mid(result, 4)
    If Left(result, 2) = "81" And Len(result) >= 12 Then result = "0" & Mid(result, 3)
    
    NormalizePhone = result
End Function


' === ダブルクォート除去 ===
Private Function StripQuotes(ByVal s As String) As String
    Dim r As String: r = Trim(s)
    If Len(r) >= 2 Then
        If Left(r, 1) = """" And Right(r, 1) = """" Then r = Mid(r, 2, Len(r) - 2)
    End If
    StripQuotes = r
End Function


' === CSVパース（クォート内カンマ対応） ===
Private Function ParseCsvLine(ByVal line As String) As String()
    Dim fields() As String, fieldCount As Long, inQuote As Boolean
    Dim current As String, c As String, pos As Long
    
    fieldCount = 0: ReDim fields(0): inQuote = False: current = "": pos = 1
    
    Do While pos <= Len(line)
        c = Mid(line, pos, 1)
        If c = """" Then
            If inQuote Then
                If pos < Len(line) And Mid(line, pos + 1, 1) = """" Then
                    current = current & """": pos = pos + 1
                Else
                    inQuote = False
                End If
            Else
                inQuote = True
            End If
        ElseIf c = "," And Not inQuote Then
            ReDim Preserve fields(fieldCount): fields(fieldCount) = current
            fieldCount = fieldCount + 1: current = ""
        Else
            current = current & c
        End If
        pos = pos + 1
    Loop
    
    ReDim Preserve fields(fieldCount): fields(fieldCount) = current
    ParseCsvLine = fields
End Function


' === ヘッダー列検索 ===
Private Function FindCol(ByRef headers() As String, ByVal colName As String) As Long
    Dim k As Long
    For k = 0 To UBound(headers)
        If LCase(Trim(StripQuotes(headers(k)))) = LCase(colName) Then
            FindCol = k: Exit Function
        End If
    Next k
    FindCol = -1
End Function



