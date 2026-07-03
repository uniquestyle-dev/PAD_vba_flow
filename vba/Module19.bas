Attribute VB_Name = "Module19"
Option Explicit

Sub Shopify変換()
    Dim filePath      As String
    Dim wb            As Workbook
    Dim ws            As Worksheet
    Dim qt            As QueryTable
    Dim lastRow       As Long, lastCol As Long
    Dim colIndex      As Long, idxSku As Long
    Dim idxZIP        As Long, idxShippingName As Long
    Dim idxProv       As Long, idxCity As Long, idxAddr1 As Long
    Dim idxOrderNum   As Long, idxBillingName As Long
    Dim idxPaymentMethod As Long
    Dim dict          As Object, key As Variant
    Dim outputWb      As Workbook, outWs  As Worksheet
    Dim outputWb2     As Workbook, outWs2 As Worksheet
    Dim headers       As Variant, out2Headers As Variant
    Dim desktopPath   As String, saveFile As String
    Dim i             As Long, j As Long, outRow As Long
    Dim skuValue      As String, category As String
    Dim fso As Object

    '─── 1. 本日のorders_export.csvを自動取得 ───
    Dim folder As Object, file As Object
    Dim downloadPath As String
    
    downloadPath = "C:\Users\lenovo\Desktop\ダウンロード\"
    filePath = ""
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(downloadPath)
    
    Dim latestDate As Date
    latestDate = 0
    
    For Each file In folder.Files
        If LCase(file.Name) Like "orders_export*.csv" Then
            If file.DateCreated > latestDate Then
                latestDate = file.DateCreated
                filePath = file.Path
            End If
        End If
    Next file
    
    If filePath = "" Then
        Debug.Print "本日のorders_export.csvが見つかりません"
        Exit Sub
    End If
    
    Debug.Print "ファイル: " & filePath

    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Sheets(1)

    Set qt = ws.QueryTables.Add("TEXT;" & filePath, ws.Range("A1"))
    With qt
        .TextFilePlatform = 65001
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True

        Dim tArr(0 To 255) As Variant
        For i = 0 To 255: tArr(i) = xlTextFormat: Next i
        .TextFileColumnDataTypes = tArr

        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
    End With

    '─── 2. 必要列インデックス取得 ───
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For i = 1 To lastCol
        Select Case ws.Cells(1, i).Value
            Case "Lineitem name":           colIndex = i
            Case "Lineitem sku":            idxSku = i
            Case "Shipping Zip":            idxZIP = i
            Case "Shipping Name":           idxShippingName = i
            Case "Shipping Province Name":  idxProv = i
            Case "Shipping City":           idxCity = i
            Case "Shipping Address1":       idxAddr1 = i
            Case "Name":                    idxOrderNum = i
            Case "Billing Name":            idxBillingName = i
            Case "Payment Method":          idxPaymentMethod = i
        End Select
    Next i

    '─── 出力先フォルダ ───
    desktopPath = "\Users\lenovo\Desktop\ダウンロード\"

    '─── 3. 代引き.csv を出力(Payment Method = Cash on Delivery (COD))───
    Dim codWb As Workbook, codWs As Worksheet
    Dim codRow As Long
    Dim adoStream As Object
    Dim csvLine As String
    Dim codLastRow As Long, codLastCol As Long
    Dim cellVal As String
    
    Set codWb = Workbooks.Add(xlWBATWorksheet)
    Set codWs = codWb.Sheets(1)
    
    ' ヘッダ行をコピー
    ws.Rows(1).Copy codWs.Rows(1)
    
    ' COD行をコピー
    codRow = 2
    For i = 2 To lastRow
        If Trim(ws.Cells(i, idxPaymentMethod).Value) = "Cash on Delivery (COD)" Then
            ws.Rows(i).Copy codWs.Rows(codRow)
            codRow = codRow + 1
        End If
    Next i
    
    If codRow > 2 Then
        codLastRow = codRow - 1
        codLastCol = codWs.Cells(1, codWs.Columns.Count).End(xlToLeft).Column
        
        ' L列(Total)に290を加算
        For i = 2 To codLastRow
            If IsNumeric(codWs.Cells(i, "L").Value) Then
                codWs.Cells(i, "L").Value = codWs.Cells(i, "L").Value + 290
            End If
        Next i
        
        ' ADODB.Stream で Shift_JIS 書き出し
        saveFile = desktopPath & "代引き.csv"
        Set adoStream = CreateObject("ADODB.Stream")
        With adoStream
            .Type = 2                 ' テキストモード
            .Charset = "Shift_JIS"
            .Open
            For i = 1 To codLastRow
                If Application.WorksheetFunction.CountA( _
                       codWs.Range(codWs.Cells(i, 1), codWs.Cells(i, codLastCol)) _
                   ) > 0 Then
                    csvLine = ""
                    For j = 1 To codLastCol
                        cellVal = codWs.Cells(i, j).Text
                        
                        ' AR列(Shipping Phone)の処理
                        If j = 44 And i > 1 Then  ' AR列は44列目
                            ' +81を0に置換
                            cellVal = Replace(cellVal, "+81", "0")
                            ' ハイフン・スペースを除去
                            cellVal = Replace(cellVal, "-", "")
                            cellVal = Replace(cellVal, " ", "")
                            ' 数値のみで9桁(固定電話)か10桁(携帯)なら先頭に0追加
                            If (Len(cellVal) = 9 Or Len(cellVal) = 10) And IsNumeric(cellVal) Then
                                cellVal = "0" & cellVal
                            End If
                        End If
                        
                        cellVal = Replace(cellVal, """", """""")
                        csvLine = csvLine & """" & cellVal & """"
                        If j < codLastCol Then csvLine = csvLine & ","
                    Next j
                    .WriteText csvLine, 1  ' adWriteLine
                End If
            Next i
            .SaveToFile saveFile, 2  ' adSaveCreateOverWrite
            .Close
        End With
        
        Debug.Print "出力: " & saveFile & " (" & (codLastRow - 1) & "件)"
    Else
        Debug.Print "スキップ: 代引き (データなし)"
    End If
    
    codWb.Close SaveChanges:=False

    '─── 4. SKUカテゴリごとにキー収集 ───
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        skuValue = LCase(Trim(ws.Cells(i, idxSku).Value))
        category = GetCategory(skuValue)
        If category <> "" Then
            If Not dict.Exists(category) Then dict.Add category, Nothing
        End If
    Next i

    '─── 発送ラベル用ヘッダ ───
    headers = Array( _
        "お届け先郵便番号", "お届け先氏名", "お届け先敬称", _
        "お届け先住所1行目", "お届け先住所2行目", "お届け先住所3行目", _
        "お届け先住所4行目", "内容品")

    '─── 5. カテゴリごとにタブ区切りテキスト出力 ───
    For Each key In dict.Keys
        Set outputWb = Workbooks.Add(xlWBATWorksheet)
        Set outWs = outputWb.Sheets(1)

        outWs.Range("A1:H1").Value = headers
        outWs.Columns("A:H").NumberFormat = "@"

        outRow = 2
        For i = 2 To lastRow
            skuValue = LCase(Trim(ws.Cells(i, idxSku).Value))
            category = GetCategory(skuValue)
            
            If category = key Then
                outWs.Cells(outRow, 1).Value = Replace(ws.Cells(i, idxZIP).Value, "-", "")
                outWs.Cells(outRow, 2).Value = SanitizeClickPost(ws.Cells(i, idxShippingName).Value)
                outWs.Cells(outRow, 3).Value = "様"
                outWs.Cells(outRow, 4).Value = SanitizeClickPost(ws.Cells(i, idxProv).Value)
                outWs.Cells(outRow, 5).Value = SanitizeClickPost(ws.Cells(i, idxCity).Value)
                outWs.Cells(outRow, 6).Value = SanitizeClickPost(SanitizeHyphen(ws.Cells(i, idxAddr1).Value))
                outWs.Cells(outRow, 7).Value = ""
                outWs.Cells(outRow, 8).Value = "書籍"
                outRow = outRow + 1
            End If
        Next i

        ' データがあれば保存
        If outRow > 2 Then
            saveFile = desktopPath & key & ".csv"
            outputWb.SaveAs fileName:=saveFile, FileFormat:=xlText, _
                            CreateBackup:=False, Local:=True
            outputWb.Close False
            Debug.Print "出力: " & saveFile & " (" & (outRow - 2) & "件)"
        Else
            outputWb.Close False
            Debug.Print "スキップ: " & key & " (データなし)"
        End If
    Next key

    '─── 6. 999.csv を出力 ───
    out2Headers = Array( _
        "Order Number", "Tracking Number", _
        "Tracking Company", "Tracking URL", "Billing Name")

    Set outputWb2 = Workbooks.Add(xlWBATWorksheet)
    Set outWs2 = outputWb2.Sheets(1)

    outWs2.Range("A1:E1").Value = out2Headers
    outWs2.Columns("A:E").NumberFormat = "@"

        Dim catKeys() As String
    Dim catCount As Long
    catCount = dict.Count
    ReDim catKeys(0 To catCount - 1)
    Dim k As Long
    k = 0
    For Each key In dict.Keys
        catKeys(k) = CStr(key)
        k = k + 1
    Next key

    ' バブルソート（大文字小文字を区別しない昇順）
    Dim tmp As String
    For i = 0 To catCount - 2
        For j = 0 To catCount - 2 - i
            If LCase(catKeys(j)) > LCase(catKeys(j + 1)) Then
                tmp = catKeys(j)
                catKeys(j) = catKeys(j + 1)
                catKeys(j + 1) = tmp
            End If
        Next j
    Next i

    ' カテゴリ順にデータ行を出力
    outRow = 2
    For k = 0 To catCount - 1
        For i = 2 To lastRow
            skuValue = LCase(Trim(ws.Cells(i, idxSku).Value))
            category = GetCategory(skuValue)
            If category = catKeys(k) Then
                outWs2.Cells(outRow, 1).Value = ws.Cells(i, idxOrderNum).Value
                outWs2.Cells(outRow, 2).Value = ""
                outWs2.Cells(outRow, 3).Value = "Japan Post (JA)"
                outWs2.Cells(outRow, 4).Value = ""
                outWs2.Cells(outRow, 5).Value = ws.Cells(i, idxBillingName).Value
                outRow = outRow + 1
            End If
        Next i
    Next k

    saveFile = desktopPath & "999.csv"
    outputWb2.SaveAs fileName:=saveFile, FileFormat:=xlCSV, _
                     CreateBackup:=False, Local:=True
    outputWb2.Close False

    '─── 後処理 ───
    ws.AutoFilterMode = False
    wb.Close False

    
    Debug.Print "===== 完了 ====="
End Sub


'----------------------------------------------
' SKU → カテゴリ名マッピング
'----------------------------------------------
Private Function GetCategory(ByVal sku As String) As String
    Select Case sku
        ' carboff
        Case "07b", "01b", "cb03b", "cbrgl", "f07b", "f01b", "g07b", "g01b", "y07b", "y01b", "cb-print", "cb03b2", "p07b", "p01b"
            GetCategory = "carboff"
        ' abura
        Case "abura"
            GetCategory = "abura"
        ' salasiru
        Case "ss03-2b", "ss03-3b", "ss03b", "ssrgl", "ss-print"
            GetCategory = "salasiru"
        ' Medico
        Case "mc03b", "mc03b-2", "mcrgl"
            GetCategory = "Medico"
        Case Else
            GetCategory = "other_" & sku
    End Select
End Function


'----------------------------------------------
' 文字化けしやすい全角/特殊ハイフン類 → ASCII "-" に統一
'----------------------------------------------
Private Function SanitizeHyphen(ByVal txt As String) As String
    Const HN As String = "-"
    txt = Replace(txt, ChrW(&H2212), HN)
    txt = Replace(txt, ChrW(&HFF0D), HN)
    txt = Replace(txt, ChrW(&H2010), HN)
    txt = Replace(txt, ChrW(&H2013), HN)
    txt = Replace(txt, ChrW(&H2014), HN)
    txt = Replace(txt, ChrW(&H30FC), HN)
    SanitizeHyphen = txt
End Function


'----------------------------------------------
' ClickPost取込エラー防止: , と " を除去
'----------------------------------------------
Private Function SanitizeClickPost(ByVal txt As String) As String
    SanitizeClickPost = Replace(Replace(txt, ",", ""), """", "")
End Function


