Attribute VB_Name = "Module12"
Option Explicit

Sub Shopify変換()
    Dim filePath      As String
    Dim wb            As Workbook
    Dim ws            As Worksheet
    Dim qt            As QueryTable
    Dim lastRow       As Long, lastCol As Long
    Dim colIndex      As Long
    Dim idxZIP        As Long, idxShippingName As Long
    Dim idxProv       As Long, idxCity As Long, idxAddr1 As Long
    Dim idxOrderNum   As Long, idxBillingName As Long
    Dim dict          As Object, key As Variant
    Dim outputWb      As Workbook, outWs  As Worksheet
    Dim outputWb2     As Workbook, outWs2 As Worksheet
    Dim headers       As Variant, out2Headers As Variant
    Dim desktopPath   As String, saveFile As String
    Dim i             As Long, fileCount As Long, outRow As Long

    '─── 1. CSV を「テキストとして」インポート ───
    filePath = Application.GetOpenFilename("CSV ファイル (*.csv),*.csv", , _
                                           "入力 CSV を選択してください")
    If filePath = "False" Then Exit Sub

    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Sheets(1)

    Set qt = ws.QueryTables.Add("TEXT;" & filePath, ws.Range("A1"))
    With qt
        .TextFilePlatform = 65001                ' UTF-8
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True

        '―― 全 256 列を xlTextFormat (=2) に固定 ――
        Dim tArr(0 To 255) As Variant
        For i = 0 To 255: tArr(i) = xlTextFormat: Next i
        .TextFileColumnDataTypes = tArr
        '――――――――――――――――――――――――――――――――――

        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
    End With

    '─── 2. 必要列インデックス取得＆ソート ───
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For i = 1 To lastCol
        Select Case ws.Cells(1, i).Value
            Case "Lineitem name":           colIndex = i
            Case "Shipping Zip":            idxZIP = i
            Case "Shipping Name":           idxShippingName = i
            Case "Shipping Province Name":  idxProv = i
            Case "Shipping City":           idxCity = i
            Case "Shipping Address1":       idxAddr1 = i
            Case "Name":                    idxOrderNum = i
            Case "Billing Name":            idxBillingName = i
        End Select
    Next i

    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Columns(colIndex), Order:=xlAscending
    With ws.Sort
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .Header = xlYes
        .Apply
    End With

    '─── 3. Lineitem name ごとにキー収集 ───
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        key = ws.Cells(i, colIndex).Value
        If Not dict.Exists(key) Then dict.Add key, Nothing
    Next i

    '─── 4-5. タブ区切りテキスト（Shift_JIS）で連番出力 ───
    headers = Array( _
        "お届け先郵便番号", "お届け先氏名", "お届け先敬称", _
        "お届け先住所1行目", "お届け先住所2行目", "お届け先住所3行目", _
        "お届け先住所4行目", "内容品")

    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    fileCount = 1001

    For Each key In dict.Keys
        ws.AutoFilterMode = False
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter _
            field:=colIndex, Criteria1:=key

        Set outputWb = Workbooks.Add(xlWBATWorksheet)
        Set outWs = outputWb.Sheets(1)

        outWs.Range("A1:H1").Value = headers
        outWs.Columns("A:H").NumberFormat = "@"    ' 文字列書式に固定

        outRow = 2
        For i = 2 To lastRow
            If Not ws.Rows(i).Hidden Then
                outWs.Cells(outRow, 1).Value = Replace(ws.Cells(i, idxZIP).Value, "-", "")
                outWs.Cells(outRow, 2).Value = ws.Cells(i, idxShippingName).Value
                outWs.Cells(outRow, 3).Value = "様"
                outWs.Cells(outRow, 4).Value = ws.Cells(i, idxProv).Value
                outWs.Cells(outRow, 5).Value = ws.Cells(i, idxCity).Value
                outWs.Cells(outRow, 6).Value = SanitizeHyphen(ws.Cells(i, idxAddr1).Value)
                outWs.Cells(outRow, 7).Value = ""
                outWs.Cells(outRow, 8).Value = "書籍"
                outRow = outRow + 1
            End If
        Next i

        saveFile = desktopPath & fileCount & ".csv"
        outputWb.SaveAs fileName:=saveFile, FileFormat:=xlText, _
                        CreateBackup:=False, Local:=True
        outputWb.Close False
        fileCount = fileCount + 1
    Next key

    '─── 6. 999.csv を出力 ───
    out2Headers = Array( _
        "Order Number", "Tracking Number", "SKU", "Quantity", _
        "Tracking Company", "Tracking URL", "Billing Name")

    Set outputWb2 = Workbooks.Add(xlWBATWorksheet)
    Set outWs2 = outputWb2.Sheets(1)

    outWs2.Range("A1:G1").Value = out2Headers
    outWs2.Columns("A:G").NumberFormat = "@"

    outRow = 2
    For i = 2 To lastRow
        outWs2.Cells(outRow, 1).Value = ws.Cells(i, idxOrderNum).Value
        outWs2.Cells(outRow, 2).Value = ""
        outWs2.Cells(outRow, 3).Value = ""
        outWs2.Cells(outRow, 4).Value = ""
        outWs2.Cells(outRow, 5).Value = "Japan Post (JA)"
        outWs2.Cells(outRow, 6).Value = ""
        outWs2.Cells(outRow, 7).Value = ws.Cells(i, idxBillingName).Value
        outRow = outRow + 1
    Next i

    saveFile = desktopPath & "999.csv"
    outputWb2.SaveAs fileName:=saveFile, FileFormat:=xlCSV, _
                     CreateBackup:=False, Local:=True
    outputWb2.Close False

    '─── 後処理 ───
    ws.AutoFilterMode = False
    wb.Close False

    '─── 999.csv を開く ───
    Workbooks.Open fileName:=saveFile
End Sub


'----------------------------------------------
' 文字化けしやすい全角/特殊ハイフン類 → ASCII "-" に統一
'----------------------------------------------
Private Function SanitizeHyphen(ByVal txt As String) As String
    Const HN As String = "-"
    txt = Replace(txt, ChrW(&H2212), HN) ' ? (数学マイナス)
    txt = Replace(txt, ChrW(&HFF0D), HN) ' － (全角ハイフン)
    txt = Replace(txt, ChrW(&H2010), HN) ' ‐
    txt = Replace(txt, ChrW(&H2013), HN) ' ?
    txt = Replace(txt, ChrW(&H2014), HN) ' ?
    txt = Replace(txt, ChrW(&H30FC), HN) ' ー (長音符)
    SanitizeHyphen = txt
End Function


