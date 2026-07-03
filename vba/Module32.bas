Attribute VB_Name = "Module32"
'--------------------------------------------------------------
' ExportCodNamePdfs
' orders_export.csv（Shopify代引き注文エクスポート）から
' SKUに基づいて名前を商品カテゴリ別のPDFファイルに出力する。
'
' 出力:
'   C:\Users\lenovo\Desktop\ダウンロード\carboff-CODname.pdf
'   C:\Users\lenovo\Desktop\ダウンロード\salasiru-CODname.pdf
'   C:\Users\lenovo\Desktop\ダウンロード\Medico-CODname.pdf
'   C:\Users\lenovo\Desktop\ダウンロード\abura-CODname.pdf
'   C:\Users\lenovo\Desktop\ダウンロード\carboff_red-CODname.pdf  ← 追加（red_cod_name.txt から生成）
'   ※ 該当なしのカテゴリはファイルを生成しない
'
' orders_export.csv が無い場合でも、carboff_red は red_cod_name.txt から生成する。
'--------------------------------------------------------------
Sub ExportCodNamePdfs()
    Const DL_DIR = "C:\Users\lenovo\Desktop\ダウンロード\"
    Const CSV_NAME = "orders_export.csv"

    ' 赤(carboff_red)は red_cod_name.txt から生成（orders_export の有無に依存しない）
    ExportRedCodNamePdf

    Dim csvPath As String
    csvPath = DL_DIR & CSV_NAME

    If Dir(csvPath) = "" Then Exit Sub

    ' --- UTF-8でCSV読み込み ---
    Dim adoStream As Object
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Open
    adoStream.Type = 2
    adoStream.Charset = "utf-8"
    adoStream.LoadFromFile csvPath
    Dim csvText As String
    csvText = adoStream.ReadText
    adoStream.Close
    Set adoStream = Nothing

    If Len(csvText) > 0 And AscW(Left(csvText, 1)) = &HFEFF Then
        csvText = Mid(csvText, 2)
    End If

    csvText = Replace(csvText, vbCrLf, vbLf)
    csvText = Replace(csvText, vbCr, vbLf)
    Dim lines() As String
    lines = Split(csvText, vbLf)

    If UBound(lines) < 1 Then Exit Sub

    ' --- ヘッダー解析 ---
    Dim headerFields() As String
    headerFields = ParseCsvLine_COD(lines(0))

    Dim colSku As Long, colShipName As Long, colOrder As Long
    colSku = -1: colShipName = -1: colOrder = -1

    Dim j As Long
    For j = 0 To UBound(headerFields)
        Select Case Trim(headerFields(j))
            Case "Lineitem sku":  colSku = j
            Case "Shipping Name": colShipName = j
            Case "Name":          colOrder = j
        End Select
    Next j

    If colSku = -1 Or colShipName = -1 Then Exit Sub

    ' --- カテゴリ別Dictionary ---
    Dim categories As Object
    Set categories = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To UBound(lines)
        If Len(Trim(lines(i))) = 0 Then GoTo NextRow

        Dim fields() As String
        fields = ParseCsvLine_COD(lines(i))

        If UBound(fields) < colSku Or UBound(fields) < colShipName Then GoTo NextRow

        Dim sku As String
        sku = Trim(fields(colSku))

        Dim category As String
        category = GetCodCategory_COD(sku)

        If Len(category) > 0 Then
            Dim nm As String
            nm = Trim(fields(colShipName))
            If Len(nm) > 0 Then
                Dim orderNum As String
                If colOrder >= 0 And UBound(fields) >= colOrder Then
                    orderNum = Trim(fields(colOrder))
                Else
                    orderNum = CStr(i)
                End If
                Dim dictKey As String
                dictKey = orderNum & "|" & nm

                If Not categories.Exists(category) Then
                    Set categories(category) = CreateObject("Scripting.Dictionary")
                End If
                If Not categories(category).Exists(dictKey) Then
                    categories(category).Add dictKey, nm
                End If
            End If
        End If
NextRow:
    Next i

    ' --- カテゴリ別PDF出力 ---
    Dim catKey As Variant
    For Each catKey In categories.Keys
        Dim pdfPath As String
        pdfPath = DL_DIR & CStr(catKey) & "-CODname.pdf"

        Dim wbTemp As Workbook
        Set wbTemp = Workbooks.Add(xlWBATWorksheet)
        Dim wsTemp As Worksheet
        Set wsTemp = wbTemp.Sheets(1)

        ' --- セル書き込み ---
        Dim displayLabel As String
        displayLabel = GetCodDisplayLabel_COD(CStr(catKey)) & "-代引き"

        wsTemp.Cells(1, 1).Value = displayLabel
        ' 2行目は空行

        Dim row As Long
        row = 3
        Dim nameKey As Variant
        For Each nameKey In categories(catKey).Keys
            wsTemp.Cells(row, 1).Value = categories(catKey)(nameKey)
            row = row + 1
        Next nameKey

        Dim lastRow As Long
        lastRow = row - 1
        wsTemp.Range("A1").Font.Size = 28
        wsTemp.Range("A3:A" & lastRow).Font.Size = 48
        wsTemp.Columns("A").AutoFit

        wsTemp.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=pdfPath, _
            Quality:=xlQualityStandard

        wbTemp.Close SaveChanges:=False
        Set wsTemp = Nothing
        Set wbTemp = Nothing
    Next catKey
End Sub


'--------------------------------------------------------------
' ExportRedCodNamePdf
' red_cod_name.txt（1行1名）から carboff_red-CODname.pdf を生成する。
' 体裁は他のCODname PDFと同一（1行目=タイトル / 2行目=空行 / 3行目以降=名前）。
' red_cod_name.txt が無い、または名前0件 の場合はPDFを生成しない。
'--------------------------------------------------------------
Private Sub ExportRedCodNamePdf()
    Const DL_DIR = "C:\Users\lenovo\Desktop\ダウンロード\"
    Const TXT_NAME = "red_cod_name.txt"

    Dim txtPath As String
    txtPath = DL_DIR & TXT_NAME
    If Dir(txtPath) = "" Then Exit Sub

    Dim wbTemp As Workbook, wsTemp As Worksheet
    Set wbTemp = Workbooks.Add(xlWBATWorksheet)
    Set wsTemp = wbTemp.Sheets(1)

    ' 1行目=タイトル（Python側 skip_first_line で除外）/ 2行目=空行
    wsTemp.Cells(1, 1).Value = "【赤】carboff-代引き"

    Dim row As Long
    row = 3

    Dim ff As Integer
    ff = FreeFile
    Dim lineStr As String
    Open txtPath For Input As #ff      ' Salaciaマクロが Print # で書いた前提（システム既定=Shift_JIS）
    Do While Not EOF(ff)
        Line Input #ff, lineStr
        lineStr = Trim(lineStr)
        If Len(lineStr) > 0 Then
            wsTemp.Cells(row, 1).Value = lineStr
            row = row + 1
        End If
    Loop
    Close #ff

    ' 名前0件ならPDFを作らず終了（他カテゴリの「該当なしは生成しない」と同じ）
    If row = 3 Then
        wbTemp.Close SaveChanges:=False
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = row - 1
    wsTemp.Range("A1").Font.Size = 28
    wsTemp.Range("A3:A" & lastRow).Font.Size = 48
    wsTemp.Columns("A").AutoFit

    wsTemp.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=DL_DIR & "carboff_red-CODname.pdf", _
        Quality:=xlQualityStandard

    wbTemp.Close SaveChanges:=False
    Set wsTemp = Nothing
    Set wbTemp = Nothing
End Sub


'--------------------------------------------------------------
' CSV行パーサー（ダブルクォート対応）
'--------------------------------------------------------------
Private Function ParseCsvLine_COD(ByVal line As String) As String()
    If Len(line) > 0 And Right(line, 1) = vbCr Then
        line = Left(line, Len(line) - 1)
    End If

    Dim result() As String
    ReDim result(0 To 0)
    Dim fieldCount As Long
    fieldCount = 0

    Dim inQuote As Boolean
    inQuote = False
    Dim field As String
    field = ""
    Dim ch As String
    Dim lineLen As Long
    lineLen = Len(line)

    Dim pos As Long
    pos = 1

    Do While pos <= lineLen
        ch = Mid(line, pos, 1)

        If inQuote Then
            If ch = """" Then
                If pos < lineLen And Mid(line, pos + 1, 1) = """" Then
                    field = field & """"
                    pos = pos + 1
                Else
                    inQuote = False
                End If
            Else
                field = field & ch
            End If
        Else
            If ch = """" Then
                inQuote = True
            ElseIf ch = "," Then
                ReDim Preserve result(0 To fieldCount)
                result(fieldCount) = field
                fieldCount = fieldCount + 1
                field = ""
            Else
                field = field & ch
            End If
        End If

        pos = pos + 1
    Loop

    ReDim Preserve result(0 To fieldCount)
    result(fieldCount) = field

    ParseCsvLine_COD = result
End Function


'--------------------------------------------------------------
' SKU → カテゴリ名マッピング
'--------------------------------------------------------------
Private Function GetCodCategory_COD(ByVal sku As String) As String
    sku = LCase(Trim(sku))

    If InStr(sku, "01b") > 0 Or InStr(sku, "07b") > 0 Or InStr(sku, "cb") > 0 Then
        GetCodCategory_COD = "carboff"
    ElseIf InStr(sku, "ss") > 0 Then
        GetCodCategory_COD = "salasiru"
    ElseIf InStr(sku, "mc") > 0 Then
        GetCodCategory_COD = "Medico"
    ElseIf InStr(sku, "abura") > 0 Then
        GetCodCategory_COD = "abura"
    Else
        GetCodCategory_COD = ""
    End If
End Function


'--------------------------------------------------------------
' カテゴリ → 表示ラベル
'--------------------------------------------------------------
Private Function GetCodDisplayLabel_COD(ByVal category As String) As String
    Select Case LCase(category)
        Case "carboff":  GetCodDisplayLabel_COD = "【青】carboff"
        Case "salasiru": GetCodDisplayLabel_COD = "【茶】salasiru"
        Case "medico":   GetCodDisplayLabel_COD = "【紫】medico"
        Case "abura":    GetCodDisplayLabel_COD = "【白】油は太る"
        Case Else:       GetCodDisplayLabel_COD = category
    End Select
End Function


