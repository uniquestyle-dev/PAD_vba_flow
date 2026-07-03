Attribute VB_Name = "Module31"
'--------------------------------------------------------------
' AppendCodNames
' orders_export.csv（Shopify代引き注文エクスポート）から
' SKUに基づいて名前を商品カテゴリ別のテキストファイルに出力する。
'
' 出力:
'   C:\Users\lenovo\Desktop\ダウンロード\carboff_cod_names.txt
'   C:\Users\lenovo\Desktop\ダウンロード\salasiru_cod_names.txt
'   C:\Users\lenovo\Desktop\ダウンロード\Medico_cod_names.txt
'   C:\Users\lenovo\Desktop\ダウンロード\abura_cod_names.txt
'   ※ 該当なしのカテゴリはファイルを生成しない
'
' orders_export.csv が存在しない場合は何もしない。
'
' ExportNamePdf 側で _cod_names.txt を読んでA列末尾に追記する。
'--------------------------------------------------------------
Sub AppendCodNames()
    Const DL_DIR = "C:\Users\lenovo\Desktop\ダウンロード\"
    Const CSV_NAME = "orders_export.csv"
    Const LOG_PATH = "C:\Users\lenovo\Desktop\cod_names_debug.txt"
    
    Dim f As Integer
    On Error GoTo ErrHandler
    
    Dim csvPath As String
    csvPath = DL_DIR & CSV_NAME
    
    ' orders_export.csv が存在しなければスキップ
    If Dir(csvPath) = "" Then
        f = FreeFile
        Open LOG_PATH For Output As #f
        Print #f, Now & " - orders_export.csv なし（スキップ）"
        Close #f
        Exit Sub
    End If
    
    f = FreeFile
    Open LOG_PATH For Output As #f
    Print #f, Now & " - Start"
    Print #f, "csvPath: " & csvPath
    Close #f
    
    ' --- UTF-8でCSV読み込み ---
    Dim adoStream As Object
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Open
    adoStream.Type = 2  ' adTypeText
    adoStream.Charset = "utf-8"
    adoStream.LoadFromFile csvPath
    Dim csvText As String
    csvText = adoStream.ReadText
    adoStream.Close
    Set adoStream = Nothing
    
    ' BOM除去
    If Len(csvText) > 0 And AscW(Left(csvText, 1)) = &HFEFF Then
        csvText = Mid(csvText, 2)
    End If
    
    ' 改行統一 → 行分割
    csvText = Replace(csvText, vbCrLf, vbLf)
    csvText = Replace(csvText, vbCr, vbLf)
    Dim lines() As String
    lines = Split(csvText, vbLf)
    
    If UBound(lines) < 1 Then
        f = FreeFile
        Open LOG_PATH For Append As #f
        Print #f, "データ行なし（スキップ）"
        Close #f
        Exit Sub
    End If
    
    ' --- ヘッダー解析 ---
    Dim headerFields() As String
    headerFields = ParseCsvLine(lines(0))
    
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
    
    If colSku = -1 Or colShipName = -1 Then
        f = FreeFile
        Open LOG_PATH For Append As #f
        Print #f, "必須列が見つからない (SKU=" & colSku & ", ShipName=" & colShipName & ")"
        Close #f
        Exit Sub
    End If
    
    ' --- カテゴリ別Dictionary ---
    ' key=カテゴリ名, value=Dictionary(注文番号|名前 → 名前)
    Dim categories As Object
    Set categories = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(lines)
        If Len(Trim(lines(i))) = 0 Then GoTo NextRow
        
        Dim fields() As String
        fields = ParseCsvLine(lines(i))
        
        If UBound(fields) < colSku Or UBound(fields) < colShipName Then GoTo NextRow
        
        Dim sku As String
        sku = Trim(fields(colSku))
        
        Dim category As String
        category = GetCodCategory(sku)
        
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
    
    ' --- カテゴリ別テキストファイル出力 ---
    f = FreeFile
    Open LOG_PATH For Append As #f
    
    Dim catKey As Variant
    For Each catKey In categories.Keys
        Dim outPath As String
        outPath = DL_DIR & CStr(catKey) & "_cod_names.txt"
        
        Dim fOut As Integer
        fOut = FreeFile
        Open outPath For Output As #fOut
        
        Dim nameKey As Variant
        For Each nameKey In categories(catKey).Keys
            Print #fOut, categories(catKey)(nameKey)
        Next nameKey
        
        Close #fOut
        Print #f, CStr(catKey) & ": " & categories(catKey).Count & "件 -> " & outPath
    Next catKey
    
    Print #f, "完了"
    Close #f
    
    Exit Sub

ErrHandler:
    f = FreeFile
    Open LOG_PATH For Append As #f
    Print #f, "ERROR " & Err.Number & ": " & Err.Description
    Close #f
End Sub


'--------------------------------------------------------------
' CSV行パーサー（ダブルクォート対応）
'--------------------------------------------------------------
Private Function ParseCsvLine(ByVal line As String) As String()
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
    
    ParseCsvLine = result
End Function


'--------------------------------------------------------------
' SKU → カテゴリ名マッピング（既存カウントマクロと同一ロジック）
'--------------------------------------------------------------
Private Function GetCodCategory(ByVal sku As String) As String
    sku = LCase(Trim(sku))
    
    If InStr(sku, "01b") > 0 Or InStr(sku, "07b") > 0 Or InStr(sku, "cb") > 0 Then
        GetCodCategory = "carboff"
    ElseIf InStr(sku, "ss") > 0 Then
        GetCodCategory = "salasiru"
    ElseIf InStr(sku, "mc") > 0 Then
        GetCodCategory = "Medico"
    ElseIf InStr(sku, "abura") > 0 Then
        GetCodCategory = "abura"
    Else
        GetCodCategory = ""
    End If
End Function

