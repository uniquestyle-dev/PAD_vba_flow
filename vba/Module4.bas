Attribute VB_Name = "Module4"
Option Explicit

Public Sub 代引き手数料追加_CS_y(Optional filePathIn As Variant = "")
    ' ↓ filePathIn の Dim は削除（引数になったので）
    ' Dim filePathIn   As Variant    ← これを消す
    Dim filePathOut  As String       ' ← これ以降は残す
    Dim wb           As Workbook
    Dim ws           As Worksheet
    Dim qt           As QueryTable
    Dim lastRow      As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim headerLine   As String
    Dim cols         As Variant
    Dim types()      As Integer
    Dim adoStream    As Object
    Dim csvLine      As String
    
    If filePathIn = "" Then
        filePathIn = Application.GetOpenFilename( _
            FileFilter:="CSV ファイル (*.csv),*.csv", _
            Title:="変換する UTF-8 CSV を選択してください" _
        )
        If filePathIn = False Then Exit Sub
    End If
    
    ' 2. 新規ワークブック＋QueryTable で UTF-8 インポート
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Sheets(1)
    
    ' ── ヘッダー行を読み込んで列数を取得 ──
    Open filePathIn For Input As #1
        Line Input #1, headerLine
    Close #1
    cols = Split(headerLine, ",")
    ReDim types(1 To UBound(cols) + 1)
    
    ' ① まず全列を「一般」に
    For i = LBound(types) To UBound(types)
        types(i) = xlGeneralFormat
    Next i
    
    ' ② テキスト扱いにしたい列をヘッダー名で指定
    Dim colName As String
    For i = LBound(cols) To UBound(cols)
        colName = Replace(cols(i), """", "")
        If colName = "Shipping Street" _
           Or colName = "Shipping Address1" _
           Or colName = "Shipping Zip" _
           Or colName = "Shipping Phone" Then
            types(i + 1) = xlTextFormat
        End If
    Next i
    
    ' ③ QueryTable の生成・インポート
    Set qt = ws.QueryTables.Add( _
        Connection:="TEXT;" & filePathIn, _
        Destination:=ws.Range("A1") _
    )
    With qt
        .TextFileParseType = xlDelimited
        .TextFilePlatform = 65001
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = types
        .Refresh BackgroundQuery:=False
        On Error Resume Next
            .Delete
        On Error GoTo 0
    End With
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' ──────────────────────────────────────────────
    ' 3.5  Lineitem name 列の値を対応表に基づき変換
    ' ──────────────────────────────────────────────
    Dim liCol As Long
    liCol = 0
    
    ' ヘッダー行から "Lineitem name" 列を検索
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For j = 1 To lastCol
        If Trim(ws.Cells(1, j).Value) = "Lineitem name" Then
            liCol = j
            Exit For
        End If
    Next j
    
    If liCol > 0 Then
        ' 変換対応表（変換前 → 変換後）
        Dim mapFrom As Variant, mapTo As Variant
        mapFrom = Array( _
            "【印刷版】「油が太る」は本当か？", _
            "【印刷版】CARBOFF-糖質の吸収をおさえる方法", _
            "【紙書籍】Medico-痩せ薬で○kg痩せた話", _
            "【紙書籍】Salasiru-効率がいいサラシア摂取法" _
        )
        mapTo = Array( _
            "【白】油が太る…", _
            "【青】CARBOFF", _
            "【紫】Medico", _
            "【茶】Salasiru" _
        )
        
        Dim cellVal As String
        Dim k As Long
        For i = 2 To lastRow
            cellVal = Trim(CStr(ws.Cells(i, liCol).Value))
            For k = LBound(mapFrom) To UBound(mapFrom)
                If cellVal = mapFrom(k) Then
                    ws.Cells(i, liCol).Value = mapTo(k)
                    Exit For
                End If
            Next k
        Next i
    End If
    
    ' 4. L列に290を加算
    For i = 1 To lastRow
        If IsNumeric(ws.Cells(i, "L").Value) Then
            ws.Cells(i, "L").Value = ws.Cells(i, "L").Value + 290
        End If
    Next i
    
    ' 5. AR列：文字列フォーマット＆「+81→'0」に置換
    With ws.Columns("AR")
        .NumberFormat = "@"
        .Replace What:="+81", _
                 Replacement:="'0", _
                 LookAt:=xlPart, MatchCase:=False
    End With
    
    ' 6. 出力先パス（デスクトップ）
    filePathOut = Environ("USERPROFILE") & "\Desktop\ダウンロード\代引き.csv"
    
    ' 7. ADODB.Stream で Shift_JIS 書き出し
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set adoStream = CreateObject("ADODB.Stream")
    With adoStream
        .Type = 2
        .Charset = "Shift_JIS"
        .Open
        For i = 1 To lastRow
            If Application.WorksheetFunction.CountA( _
                   ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)) _
               ) > 0 Then
                csvLine = ""
                For j = 1 To lastCol
                    csvLine = csvLine & """" & _
                              Replace(ws.Cells(i, j).Text, """", """""") & """"
                    If j < lastCol Then csvLine = csvLine & ","
                Next j
                .WriteText csvLine, 1
            End If
        Next i
        .SaveToFile filePathOut, 2
        .Close
    End With
    
    ' 8. 終了
    wb.Close SaveChanges:=False
    Application.Quit
End Sub




