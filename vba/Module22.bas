Attribute VB_Name = "Module22"
Option Explicit

Public Sub DaibikiConvert(filePathIn As String)

    Dim filePathOut  As String
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

    ' ============================================================
    ' 0. 引数チェック
    ' ============================================================
    If filePathIn = "" Then
        MsgBox "ファイルパスが渡されていません。PAD側の変数を確認してください。", vbCritical, "代引き変換エラー"
        Exit Sub
    End If
    If Dir(filePathIn) = "" Then
        MsgBox "ファイルが見つかりません:" & vbCrLf & filePathIn, vbCritical, "代引き変換エラー"
        Exit Sub
    End If

    filePathOut = "C:\Users\lenovo\Desktop\ダウンロード\代引き.csv"

    ' ============================================================
    ' 2. ヘッダー行を読み込んで列データ型を設定
    ' ============================================================
    On Error GoTo ErrHandler
    Open filePathIn For Input As #1
        Line Input #1, headerLine
    Close #1

    cols = Split(headerLine, ",")
    ReDim types(1 To UBound(cols) + 1)

    Dim k As Long
    For k = LBound(types) To UBound(types)
        types(k) = xlGeneralFormat
    Next k

    Dim colName As String
    For k = LBound(cols) To UBound(cols)
        colName = Replace(cols(k), """", "")
        Select Case colName
            Case "Shipping Street", "Shipping Address1", "Shipping Zip", "Shipping Phone"
                types(k + 1) = xlTextFormat
        End Select
    Next k

    ' ============================================================
    ' 3. 新規ワークブック＋QueryTable で UTF-8 インポート
    ' ============================================================
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Sheets(1)

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
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "データが読み込めませんでした（1行以下）。" & vbCrLf & _
               "ファイル: " & filePathIn, vbCritical, "代引き変換エラー"
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    ' ============================================================
    ' 4. Lineitem name 列を変換
    ' ============================================================
    Dim liCol As Long
    liCol = 0
    For j = 1 To lastCol
        If Trim(ws.Cells(1, j).Value) = "Lineitem name" Then
            liCol = j
            Exit For
        End If
    Next j

    If liCol > 0 Then
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

    ' ============================================================
    ' 5. L列に290を加算（ヘッダー行はスキップ）
    ' ============================================================
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, "L").Value) And ws.Cells(i, "L").Value <> "" Then
            ws.Cells(i, "L").Value = ws.Cells(i, "L").Value + 290
        End If
    Next i

    ' ============================================================
    ' 6. AR列：+81 → 0 に置換（文字列フォーマット）
    ' ============================================================
    With ws.Columns("AR")
        .NumberFormat = "@"
        .Replace What:="+81", _
                 Replacement:="0", _
                 LookAt:=xlPart, MatchCase:=False
    End With

    ' ============================================================
    ' 7. ADODB.Stream で Shift_JIS CSV 書き出し
    ' ============================================================
    Set adoStream = CreateObject("ADODB.Stream")
    With adoStream
        .Type = 2             ' adTypeText
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
                .WriteText csvLine, 1   ' 1 = adWriteLine（改行付き）
            End If
        Next i

        .SaveToFile filePathOut, 2      ' 2 = adSaveCreateOverWrite
        .Close
    End With
    
    wb.Close SaveChanges:=False
    Exit Sub

' ============================================================
' エラーハンドラ
' ============================================================
ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "予期しないエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "代引き変換エラー"

End Sub
