Attribute VB_Name = "Module24"
Option Explicit

Public Sub サラシア変換()
    Dim filePath As String
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim lastRow As Long
    Dim i As Long
    
    '────────────────────────────
    ' 1. order_20で始まる最新CSVファイルを自動取得
    '────────────────────────────
    Dim folderPath As String
    Dim fileName As String
    Dim latestFile As String
    Dim latestDate As Date
    Dim fileDate As Date
    
    folderPath = "C:\Users\lenovo\Desktop\ダウンロード\"
    fileName = Dir(folderPath & "order_20*.csv")
    
    If fileName = "" Then
        MsgBox "order_20 で始まるCSVファイルが見つかりません。" & vbCrLf & _
               "フォルダ: " & folderPath, vbExclamation
        Exit Sub
    End If
    
    latestDate = #1/1/1900#
    latestFile = ""
    
    Do While fileName <> ""
        fileDate = FileDateTime(folderPath & fileName)
        If fileDate > latestDate Then
            latestDate = fileDate
            latestFile = fileName
        End If
        fileName = Dir()
    Loop
    
    filePath = folderPath & latestFile
    
    '────────────────────────────
    ' 2. 新規ブック＋QueryTable で Shift JIS インポート
    '    （全列を文字列扱いにして「2-5-16」を日付化させない）
    '────────────────────────────
    With Workbooks.Add(xlWBATWorksheet).Sheets(1)
        Set qt = .QueryTables.Add( _
            Connection:="TEXT;" & filePath, _
            Destination:=.Range("A1") _
        )
        With qt
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .TextFilePlatform = 932            ' Shift JIS
            ' 読み込む列数に合わせて要素数を調整（例:11列分）
            .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
            .Refresh BackgroundQuery:=False
        End With
        Set ws = .Parent.Sheets(1)
    End With
    
    ' シート名を「1」に（重複時は無視）
    On Error Resume Next
    ws.Name = "1"
    On Error GoTo 0
    
    '────────────────────────────
    ' 3. 既存の処理
    '────────────────────────────
    
    ' ─── ①「内容品」列 (I列) から「ご紹介」「追加」を削除＋Trim ───
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
    For i = 2 To lastRow
        With ws.Cells(i, "I")
            If Not IsError(.Value) And .Value <> "" Then
                .Value = Application.WorksheetFunction.Trim( _
                            Replace(Replace(.Value, "ご紹介", ""), "追加", "") _
                         )
            End If
        End With
    Next i
    
    ' ─── ②既存の置換処理 ───
    With ws.Range("I2:I" & lastRow)
        .Replace What:="20％引きの年間コース　【サラシア茶】", Replacement:="15％引きの3ヶ月毎コース　【サラシア茶】", LookAt:=xlPart
        .Replace What:="20％引きの年間コース　【サラシア粒】", Replacement:="15％引きの3ヶ月毎コース　【サラシア粒】", LookAt:=xlPart
        .Replace What:="0割引なしの単品購入　【サラシア茶】", Replacement:="10％引きの1ヶ月毎コース　【サラシア茶】", LookAt:=xlPart
        .Replace What:="0割引なしの単品購入　【サラシア粒】", Replacement:="10％引きの1ヶ月毎コース　【サラシア粒】", LookAt:=xlPart
    End With
    
    ' ─── ②-2 書籍名の末尾バリアント文字を除去 ───
    Const BOOK_PREFIX As String = "【紙の書籍】CARBOFF-糖質の吸収をおさえる方法"
    For i = 2 To lastRow
    With ws.Cells(i, "I")
        If Left(.Value, Len(BOOK_PREFIX)) = BOOK_PREFIX Then
            .Value = BOOK_PREFIX
        End If
    End With
    Next i
    
    ' ─── ③ 定期購入回数を正規化＋数値化 (J列) ───
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).row
    For i = 2 To lastRow
        With ws.Cells(i, "J")
            If Not IsEmpty(.Value) Then
                Dim v As Long
                v = CLng(Val(.Value))
                Select Case v
                    Case Is >= 2: v = 2
                    Case 0:       v = 1
                End Select
                .Value = v
                .NumberFormat = "0"
            End If
        End With
    Next i
    
    ' ─── ④並べ替え ───
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 key:=ws.Range("I2:I" & lastRow), Order:=xlAscending
        .SortFields.Add2 key:=ws.Range("J2:J" & lastRow), Order:=xlAscending
        .SortFields.Add2 key:=ws.Range("K2:K" & ws.Cells(ws.Rows.Count, "K").End(xlUp).row), Order:=xlDescending
        .SortFields.Add2 key:=ws.Range("B2:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).row), Order:=xlAscending
        .SetRange ws.Range("A1:K" & ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
        .Header = xlYes
        .Apply
    End With
    
    ' ─── ⑤「商品」列を H 列に挿入 ───
    ws.Columns("H").Insert Shift:=xlToRight
    ws.Range("H1").Value = "内容品"
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For i = 2 To lastRow
        If Application.WorksheetFunction.CountA(ws.Range("A" & i & ":K" & i)) > 0 Then
            ws.Cells(i, "H").Value = "商品"
        End If
    Next i
    
    ' ─── ⑥郵便番号を7桁に整形（A列） ───
    ws.Columns("A").NumberFormat = "@"
    For i = 2 To lastRow
        With ws.Cells(i, "A")
            If Not IsEmpty(.Value) Then
                Dim postalCode As String
                postalCode = Trim(.Value)
                If IsNumeric(postalCode) Then
                    .Value = Right("0000000" & postalCode, 7)
                End If
            End If
        End With
    Next i

    '────────────────────────────
    ' 4. 追加の書式設定
    '────────────────────────────

    ' ⑦ ヘッダー（A1:L1）の背景色と文字色
    With ws.Range("A1:L1")
        .Interior.Color = RGB(112, 173, 71)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' ⑧ 表全体（A1:L最終行）の罫線色
    With ws.Range("A1:L" & lastRow).Borders
        .LineStyle = xlContinuous
        .Color = RGB(146, 208, 80)
    End With

    ' ⑨ 全列幅を約73ピクセル相当に設定
    ws.Columns.ColumnWidth = 10.43
    
    ' ⑨-2 I列（購入数）を数値に変換
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
    For i = 2 To lastRow
        With ws.Cells(i, "I")
            If Not IsEmpty(.Value) And IsNumeric(.Value) Then
                .Value = CDbl(.Value)
            End If
        End With
    Next i
    ws.Columns("I").NumberFormat = "0"
    
    ' ⑩ J列のグループ切り替わり前にヘッダー行を挿入
    Dim headerRange As Range
    Set headerRange = ws.Range("A1:L1")
    lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).row

    For i = lastRow To 3 Step -1
        If ws.Cells(i, "J").Value <> ws.Cells(i - 1, "J").Value Then
            ws.Rows(i).Insert Shift:=xlDown
            headerRange.Copy Destination:=ws.Range("A" & i)
        End If
    Next i

    ' ─── ⑪ I～L列を削除 ───
    ws.Columns("I:L").Delete Shift:=xlToLeft
    
    ' ─── ⑫ CSV別名保存してブックを閉じる ───
    Application.DisplayAlerts = False
    ws.Parent.SaveAs fileName:="C:\Users\lenovo\Desktop\ダウンロード\carboff_red.csv", _
                     FileFormat:=xlCSV, _
                     Local:=True
    ws.Parent.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    ' ─── ⑬ Excelを終了 ───
    Application.Quit


End Sub


