Attribute VB_Name = "Module30"
Option Explicit

Public Sub サラシア変換サプリ一覧表()
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
    
    ' --- 【追加】CARBOFF×代引き：氏名抽出 / 件数 / Book_Shipment_List ---
    '   ※取込直後の生データ列構成:
    '       B列=お届け先氏名 / I列=内容品(CARBOFFを含む) / K列=支払方法(「代引き」を含む)
    Dim codCount As Long
    Dim codLastRow As Long
    Dim ffCodName As Integer, ffCodCount As Integer

    ' Book_Shipment_List.xlsx 用ブック（ヘッダのみでも保存する）
    Dim wbBook As Workbook, wsBook As Worksheet
    Dim bookRow As Long
    Set wbBook = Workbooks.Add(xlWBATWorksheet)
    Set wsBook = wbBook.Sheets(1)
    wsBook.Columns("A:C").NumberFormat = "@"
    wsBook.Range("A1").Value = "本の色"
    wsBook.Range("B1").Value = "名前"
    wsBook.Range("C1").Value = "支払方法"
    bookRow = 2

    codCount = 0
    codLastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row

    ' red_cod_name.txt を新規作成（該当行のお届け先氏名を1件1行で出力）
    ffCodName = FreeFile
    Open folderPath & "red_cod_name.txt" For Output As #ffCodName
    For i = 2 To codLastRow
        If InStr(ws.Cells(i, "I").Value, "CARBOFF") > 0 _
           And InStr(ws.Cells(i, "K").Value, "代引き") > 0 Then
            Print #ffCodName, ws.Cells(i, "B").Value
            codCount = codCount + 1
            ' Book_Shipment_List 行追加（本の色=赤 / 名前=顧客名 / 支払方法=代引き）
            wsBook.Cells(bookRow, 1).Value = "赤"
            wsBook.Cells(bookRow, 2).Value = ws.Cells(i, "B").Value
            wsBook.Cells(bookRow, 3).Value = "代引き"
            bookRow = bookRow + 1
        End If
    Next i
    Close #ffCodName

    ' books_cod_count.txt を新規作成（5行：1?4行目=0、5行目=該当件数）
    ffCodCount = FreeFile
    Open folderPath & "books_cod_count.txt" For Output As #ffCodCount
    Print #ffCodCount, "0"
    Print #ffCodCount, "0"
    Print #ffCodCount, "0"
    Print #ffCodCount, "0"
    Print #ffCodCount, CStr(codCount)
    Close #ffCodCount

    ' Book_Shipment_List.xlsx を保存（該当0件ならヘッダ行のみで保存）
    Application.DisplayAlerts = False
    wbBook.SaveAs fileName:=folderPath & "Book_Shipment_List.xlsx", _
                  FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbBook.Close SaveChanges:=False
    ' --- 【追加ここまで】 ---
    
    
    
    
    ' --- I列に CARBOFF を含む行を削除 ---
    Dim qDelRow As Long
    qDelRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).row
    For i = qDelRow To 2 Step -1
    If InStr(ws.Cells(i, "I").Value, "CARBOFF") > 0 Then
            ws.Rows(i).Delete
        End If
    Next i
    
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
    
    ' ??????????????????????????????????????????????
    ' ⑩-B  サラシア数量リスト（印刷 → xlsx保存）
    ' ??????????????????????????????????????????????
    
    ' --- 現在のシートを丸ごと新規ブックにコピー ---
    ws.Copy
    Dim wbQty As Workbook:  Set wbQty = ActiveWorkbook
    Dim wsQty As Worksheet: Set wsQty = wbQty.Sheets(1)
    Dim savePath As String
    savePath = "C:\Users\lenovo\Desktop\ダウンロード\"
    
    ' --- 不要列を削除（右→左の順でシフトの影響を回避） ---
    '   元12列 A-L から B,I,J,L だけ残す
    '   K(購入回数) → H(挿入された内容品) → C:G(住所等) → A(郵便番号)
    wsQty.Columns("K:K").Delete Shift:=xlToLeft
    wsQty.Columns("H:H").Delete Shift:=xlToLeft
    wsQty.Columns("C:G").Delete Shift:=xlToLeft
    wsQty.Columns("A:A").Delete Shift:=xlToLeft
    ' 結果: A=氏名等  B=購入数  C=内容品(商品名)  D=支払方法
    
    Dim qLastRow As Long
    qLastRow = wsQty.Cells(wsQty.Rows.Count, "C").End(xlUp).row
    
    ' --- C列 → 大／小 ---
    Dim qCell As Range
    For Each qCell In wsQty.Range("C2:C" & qLastRow)
        If qCell.Value Like "*15" & ChrW(&HFF05) & "*" Or _
           qCell.Value Like "*20" & ChrW(&HFF05) & "*" Then
            qCell.Value = ChrW(&H5927)          ' 大
        ElseIf qCell.Value Like "*10" & ChrW(&HFF05) & "*" Or _
               qCell.Value Like "*" & ChrW(&H5272) & ChrW(&H5F15) & "*" Then
            qCell.Value = ChrW(&H5C0F)          ' 小
        End If
    Next qCell
    
    ' --- A列・B列の背景白 ---
    wsQty.Range("A2:B" & qLastRow).Interior.Color = RGB(255, 255, 255)
    
    ' --- B列 ? 2 → 蛍光黄 ---
    For Each qCell In wsQty.Range("B2:B" & qLastRow)
        If IsNumeric(qCell.Value) And qCell.Value >= 2 Then
            qCell.Interior.Color = RGB(255, 255, 0)
        End If
    Next qCell
    
    ' --- C列の背景色（大=薄緑 / 小=白） ---
    For Each qCell In wsQty.Range("C2:C" & qLastRow)
        If qCell.Value = ChrW(&H5927) Then
            qCell.Interior.Color = RGB(226, 239, 218)
        ElseIf qCell.Value = ChrW(&H5C0F) Then
            qCell.Interior.Color = RGB(255, 255, 255)
        End If
    Next qCell
    
    ' --- B列の左右罫線 ---
    With wsQty.Range("B2:B" & qLastRow).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(226, 239, 218)
    End With
    With wsQty.Range("B2:B" & qLastRow).Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(226, 239, 218)
    End With
    
    ' --- D列→E列にコピー（案内状★ / 支払い方法） ---
    wsQty.Columns("D:D").Copy Destination:=wsQty.Columns("E:E")
    
    ' D列 → 案内状★
    wsQty.Range("D1").Value = ChrW(&H6848) & ChrW(&H5185) & ChrW(&H72B6) & ChrW(&H2605)
    Dim qLastRowD As Long
    qLastRowD = wsQty.Cells(wsQty.Rows.Count, "D").End(xlUp).row
    For Each qCell In wsQty.Range("D2:D" & qLastRowD)
        If qCell.Value = ChrW(&H30B3) & ChrW(&H30F3) & ChrW(&H30D3) & ChrW(&H30CB) & _
           ChrW(&H5F8C) & ChrW(&H6255) & ChrW(&H3044) & "-" & ChrW(&H2605) Then
            qCell.Value = ChrW(&H2605)
        Else
            qCell.ClearContents
        End If
    Next qCell
    wsQty.Range("D2:D" & qLastRowD).Interior.Color = RGB(255, 255, 255)
    With wsQty.Range("D2:D" & qLastRowD).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(226, 239, 218)
    End With
    
    ' E列 → 支払い方法
    wsQty.Range("E1").Value = ChrW(&H652F) & ChrW(&H6255) & ChrW(&H3044) & ChrW(&H65B9) & ChrW(&H6CD5)
    With wsQty.Range("E1")
        .Interior.Color = RGB(112, 173, 71)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    Dim qLastRowE As Long
    qLastRowE = wsQty.Cells(wsQty.Rows.Count, "E").End(xlUp).row
    Dim origVal As String
    For Each qCell In wsQty.Range("E2:E" & qLastRowE)
        origVal = qCell.Value
        If InStr(origVal, "-" & ChrW(&H2605)) > 0 Then
            origVal = Replace(origVal, "-" & ChrW(&H2605), "")
        End If
        Select Case origVal
            Case ChrW(&H30B3) & ChrW(&H30F3) & ChrW(&H30D3) & ChrW(&H30CB) & _
                 ChrW(&H5F8C) & ChrW(&H6255) & ChrW(&H3044), _
                 ChrW(&H30B3) & ChrW(&H30F3) & ChrW(&H30D3) & ChrW(&H30CB) & _
                 ChrW(&H652F) & ChrW(&H6255) & ChrW(&H3044) & "-" & _
                 ChrW(&H540C) & ChrW(&H68B1)
                qCell.Value = ChrW(&H30B3) & ChrW(&H30F3) & ChrW(&H30D3) & ChrW(&H30CB)
            Case ChrW(&H4EE3) & ChrW(&H5F15) & ChrW(&H304D), _
                 ChrW(&H30B3) & ChrW(&H30F3) & ChrW(&H30D3) & ChrW(&H30CB)
                ' そのまま
            Case Else
                qCell.ClearContents
        End Select
    Next qCell
    wsQty.Range("E2:E" & qLastRowE).Interior.Color = RGB(255, 255, 255)
    
    ' E列罫線
    Dim bdr As Variant
    For Each bdr In Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom, xlInsideHorizontal)
        With wsQty.Range("E1:E" & qLastRowE).Borders(bdr)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(169, 208, 142)
        End With
    Next bdr
    
    ' --- 列幅自動調整 ---
    wsQty.Rows("2:" & qLastRow).Hidden = True
    wsQty.Columns("C:E").AutoFit
    wsQty.Rows("2:" & qLastRow).Hidden = False
    wsQty.Columns("A:A").AutoFit
    
    ' --- 重複ヘッダー行の削除 ---
    Dim qLastCol As Long, qDataRow As Long, qj As Long, qIsHdr As Boolean
    qLastCol = wsQty.Cells(1, wsQty.Columns.Count).End(xlToLeft).Column
    qDataRow = wsQty.Cells(wsQty.Rows.Count, "A").End(xlUp).row
    For i = qDataRow To 2 Step -1
        qIsHdr = True
        For qj = 1 To qLastCol
            If wsQty.Cells(i, qj).Value <> wsQty.Cells(1, qj).Value Then
                qIsHdr = False: Exit For
            End If
        Next qj
        If qIsHdr Then wsQty.Rows(i).Delete
    Next i
    For i = wsQty.Cells(wsQty.Rows.Count, "C").End(xlUp).row To 2 Step -1
        If wsQty.Cells(i, "C").Value = wsQty.Cells(1, "C").Value Then
            wsQty.Rows(i).Delete
        End If
    Next i
    
    ' --- サラシア小・大の購入数合計 → salacia_count.txt ---
    Dim countSmall As Long, countLarge As Long
    Dim qCountRow As Long
    qCountRow = wsQty.Cells(wsQty.Rows.Count, "C").End(xlUp).row
    countSmall = 0
    countLarge = 0
    For Each qCell In wsQty.Range("C2:C" & qCountRow)
        If qCell.Value = ChrW(&H5C0F) Then
            countSmall = countSmall + Val(wsQty.Cells(qCell.row, "B").Value)
        ElseIf qCell.Value = ChrW(&H5927) Then
            countLarge = countLarge + Val(wsQty.Cells(qCell.row, "B").Value)
        End If
    Next qCell

    Dim ff As Integer
    ff = FreeFile
    Open savePath & "salacia_count.txt" For Output As #ff
    Print #ff, countSmall
    Print #ff, countLarge
    Close #ff
        
    ' --- カラー印刷 ---
     Dim qPrintRow As Long
     qPrintRow = wsQty.Cells(wsQty.Rows.Count, "A").End(xlUp).row
     With wsQty.PageSetup
         .PrintArea = "A1:E" & qPrintRow
         .Zoom = False
         .FitToPagesWide = 1
         .FitToPagesTall = False
         .Orientation = xlPortrait
         .TopMargin = Application.CentimetersToPoints(1)
         .BottomMargin = Application.CentimetersToPoints(1)
         .LeftMargin = Application.CentimetersToPoints(1)
         .RightMargin = Application.CentimetersToPoints(1)
     End With
'    wsQty.PrintOut Copies:=1
    
    ' --- xlsx保存 → 閉じる ---
    Application.DisplayAlerts = False
    wbQty.SaveAs fileName:=savePath & "Salacia_Shipment_List.xlsx", _
                  FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    ' ★ PDF保存（追加）
    wsQty.ExportAsFixedFormat Type:=xlTypePDF, _
    fileName:=savePath & "Salacia_Shipment_List.pdf", _
    Quality:=xlQualityStandard, _
    OpenAfterPublish:=False
    
    wbQty.Close SaveChanges:=False
    
    ' ─── 元の作業ブックを保存せず閉じる ───
    ws.Parent.Close SaveChanges:=False
    

End Sub






