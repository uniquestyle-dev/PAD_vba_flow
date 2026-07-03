Attribute VB_Name = "Module37"
Option Explicit

'==================================================================
' 【一覧表PDF用】STORES注文CSV → Book_Shipment_List.xlsx へ追記（単独実行マクロ）
'   ・既存の「STORES変換」(ClickPost/1.csv出力)とは別名・別モジュール想定。
'     既存処理には一切影響しません。必ず新しい標準モジュールに配置すること。
'   ・stores-list.csv の出力は廃止し、既存の Book_Shipment_List.xlsx に行を追記する。
'     列構成は 本の色 / 名前 / 支払方法（Salaciaマクロが作るヘッダと同一）。
'   ・前提: Salaciaマクロが先に Book_Shipment_List.xlsx を作成・リセット済みであること。
'     ファイルが無い場合はヘッダ付きで新規作成する。
'==================================================================
Public Sub STORES変換一覧表用()
    Dim inputFile As String
    Dim outputFolder As String
    Dim wbInput As Workbook
    Dim ws As Worksheet
    Dim colItem As Long, colZip As Long
    Dim colSei As Long, colMei As Long
    Dim colPref As Long, colAddr As Long
    Dim colOrderNumber As Long, colOrderDate As Long
    Dim lastRow As Long
    Dim fso As Object
    Dim i As Long
    Dim headerVal As String
    Dim col As Long
    Dim itemName As String

    '─── 入力CSVファイルを自動取得（作成日時が最新の20*.csv）※読み込み部は変更なし ───
    Dim dlFolder As Object, dlFile As Object, latestDate As Date
    Dim dlPath As String
    dlPath = "C:\Users\lenovo\Desktop\ダウンロード\"
    latestDate = 0
    inputFile = ""

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dlFolder = fso.GetFolder(dlPath)
    For Each dlFile In dlFolder.Files
        If LCase(dlFile.Name) Like "20*.csv" Then
            If dlFile.DateCreated > latestDate Then
                latestDate = dlFile.DateCreated
                inputFile = dlFile.Path
            End If
        End If
    Next dlFile

    If inputFile = "" Then
        Exit Sub
    End If

    Debug.Print "ファイル: " & inputFile

    ' 出力先フォルダ
    outputFolder = "C:\Users\lenovo\Desktop\ダウンロード\"

    ' CSVを読み込む（※読み込み部は変更なし）
    Set wbInput = Workbooks.Add(xlWBATWorksheet)
    Set ws = wbInput.Sheets(1)

    Dim qt As QueryTable
    Set qt = ws.QueryTables.Add("TEXT;" & inputFile, ws.Range("A1"))
    With qt
        .TextFilePlatform = 932
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTextQualifier = xlTextQualifierDoubleQuote

        Dim tArr(0 To 255) As Variant
        For i = 0 To 255: tArr(i) = xlTextFormat: Next i
        .TextFileColumnDataTypes = tArr

        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
    End With

    ' ヘッダ行から列を自動検索（※検出部は変更なし）
    colItem = 0: colZip = 0: colSei = 0: colMei = 0
    colPref = 0: colAddr = 0: colOrderNumber = 0: colOrderDate = 0

    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        headerVal = ws.Cells(1, col).Value
        Select Case headerVal
            Case "アイテム名":                       colItem = col
            Case "郵便番号(購入者)", "郵便番号(配送先)":  colZip = col
            Case "氏(購入者)", "氏(配送先)":            colSei = col
            Case "名(購入者)", "名(配送先)":            colMei = col
            Case "都道府県(購入者)", "都道府県(配送先)":  colPref = col
            Case "住所(購入者)", "住所(配送先)":         colAddr = col
            Case "オーダー番号":                       colOrderNumber = col
            Case "オーダー日時", "オーダー日":           colOrderDate = col
        End Select
    Next col

    ' データ最終行
    lastRow = ws.Cells(ws.Rows.Count, colItem).End(xlUp).row

    '── Book_Shipment_List.xlsx に追記（CSV出力は廃止し統合）──
    Dim bookPath As String
    bookPath = outputFolder & "Book_Shipment_List.xlsx"

    Dim wbBook As Workbook, wsBook As Worksheet
    If fso.FileExists(bookPath) Then
        Set wbBook = Workbooks.Open(bookPath)
        Set wsBook = wbBook.Sheets(1)
    Else
        ' 無ければヘッダ付きで新規作成（本来はSalaciaマクロが先に作成する想定）
        Set wbBook = Workbooks.Add(xlWBATWorksheet)
        Set wsBook = wbBook.Sheets(1)
        wsBook.Columns("A:C").NumberFormat = "@"
        wsBook.Range("A1").Value = "本の色"
        wsBook.Range("B1").Value = "名前"
        wsBook.Range("C1").Value = "支払方法"
    End If

    ' 追記開始行（A列の最終行の次）
    Dim bookRow As Long
    bookRow = wsBook.Cells(wsBook.Rows.Count, "A").End(xlUp).row + 1
    If bookRow < 2 Then bookRow = 2

    ' データ行（本の色=対応表変換 / 名前=氏+名 / 支払方法=空欄固定）
    For i = 2 To lastRow
        itemName = Trim(ws.Cells(i, colItem).Value)
        wsBook.Cells(bookRow, 1).Value = ConvertItemName(itemName)
        wsBook.Cells(bookRow, 2).Value = ws.Cells(i, colSei).Value & ws.Cells(i, colMei).Value
        wsBook.Cells(bookRow, 3).Value = ""
        bookRow = bookRow + 1
    Next i

    Application.DisplayAlerts = False
    wbBook.SaveAs fileName:=bookPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbBook.Close SaveChanges:=False

    wbInput.Close SaveChanges:=False

    Debug.Print "Book_Shipment_List.xlsx に追記: " & (lastRow - 1) & "件"
    Debug.Print "===== 完了 ====="
End Sub


'----------------------------------------------
' アイテム名変換（変換対応表.xlsx の内容をハードコード）
'   変換前 → 変換後
'   ※未定義のアイテム名は変換せずそのまま出力（取りこぼし防止）
'----------------------------------------------
Private Function ConvertItemName(ByVal itemName As String) As String
    Select Case itemName
        Case "電子書籍印刷サービス_c":   ConvertItemName = "青"
        Case "電子書籍印刷サービス_s":   ConvertItemName = "茶"
        Case "電子書籍印刷サービス_sr":  ConvertItemName = "茶"
        Case "電子書籍印刷サービス_i":   ConvertItemName = "茶のおまけ赤だけ"
        Case "電子書籍印刷サービス_sw":  ConvertItemName = "茶のおまけ緑だけ"
        Case "電子書籍印刷サービス_set": ConvertItemName = "茶のおまけ赤緑だけ"
        Case "電子書籍印刷サービス_m":   ConvertItemName = "紫"
        Case "電子書籍印刷サービス_mr":  ConvertItemName = "紫"
        Case Else:                       ConvertItemName = itemName
    End Select
End Function

