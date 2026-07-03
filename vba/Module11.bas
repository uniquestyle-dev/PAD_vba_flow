Attribute VB_Name = "Module11"
Option Explicit

Public Sub Shopify登録確認_Shift_w()
    Dim filePathIn As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim paidAtCol As Long

    ' 1. 変換するファイルを選択（UTF-8 BOMなし）
    filePathIn = Application.GetOpenFilename( _
        FileFilter:="CSV ファイル (*.csv),*.csv", _
        Title:="変換する UTF-8 CSV を選択してください" _
    )
    If filePathIn = False Then Exit Sub

    ' 2. 新規ワークブック + QueryTable で UTF-8 インポート
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Sheets(1)
    Set qt = ws.QueryTables.Add( _
        Connection:="TEXT;" & filePathIn, _
        Destination:=ws.Range("A1") _
    )
    With qt
        .TextFileParseType = xlDelimited
        .TextFilePlatform = 65001                    ' UTF-8
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileCommaDelimiter = True
        .Refresh BackgroundQuery:=False              ' 同期実行
    End With

    ' ※ QueryTable を削除せずそのままデータを利用します

    ' 3. "Paid at" 列を探して降順ソート
    On Error Resume Next
    paidAtCol = Application.Match("Paid at", ws.Rows(1), 0)
    On Error GoTo 0
    If paidAtCol = 0 Then
        MsgBox "ヘッダーに ""Paid at"" 列が見つかりません。", vbExclamation
        Exit Sub
    End If
    ' ヘッダー行ありのまま、2行目以降をキー列で降順に並べ替え
    ws.UsedRange.Sort _
        Key1:=ws.Cells(2, paidAtCol), _
        Order1:=xlDescending, _
        Header:=xlYes

    ' 4. ソート後に A1 セルを選択
    ws.Range("A1").Select
End Sub

