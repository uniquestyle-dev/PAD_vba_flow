Attribute VB_Name = "Module10"
Option Explicit

Public Sub Shopify_件数集計()
    Dim filePathIn As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim qt As QueryTable

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

    ' 3. 不要列を削除
    With ws
        .Columns("A:A").Delete Shift:=xlToLeft
        .Columns("B:N").Delete Shift:=xlToLeft
        .Columns("C:F").Delete Shift:=xlToLeft
    End With

    ' 4. フィルターを設定（1行目を見出し行として）
    ws.Rows(1).AutoFilter

    ' 4-1. B列「Created at」で降順ソート
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Columns("B"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange ws.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' 5. D8 セルを選択
    ws.Range("D8").Select
End Sub


