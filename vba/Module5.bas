Attribute VB_Name = "Module5"
Option Explicit

Public Sub STORES変換()
    Dim inputFile As String
    Dim outputFolder As String
    Dim wbInput As Workbook
    Dim ws As Worksheet
    Dim headerRow As Range
    Dim colItem As Long, colZip As Long
    Dim colSei As Long, colMei As Long
    Dim colPref As Long, colAddr As Long
    Dim colOrderNumber As Long, colOrderDate As Long
    Dim lastRow As Long
    Dim dict As Object
    Dim cell As Range
    Dim item As Variant
    Dim idx As Long
    Dim fso As Object, ts As Object, tsSummary As Object
    Dim fields As Variant
    Dim lineParts As Variant
    Dim i As Long

    ' 入力CSVファイルを選択
    inputFile = Application.GetOpenFilename("CSV Files (*.csv),*.csv", , "入力CSVファイルを選択")
    If inputFile = "False" Then Exit Sub

    ' 出力先フォルダをデスクトップに固定
    outputFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    ' CSVをShift-JISで読み込む
    Workbooks.OpenText _
        fileName:=inputFile, _
        Origin:=932, _
        DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, _
        Comma:=True

    Set wbInput = ActiveWorkbook
    Set ws = wbInput.Sheets(1)

    ' ヘッダ行
    Set headerRow = ws.Rows(1)
    colItem = Application.Match("アイテム名", headerRow, 0)
    colZip = Application.Match("郵便番号(配送先)", headerRow, 0)
    colSei = Application.Match("氏(配送先)", headerRow, 0)
    colMei = Application.Match("名(配送先)", headerRow, 0)
    colPref = Application.Match("都道府県(配送先)", headerRow, 0)
    colAddr = Application.Match("住所(配送先)", headerRow, 0)
    colOrderNumber = Application.Match("オーダー番号", headerRow, 0)
    colOrderDate = Application.Match("オーダー日", headerRow, 0)

    ' データ最終行
    lastRow = ws.Cells(ws.Rows.Count, colItem).End(xlUp).row

    ' FileSystemObject 準備
    Set fso = CreateObject("Scripting.FileSystemObject")

    '── 1.csv（サマリ）を出力 ──
    Set tsSummary = fso.CreateTextFile(outputFolder & "\1.csv", True, False)
    ' ヘッダ行（カンマ区切り）
    tsSummary.WriteLine Join(Array( _
        "オーダー番号", "オーダー日", _
        "氏(配送先)", "名(配送先)", _
        "配送方法", "到着予定日時", "問い合わせ番号", "備考", _
        "発送完了" _
    ), ",")
    ' データ行（カンマ区切り）
    For i = 2 To lastRow
        tsSummary.WriteLine _
            ws.Cells(i, colOrderNumber).Value & "," & _
            Format(ws.Cells(i, colOrderDate).Value, "yyyy/m/d h:mm") & "," & _
            ws.Cells(i, colSei).Value & "," & _
            ws.Cells(i, colMei).Value & "," & _
            "日本郵便" & "," & _
            "" & "," & _
            "" & "," & _
            "" & "," & _
            "1"
    Next i
    tsSummary.Close
    '────────────────────────────────────

    ' アイテム名でソート（任意）
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Range(ws.Cells(2, colItem), ws.Cells(lastRow, colItem)), _
        SortOn:=xlSortOnValues, Order:=xlAscending
    With ws.Sort
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, ws.Columns.Count).End(xlToLeft))
        .Header = xlYes
        .Apply
    End With

    ' アイテム名をディクショナリに
    Set dict = CreateObject("Scripting.Dictionary")
    For Each cell In ws.Range(ws.Cells(2, colItem), ws.Cells(lastRow, colItem))
        If Not dict.Exists(cell.Value) Then dict.Add cell.Value, Empty
    Next cell

    ' 出力用ヘッダ配列
    fields = Array( _
        "お届け先郵便番号", "お届け先氏名", "お届け先敬称", _
        "お届け先住所1行目", "お届け先住所2行目", _
        "お届け先住所3行目", "お届け先住所4行目", "内容品" _
    )

    idx = 1
    ' 各アイテムごとにテキストファイルを生成
    For Each item In dict.Keys
        Set ts = fso.CreateTextFile(outputFolder & "\Book" & idx & ".csv", True, False)
        ' ヘッダ書き込み
        ts.WriteLine Join(fields, vbTab)
        ' データ行書き込み
        For Each cell In ws.Range(ws.Cells(2, colItem), ws.Cells(lastRow, colItem))
            If cell.Value = item Then
                lineParts = Array( _
                    ws.Cells(cell.row, colZip).Value, _
                    ws.Cells(cell.row, colSei).Value & ws.Cells(cell.row, colMei).Value, _
                    "様", _
                    ws.Cells(cell.row, colPref).Value, _
                    ws.Cells(cell.row, colAddr).Value, _
                    "", "", "書籍" _
                )
                ts.WriteLine Join(lineParts, vbTab)
            End If
        Next cell
        ts.Close
        idx = idx + 1
    Next item

    wbInput.Close SaveChanges:=False
   
    ' 1.csv を開く
    Workbooks.Open fileName:=outputFolder & "\1.csv"
End Sub


