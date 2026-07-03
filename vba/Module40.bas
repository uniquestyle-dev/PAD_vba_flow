Attribute VB_Name = "Module40"
Option Explicit

'==================================================================
' 【一覧表PDF用】Book_Shipment_List.xlsx → Book_Shipment_List.pdf 出力
'   ・Salacia_Shipment_List.pdf の配色（緑系ヘッダー＋緑罫線）を参考にした体裁でPDF化。
'   ・項目名（本の色 / 名前 / 支払方法）は変更しない。
'   ・本の色 列のセルを、その色名（青/茶/紫/白/赤 等）に応じて淡色で塗り分け。
'   ・元データ(.xlsx)は変更を保存せず、PDFのみ出力する。
'==================================================================
Public Sub Book一覧表PDF出力()
    Const DL_DIR = "C:\Users\lenovo\Desktop\ダウンロード\"
    Dim bookPath As String, pdfPath As String
    bookPath = DL_DIR & "Book_Shipment_List.xlsx"
    pdfPath = DL_DIR & "Book_Shipment_List.pdf"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(bookPath) Then
        Debug.Print "Book_Shipment_List.xlsx が見つかりません: " & bookPath
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(bookPath)
    Set ws = wb.Sheets(1)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    If lastRow < 1 Then lastRow = 1

    ' --- 本の色で並べ替え（色ごとにまとめる。同色内は名前順）---
    '     並び順は ColorRank で定義（数字を変えれば順番を変更可能）
    If lastRow >= 2 Then
        Dim r As Long
        ws.Cells(1, 4).Value = "_sortkey"            ' 一時的な並べ替えキー列(D)
        For r = 2 To lastRow
            ws.Cells(r, 4).Value = ColorRank(Trim(ws.Cells(r, "A").Value))
        Next r
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add2 key:=ws.Range("D2:D" & lastRow), Order:=xlAscending  ' 色ランク
            .SortFields.Add2 key:=ws.Range("B2:B" & lastRow), Order:=xlAscending  ' 同色内は名前順
            .SetRange ws.Range("A1:D" & lastRow)
            .Header = xlYes
            .Apply
        End With
        ws.Columns("D").Delete                       ' 並べ替え用の列を削除
    End If

    ' --- ヘッダー（A1:C1）緑背景・白文字・太字・中央 ---
    With ws.Range("A1:C1")
        .Interior.Color = RGB(112, 173, 71)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' --- データ部：背景白 → 本の色セルを色名で塗り分け、配置調整 ---
    Dim i As Long, v As String
    If lastRow >= 2 Then
        With ws.Range("A2:C" & lastRow)
            .Interior.Color = RGB(255, 255, 255)
            .VerticalAlignment = xlCenter
        End With
        ws.Range("A2:A" & lastRow).HorizontalAlignment = xlCenter   ' 本の色：中央
        ws.Range("C2:C" & lastRow).HorizontalAlignment = xlCenter   ' 支払方法：中央
        For i = 2 To lastRow
            v = Trim(ws.Cells(i, "A").Value)
            ws.Cells(i, "A").Interior.Color = ColorForBook(v)
        Next i
    End If

    ' --- 表全体に緑の罫線 ---
    With ws.Range("A1:C" & lastRow).Borders
        .LineStyle = xlContinuous
        .Color = RGB(146, 208, 80)
        .Weight = xlThin
    End With

    ' --- 列幅自動調整 ---
    ws.Columns("A:C").AutoFit

    ' --- 印刷設定（A4縦・横1ページ・各ページにヘッダー・余白1cm）---
    With ws.PageSetup
        .PrintArea = "A1:C" & lastRow
        .PrintTitleRows = "$1:$1"          ' 各ページ先頭に見出し行を繰り返す
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
    End With

    ' --- PDF出力 ---
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard, _
        OpenAfterPublish:=False

    ' 元データ(.xlsx)は変更を保存せず閉じる（書式はPDFのみに反映）
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    Debug.Print "PDF出力: " & pdfPath
    Debug.Print "===== 完了 ====="
End Sub


'----------------------------------------------
' 本の色 → セル塗り色（淡色・黒文字が読める明度）
'----------------------------------------------
Private Function ColorForBook(ByVal v As String) As Long
    Select Case v
        Case "青":                                                          ColorForBook = RGB(189, 215, 238)  ' 淡い青
        Case "茶", "茶のおまけ赤だけ", "茶のおまけ緑だけ", "茶のおまけ赤緑だけ": ColorForBook = RGB(223, 194, 165)  ' 淡い茶
        Case "紫":                                                          ColorForBook = RGB(213, 191, 219)  ' 淡い紫
        Case "白":                                                          ColorForBook = RGB(255, 255, 255)  ' 白
        Case "赤":                                                          ColorForBook = RGB(248, 203, 203)  ' 淡い赤
        Case Else:                                                          ColorForBook = RGB(255, 255, 255)  ' 未定義は白
    End Select
End Function


'----------------------------------------------
' 本の色 → 並べ替え用ランク（小さいほど上に並ぶ。数字を変えれば順番変更可）
'----------------------------------------------
Private Function ColorRank(ByVal v As String) As Long
    Select Case v
        Case "青":               ColorRank = 1
        Case "茶":               ColorRank = 2
        Case "茶のおまけ赤だけ":   ColorRank = 3
        Case "茶のおまけ緑だけ":   ColorRank = 4
        Case "茶のおまけ赤緑だけ": ColorRank = 5
        Case "紫":               ColorRank = 6
        Case "白":               ColorRank = 7
        Case "赤":               ColorRank = 8
        Case Else:               ColorRank = 99   ' 未定義は最後にまとめる
    End Select
End Function


