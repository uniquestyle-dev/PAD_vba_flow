Attribute VB_Name = "Module39"
Option Explicit

'==================================================================
' 【一覧表PDF用】リピスト赤本(CARBOFF・コンビニ支払い) → Book_Shipment_List.xlsx へ追記
'   ・添付の「サラシア変換」マクロをベースに、インポート部はそのまま流用。
'     order_20*.csv（フォルダ内で最新）を Shift-JIS で読み込む。
'   ・前提: 入力ファイルは「赤本・コンビニ支払い」のみで構成されている。
'     （絞り込みは行わず、全データ行を 赤 / 氏名 / コンビニ支払い で追記する）
'   ・出力: Book_Shipment_List.xlsx
'     列構成 = 本の色 / 名前 / 支払方法（Salaciaマクロが作るヘッダと同一）
'       本の色   : 「赤」固定
'       名前     : B列（お届け先氏名）
'       支払方法 : 「コンビニ支払い」固定
'   ・既存処理には一切影響しません。必ず新しい標準モジュールに配置すること。
'   ・前提: Salaciaマクロが先に Book_Shipment_List.xlsx を作成・リセット済みであること。
'     ファイルが無い場合はヘッダ付きで新規作成する保険を入れている。
'==================================================================
Public Sub サラシア変換一覧表用()
    Dim filePath As String
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim lastRow As Long
    Dim i As Long

    '─── 1. order_20で始まる最新CSVファイルを自動取得（添付マクロと同一ロジック）───
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

    '─── 2. 新規ブック＋QueryTable で Shift JIS インポート（添付マクロと同一）───
    '     全列を文字列扱いにして「2-5-16」を日付化させない
    Dim wbInput As Workbook
    Set wbInput = Workbooks.Add(xlWBATWorksheet)
    Set ws = wbInput.Sheets(1)

    Set qt = ws.QueryTables.Add( _
        Connection:="TEXT;" & filePath, _
        Destination:=ws.Range("A1") _
    )
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFilePlatform = 932            ' Shift JIS
        ' 読み込む列数に合わせて要素数を調整（11列分）
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
        .Refresh BackgroundQuery:=False
    End With

    '─── 3. Book_Shipment_List.xlsx を開く（無ければヘッダ付きで新規作成）───
    Dim bookPath As String
    bookPath = folderPath & "Book_Shipment_List.xlsx"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim wbBook As Workbook, wsBook As Worksheet
    If fso.FileExists(bookPath) Then
        Set wbBook = Workbooks.Open(bookPath)
        Set wsBook = wbBook.Sheets(1)
    Else
        ' 保険: 無ければヘッダ付きで新規作成（本来はSalaciaマクロが先に作成する想定）
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

    '─── 4. 全データ行を 赤 / お届け先氏名(B列) / コンビニ支払い で追記 ───
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
    Dim addedCount As Long
    addedCount = 0
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, "B").Value)) <> "" Then
            wsBook.Cells(bookRow, 1).Value = "赤"
            wsBook.Cells(bookRow, 2).Value = ws.Cells(i, "B").Value
            wsBook.Cells(bookRow, 3).Value = "コンビニ支払い"
            bookRow = bookRow + 1
            addedCount = addedCount + 1
        End If
    Next i

    '─── 5. 保存して閉じる ───
    Application.DisplayAlerts = False
    wbBook.SaveAs fileName:=bookPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbBook.Close SaveChanges:=False

    wbInput.Close SaveChanges:=False

    Debug.Print "Book_Shipment_List.xlsx に追記(赤・コンビニ支払い): " & addedCount & "件"
    Debug.Print "===== 完了 ====="

    '─── 6. Excelを終了（元の サラシア変換 マクロと同じ挙動）───
    Application.Quit
End Sub


