Attribute VB_Name = "Module20"
Option Explicit

Public Sub STORES変換()
    Dim inputFile As String
    Dim outputFolder As String
    Dim wbInput As Workbook
    Dim ws As Worksheet
    Dim colItem As Long, colZip As Long
    Dim colSei As Long, colMei As Long
    Dim colPref As Long, colAddr As Long
    Dim colOrderNumber As Long, colOrderDate As Long
    Dim lastRow As Long
    Dim dict As Object
    Dim cell As Range
    Dim item As Variant
    Dim fso As Object, ts As Object, tsSummary As Object
    Dim fields As Variant
    Dim lineParts As Variant
    Dim i As Long
    Dim headerVal As String
    Dim col As Long
    Dim category As String
    Dim itemName As String

    ' 入力CSV ダウンロードフォルダから最新の20*.csvを取得
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

    ' CSVをcp932で読み込み
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

    ' ヘッダ行から列番号を取得
    colItem = 0: colZip = 0: colSei = 0: colMei = 0
    colPref = 0: colAddr = 0: colOrderNumber = 0: colOrderDate = 0
    
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        headerVal = ws.Cells(1, col).Value
        Select Case headerVal
            Case "アイテム名":                      colItem = col
            Case "郵便番号(購入者)", "郵便番号(配送先)":  colZip = col
            Case "氏(購入者)", "氏(配送先)":           colSei = col
            Case "名(購入者)", "名(配送先)":           colMei = col
            Case "都道府県(購入者)", "都道府県(配送先)":  colPref = col
            Case "住所(購入者)", "住所(配送先)":        colAddr = col
            Case "オーダー番号":                      colOrderNumber = col
            Case "オーダー日時", "オーダー日":          colOrderDate = col
        End Select
    Next col

    ' データ最終行
    lastRow = ws.Cells(ws.Rows.Count, colItem).End(xlUp).row

    If lastRow <= 1 Then
        wbInput.Close SaveChanges:=False
        Debug.Print "STORES 受信データ0件 - スキップ"
        Exit Sub
    End If

    ' 1.csv（サマリ）を出力
    Set tsSummary = fso.CreateTextFile(outputFolder & "\1.csv", True, False)
    tsSummary.WriteLine Join(Array( _
        "オーダー番号", "オーダー日", _
        "氏(配送先)", "名(配送先)", _
        "配送方法", "決済予定日", "お問い合わせ番号", "備考", _
        "注文個数" _
    ), ",")
    For i = 2 To lastRow
        tsSummary.WriteLine _
            ws.Cells(i, colOrderNumber).Value & "," & _
            ws.Cells(i, colOrderDate).Value & "," & _
            ws.Cells(i, colSei).Value & "," & _
            ws.Cells(i, colMei).Value & "," & _
            "日本郵便" & "," & _
            "" & "," & _
            "" & "," & _
            "" & "," & _
            "1"
    Next i
    tsSummary.Close

    ' カテゴリごとに集約
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        itemName = Trim(ws.Cells(i, colItem).Value)
        category = GetStoresCategory(itemName)
        If Not dict.Exists(category) Then dict.Add category, Nothing
    Next i

    ' 出力用ヘッダ配列
    fields = Array( _
        "お届け先郵便番号", "お届け先氏名", "お届け先敬称", _
        "お届け先住所1行目", "お届け先住所2行目", _
        "お届け先住所3行目", "お届け先住所4行目", "内容品" _
    )

    ' カテゴリごとにファイル出力
    For Each item In dict.Keys
        Set ts = fso.CreateTextFile(outputFolder & "\" & item & ".csv", True, False)
        ts.WriteLine Join(fields, vbTab)
        
        For i = 2 To lastRow
            itemName = Trim(ws.Cells(i, colItem).Value)
            category = GetStoresCategory(itemName)
            
            If category = item Then
                lineParts = Array( _
                    Format(ws.Cells(i, colZip).Value, "0000000"), _
                    SanitizeClickPost(ws.Cells(i, colSei).Value & ws.Cells(i, colMei).Value), _
                    "様", _
                    SanitizeClickPost(ws.Cells(i, colPref).Value), _
                    SanitizeClickPost(ws.Cells(i, colAddr).Value), _
                    "", "", "書籍" _
                )
                ts.WriteLine Join(lineParts, vbTab)
            End If
        Next i
        
        ts.Close
        Debug.Print "出力: " & item & ".csv"
    Next item

    wbInput.Close SaveChanges:=False
   
    ' 1.csv を開く
    Workbooks.Open fileName:=outputFolder & "\1.csv"
    
    Debug.Print "===== 完了 ====="
End Sub


'----------------------------------------------
' アイテム名 → カテゴリのマッピング
'----------------------------------------------
Private Function GetStoresCategory(ByVal itemName As String) As String
    Select Case itemName
        Case "電子書籍印刷サービス_c"
            GetStoresCategory = "carboff_stores"
        Case "電子書籍印刷サービス_s", "電子書籍印刷サービス_sr"
            GetStoresCategory = "salasiru_stores"
        Case "電子書籍印刷サービス_m", "電子書籍印刷サービス_mr"
            GetStoresCategory = "Medico_stores"
        Case "電子書籍印刷サービス_i"
            GetStoresCategory = "ice_stores"
        Case "電子書籍印刷サービス_sw"
            GetStoresCategory = "sweets_stores"
        Case "電子書籍印刷サービス_set"
            GetStoresCategory = "set_stores"
        Case Else
            GetStoresCategory = "other_stores_" & itemName
    End Select
End Function


'----------------------------------------------
' ClickPost取込エラー回避: , と " を除去
'----------------------------------------------
Private Function SanitizeClickPost(ByVal txt As String) As String
    SanitizeClickPost = Replace(Replace(txt, ",", ""), """", "")
End Function