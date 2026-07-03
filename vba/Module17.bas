Attribute VB_Name = "Module17"
Option Explicit
Public Sub Shopify登録確認_Shift_w(Optional filePathIn As Variant, Optional firstDate As Variant)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim paidAtCol As Variant
    Dim emailCol As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim filterDate As Date
    Dim i As Long
    Dim outputWb As Workbook
    Dim outputWs As Worksheet
    Dim outputRow As Long
    Dim cellDate As Date
    Dim savePath As String
    Dim cellValue As String
    Dim fileNum As Integer
    Dim tmpDate As String
    
    On Error GoTo ErrorHandler
    
    Debug.Print "========== マクロ開始 =========="
    Debug.Print "実行時刻: " & Now
    
    ' ===== STEP 1: ファイルパス取得 =====
    Debug.Print "[STEP 1] ファイルパス処理開始"
    If IsMissing(filePathIn) Then
        Debug.Print "  → filePathIn: Missing"
    Else
        Debug.Print "  → filePathIn: " & filePathIn
    End If
    
    If IsMissing(firstDate) Then
        Debug.Print "  → firstDate: Missing"
        filterDate = DateSerial(1900, 1, 1)
    Else
        Debug.Print "  → firstDate: " & firstDate
        tmpDate = CStr(firstDate)
        If InStr(tmpDate, "+") > 0 Then
            tmpDate = Trim(Left(tmpDate, InStr(tmpDate, "+") - 1))
        End If
        filterDate = CDate(tmpDate)
    End If
    Debug.Print "  → filterDate: " & filterDate
    
    If IsMissing(filePathIn) Or filePathIn = "" Then
        Debug.Print "  → ファイル選択ダイアログを表示"
        filePathIn = Application.GetOpenFilename( _
            FileFilter:="CSV ファイル (*.csv),*.csv", _
            Title:="変換する UTF-8 CSV を選択してください" _
        )
        If filePathIn = False Then
            Debug.Print "  → キャンセルされました"
            Exit Sub
        End If
    End If
    Debug.Print "[STEP 1 完了] ファイルパス: " & filePathIn
    
    If Dir(filePathIn) = "" Then
        Debug.Print "  → エラー: ファイルが存在しません: " & filePathIn
        Exit Sub
    End If
    Debug.Print "  → ファイル存在確認OK"
    
    ' ===== STEP 2: 新規ワークブック作成 =====
    Debug.Print "[STEP 2] 新規ワークブック作成"
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Sheets(1)
    
    ' ===== STEP 3: QueryTable作成 =====
    Debug.Print "[STEP 3] QueryTable作成"
    Set qt = ws.QueryTables.Add( _
        Connection:="TEXT;" & filePathIn, _
        Destination:=ws.Range("A1") _
    )
    
    ' ===== STEP 4: QueryTable設定 =====
    Debug.Print "[STEP 4] QueryTable設定"
    With qt
        .TextFileParseType = xlDelimited
        .TextFilePlatform = 65001
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileCommaDelimiter = True
        .Refresh BackgroundQuery:=False
    End With
    
    lastRow = ws.UsedRange.Rows.Count
    lastCol = ws.UsedRange.Columns.Count
    Debug.Print "[STEP 4 完了] データ範囲: " & lastRow & "行 x " & lastCol & "列"
    
    If lastRow <= 1 Then
        Debug.Print "  → 警告: データが空またはヘッダーのみ（ヘッダーのみのCSVを出力します）"
    End If
    
    ' ===== STEP 5: 列検索 =====
    Debug.Print "[STEP 5] 列検索"
    If lastRow > 1 Then
        On Error Resume Next
        paidAtCol = Application.Match("Paid at", ws.Rows(1), 0)
        emailCol = Application.Match("Email", ws.Rows(1), 0)
        On Error GoTo ErrorHandler
        
        If IsError(paidAtCol) Or IsError(emailCol) Then
            Debug.Print "  → エラー: 必要な列が見つかりません"
            wb.Close SaveChanges:=False
            Exit Sub
        End If
        
        Debug.Print "[STEP 5 完了] Paid at列: " & CLng(paidAtCol) & ", Email列: " & CLng(emailCol)
        
        ' ===== STEP 6: ソート =====
        Debug.Print "[STEP 6] ソート実行"
        ws.UsedRange.Sort _
            Key1:=ws.Cells(1, CLng(paidAtCol)), _
            Order1:=xlDescending, _
            Header:=xlYes
        Debug.Print "[STEP 6 完了] ソート成功"
    Else
        Debug.Print "[STEP 5-6 スキップ] データ行なし"
    End If
    
    ' ===== STEP 7: フィルタリング＆メール抽出 =====
    Debug.Print "[STEP 7] フィルタリング＆メール抽出"
    
    Set outputWb = Workbooks.Add(xlWBATWorksheet)
    Set outputWs = outputWb.Sheets(1)
    outputRow = 1
    
    outputWs.Cells(outputRow, 1).Value = "Email"
    outputRow = outputRow + 1
    
    For i = 2 To lastRow
        cellValue = ws.Cells(i, CLng(paidAtCol)).Value
        
        If InStr(cellValue, "+") > 0 Then
            cellValue = Trim(Left(cellValue, InStr(cellValue, "+") - 1))
        ElseIf InStrRev(cellValue, "-") > 10 Then
            cellValue = Trim(Left(cellValue, InStrRev(cellValue, "-") - 1))
        End If
        
        On Error Resume Next
        cellDate = CDate(cellValue)
        If Err.Number <> 0 Then
            Err.Clear
            cellDate = DateSerial(1900, 1, 1)
        End If
        On Error GoTo ErrorHandler
        
        If cellDate >= filterDate Then
            outputWs.Cells(outputRow, 1).Value = ws.Cells(i, CLng(emailCol)).Value
            outputRow = outputRow + 1
        End If
    Next i
    
    Debug.Print "  → 抽出件数: " & (outputRow - 2) & "件"
    
    ' ===== STEP 8: CSV保存（空行なし） =====
    Debug.Print "[STEP 8] CSV保存"
    savePath = Left(filePathIn, InStrRev(filePathIn, "\")) & "emails_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
    Debug.Print "  → 保存先: " & savePath
    
    fileNum = FreeFile
    Open savePath For Output As #fileNum
    
    For i = 1 To outputRow - 1
        If i < outputRow - 1 Then
            Print #fileNum, outputWs.Cells(i, 1).Value
        Else
            Print #fileNum, outputWs.Cells(i, 1).Value;
        End If
    Next i
    
    Close #fileNum
    
    Debug.Print "[STEP 8 完了] CSV保存成功"
    
    ' ===== STEP 9: 完了処理 =====
    Debug.Print "[STEP 9] 完了処理"
    outputWb.Close SaveChanges:=False
    wb.Close SaveChanges:=False
    
    On Error Resume Next
    qt.Delete
    On Error GoTo 0
    
    Debug.Print "========== マクロ正常終了 =========="
    Debug.Print "処理完了 - 抽出件数: " & (outputRow - 2) & "件"
    Exit Sub
    
ErrorHandler:
    Debug.Print "========== エラー発生 =========="
    Debug.Print "エラー番号: " & Err.Number
    Debug.Print "エラー内容: " & Err.Description
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
End Sub


