Attribute VB_Name = "Module18"
Option Explicit
Sub ユーザ登録と確認(Optional filePathIn As Variant, Optional firstDate As Variant)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim lastRow As Long
    Dim emailCol As Long
    Dim orderDateCol As Long
    Dim filterDate As Date
    Dim savePath As String
    Dim i As Long
    Dim emails As Collection
    Dim col As Long
    Dim cellDate As Date
    Dim cellValue As String
    Dim tmpDate As String
    Dim headerVal As String
    Dim emailVal As String
    Dim fileNum As Integer
    
    On Error GoTo ErrorHandler
    
    Debug.Print "===== マクロ開始 ====="
    
    ' ===== STEP 1: firstDate処理 =====
    filterDate = DateSerial(1900, 1, 1)
    
    If Not IsMissing(firstDate) Then
        If firstDate <> "" Then
            tmpDate = CStr(firstDate)
            Debug.Print "firstDate元値: [" & tmpDate & "]"
            If InStr(tmpDate, "(") > 0 Then
                tmpDate = Trim(Left(tmpDate, InStr(tmpDate, "(") - 1))
            End If
            If InStr(tmpDate, "+") > 0 Then
                tmpDate = Trim(Left(tmpDate, InStr(tmpDate, "+") - 1))
            End If
            On Error Resume Next
            filterDate = CDate(tmpDate)
            If Err.Number <> 0 Then
                Err.Clear
                filterDate = DateSerial(1900, 1, 1)
            End If
            On Error GoTo ErrorHandler
        End If
    End If
    Debug.Print "filterDate: " & filterDate
    
    ' ===== STEP 2: 顧客リストファイルを開く =====
    If IsMissing(filePathIn) Or filePathIn = "" Then
        filePathIn = Application.GetOpenFilename( _
            FileFilter:="Excel/CSV ファイル (*.xlsx;*.xls;*.csv),*.xlsx;*.xls;*.csv", _
            Title:="顧客リストファイルを選択してください" _
        )
        If filePathIn = False Then
            Exit Sub
        End If
    End If
    
    Debug.Print "ファイルパス: " & filePathIn
    
    If Dir(filePathIn) = "" Then
        Debug.Print "ファイルが存在しません"
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePathIn)
    Set ws = wb.Sheets(1)
    Debug.Print "ファイルを開きました: " & wb.Name
    
    ' 列を特定（ヘッダーから検索）
    orderDateCol = 0
    emailCol = 0
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        headerVal = CStr(ws.Cells(1, col).Value)
        Debug.Print "列" & col & ": [" & headerVal & "]"
        If InStr(headerVal, "オーダー日時") > 0 Or InStr(headerVal, "注文日") > 0 Then
            orderDateCol = col
        End If
        If InStr(headerVal, "メールアドレス") > 0 Or headerVal = "Email" Or headerVal = "email" Then
            emailCol = col
        End If
    Next col
    
    Debug.Print "オーダー日時列: " & orderDateCol & ", メールアドレス列: " & emailCol
    
    If orderDateCol = 0 Or emailCol = 0 Then
        Debug.Print "必要な列が見つかりません"
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    Debug.Print "最終行: " & lastRow
    
    ' ===== STEP 3: メールアドレス抽出 =====
    Set emails = New Collection
    
    For i = 2 To lastRow
        cellValue = CStr(ws.Cells(i, orderDateCol).Value)
        
        If cellValue <> "" Then
            On Error Resume Next
            cellDate = CDate(cellValue)
            If Err.Number <> 0 Then
                Err.Clear
                cellDate = DateSerial(1900, 1, 1)
            End If
            On Error GoTo ErrorHandler
            
            If cellDate >= filterDate Then
                emailVal = CStr(ws.Cells(i, emailCol).Value)
                If emailVal <> "" Then
                    On Error Resume Next
                    emails.Add emailVal, emailVal
                    On Error GoTo ErrorHandler
                    Debug.Print "  追加: " & emailVal
                End If
            End If
        End If
    Next i
    
    Debug.Print "抽出メール件数: " & emails.Count
    
    wb.Close SaveChanges:=False
    
    ' ===== STEP 4: CSV保存（空行なし） =====
    savePath = "C:\Users\lenovo\Desktop\ダウンロード\emails_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
    Debug.Print "保存先: " & savePath
    
    fileNum = FreeFile
    Open savePath For Output As #fileNum
    
    If emails.Count = 0 Then
        ' 該当なしの場合は半角スペース1つ
        Print #fileNum, " ";
        Debug.Print "該当なし - 半角スペースのみ出力"
    Else
        ' メールアドレスを出力
        For i = 1 To emails.Count
            If i < emails.Count Then
                Print #fileNum, emails(i)
            Else
                Print #fileNum, emails(i);
            End If
        Next i
    End If
    
    Close #fileNum
    
    Debug.Print "===== マクロ正常終了 ====="
    Debug.Print "保存完了: " & emails.Count & "件 → " & savePath
    Exit Sub
    
ErrorHandler:
    Debug.Print "エラー: " & Err.Number & " - " & Err.Description
    Application.DisplayAlerts = True
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
End Sub

