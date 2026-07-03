Attribute VB_Name = "Module21"
'=============================================================
' 対応状況=5 & FirstDate以降のメールアドレスをCSV出力
' VBSから呼び出し: objExcel.Run "PERSONAL.XLSB!Module18.FilterOrderEmails", latestFile, firstDate
'=============================================================
Sub リピスト書籍購入ユーザ変換(ByVal strInputCSV As String, _
                      ByVal strFirstDate As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dtFirstDate As Date
    Dim lastRow As Long
    Dim i As Long
    Dim emailDict As Object
    Dim statusVal As String
    Dim dateVal As Date
    Dim emailVal As String
    Dim savePath As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    dtFirstDate = CDate(strFirstDate)

    ' --- 保存パス自動生成 ---
    savePath = "C:\Users\lenovo\Desktop\ダウンロード\emails_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"

    ' --- 入力CSV読み込み ---
    Set wb = Workbooks.Open(fileName:=strInputCSV, ReadOnly:=True)
    Set ws = wb.Sheets(1)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' --- ヘッダーから列位置を特定 ---
    Dim colStatus As Long, colDate As Long, colEmail As Long
    Dim col As Long
    colStatus = 0: colDate = 0: colEmail = 0

    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Select Case Trim(ws.Cells(1, col).Value)
            Case "対応状況":     colStatus = col
            Case "注文日時":     colDate = col
            Case "メールアドレス": colEmail = col
        End Select
    Next col

    If colStatus = 0 Or colDate = 0 Or colEmail = 0 Then
        MsgBox "必要な列が見つかりません", vbCritical
        wb.Close SaveChanges:=False
        GoTo Cleanup
    End If

    ' --- データ抽出（重複排除） ---
    Set emailDict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        statusVal = Trim(CStr(ws.Cells(i, colStatus).Value))

        ' 条件1: 対応状況 = 5
        If statusVal = "5" Then
            ' 条件2: 注文日時 >= FirstDate
            If IsDate(ws.Cells(i, colDate).Value) Then
                dateVal = CDate(ws.Cells(i, colDate).Value)
                If dateVal >= dtFirstDate Then
                    emailVal = Trim(CStr(ws.Cells(i, colEmail).Value))
                    If emailVal <> "" Then
                        If Not emailDict.Exists(emailVal) Then
                            emailDict.Add emailVal, True
                        End If
                    End If
                End If
            End If
        End If
    Next i

    wb.Close SaveChanges:=False

' --- CSV出力 ---
    Dim ff As Integer
    ff = FreeFile
    Open savePath For Output As #ff
        Dim key As Variant
        For Each key In emailDict.Keys
            Print #ff, key
        Next key
    Close #ff

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


