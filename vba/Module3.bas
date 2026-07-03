Attribute VB_Name = "Module3"
Option Explicit

Public Sub 友だちリスト列削除_k()
    Dim desktopPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim projCol As Long
    Dim i As Long
    Dim deleteStart As Long
    Dim deleteEnd As Long

    ' 画面更新・警告メッセージをオフ
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' デスクトップのパスを取得
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    ' Subscriber_ で始まる CSV ファイルを検索
    fileName = Dir(desktopPath & "Subscriber_*.csv")

    Do While fileName <> ""
        On Error GoTo NextFile
        ' ファイルを開く（Shift-JIS は Local:=True）
        Set wb = Workbooks.Open(fileName:=desktopPath & fileName, Local:=True)
        Set ws = wb.Worksheets(1)

        ' ヘッダー行の最終列を取得
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        ' 「プロジェクト登録日時」列を探す
        projCol = 0
        For i = 1 To lastCol
            If ws.Cells(1, i).Value = "プロジェクト登録日時" Then
                projCol = i
                Exit For
            End If
        Next i

        If projCol <> 0 Then
            ' E列(5) から 「プロジェクト登録日時」列の左隣までを削除
            deleteStart = 5
            deleteEnd = projCol - 1
            If deleteEnd >= deleteStart Then
                ws.Range(ws.Columns(deleteStart), ws.Columns(deleteEnd)).Delete Shift:=xlToLeft
            End If

            ' 列削除後に改めて見出しを探し、「プロジェクト登録日時」を「登録日時」に置換
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            For i = 1 To lastCol
                If ws.Cells(1, i).Value = "プロジェクト登録日時" Then
                    ws.Cells(1, i).Value = "登録日時"
                    Exit For
                End If
            Next i
        End If

        ' 上書き保存して閉じる
        wb.SaveAs fileName:=desktopPath & fileName, FileFormat:=xlCSV, Local:=True
        wb.Close SaveChanges:=False

NextFile:
        Err.Clear
        ' 次のファイルへ
        fileName = Dir()
    Loop

    ' 画面更新・警告メッセージを戻す
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Subscriber_*.csv の一括処理が完了しました。", vbInformation
    
    ' Excel を閉じる
    Application.Quit
End Sub

