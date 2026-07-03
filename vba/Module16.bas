Attribute VB_Name = "Module16"
Sub CSV保存_CS_s()
Attribute CSV保存_CS_s.VB_ProcData.VB_Invoke_Func = "S\n14"
    Dim wb As Workbook
    Dim folderPath As String
    Dim baseName As String
    Dim fullPath As String
    Dim i As Integer

    Set wb = ActiveWorkbook

    ' 保存済みファイルであれば通常保存
    If wb.Path <> "" Then
        wb.Save
        MsgBox "既存ファイルを上書き保存しました。", vbInformation
        Exit Sub
    End If

    ' 保存先フォルダ（例：デスクトップ）
    folderPath = Environ("USERPROFILE") & "\Desktop\"
    baseName = "Book"
    i = 1

    ' Book1.csv, Book2.csv ... のように既存ファイルを避けて保存名を決定
    Do
        fullPath = folderPath & baseName & i & ".csv"
        If Dir(fullPath) = "" Then Exit Do
        i = i + 1
    Loop

    ' CSV形式で保存
    wb.SaveAs fileName:=fullPath, FileFormat:=xlCSVUTF8, CreateBackup:=False

    MsgBox "CSVとして保存しました：" & vbCrLf & fullPath, vbInformation
    
    wb.Close SaveChanges:=False
    
End Sub

