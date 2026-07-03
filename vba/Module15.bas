Attribute VB_Name = "Module15"
Sub 既定のプリンタで2部印刷()
Attribute 既定のプリンタで2部印刷.VB_ProcData.VB_Invoke_Func = "P\n14"
    ' A列文字サイズ48()
    Columns("A:A").Font.Size = 48
    
    ' アクティブなシートを既定プリンターで2部印刷
    ActiveSheet.PrintOut Copies:=2, Collate:=True
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub
