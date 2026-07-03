Attribute VB_Name = "Module14"
Sub サラ茶_列削除_Shift_i()
Attribute サラ茶_列削除_Shift_i.VB_ProcData.VB_Invoke_Func = "I\n14"
    ' B列からI列を削除
    Columns("B:I").Delete Shift:=xlToLeft
    ' アクティブシートの B 列全体を数値（小数点以下 0桁）に書式設定
    With ActiveSheet.Columns("B:B")
        .NumberFormat = "0"
    End With
     ' 今アクティブなブックを保存して閉じる
    With ActiveWorkbook
        .Close SaveChanges:=True
    End With
End Sub
