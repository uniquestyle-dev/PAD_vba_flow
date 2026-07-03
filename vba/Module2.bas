Attribute VB_Name = "Module2"
Sub CARBOFFのFB_件数集計_r()
Attribute CARBOFFのFB_件数集計_r.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' CARBOFFのFB_件数集計_r Macro
'
' Keyboard Shortcut: Ctrl+r
'
    ActiveSheet.Select
    ActiveSheet.Name = "1"
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("1").AutoFilter.Sort.SortFields.Add2 key:=Range( _
        "D1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("D:D").EntireColumn.AutoFit
    ' ↓ フィルター条件を「【電子書籍】CARBOFF 7日限定価格」を含む に変更
    ActiveSheet.Range("$A$1:$AZ$135").AutoFilter _
        field:=9, _
        Criteria1:="=*【電子書籍】CARBOFF*", _
        Operator:=xlAnd
End Sub


