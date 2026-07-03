Attribute VB_Name = "Module1"
Sub ƒ†[ƒU“o˜^‚ÆŠm”F()
Attribute ƒ†[ƒU“o˜^‚ÆŠm”F.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' ƒ†[ƒU“o˜^‚ÆŠm”F Macro
'

'
    Rows("1:1").Select
    ActiveSheet.Select
    ActiveSheet.Name = "1"
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AZ$433").AutoFilter field:=2, Criteria1:="Š®—¹"
    ActiveWorkbook.Worksheets("1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("1").AutoFilter.Sort.SortFields.Add2 key:=Range( _
        "E1:E433"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("E:E").EntireColumn.AutoFit
    ActiveSheet.Range("$A$1:$AZ$433").AutoFilter field:=9, Criteria1:=Array( _
        "y†‚Ì‘ĞzCARBOFFf", "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“ü1", "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“ü1g", _
        "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“ü1y", "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“üg", "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“ül", _
         "y“dq‘ĞzCARBOFF 7“úŒÀ’è‰¿Šil", "y“dq‘ĞzCARBOFF 1“úŒÀ’è‰¿Šil", "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“ür", "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“üy", "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“üll", "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“ülll", "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“üt", "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“ü1t", "y“dq‘ĞzCARBOFF 7“úŒÀ’è‰¿Šig", "y“dq‘ĞzCARBOFF 7“úŒÀ’è‰¿Šiy", "y“dq‘ĞzCARBOFF 7“úŒÀ’è‰¿Šiy|", "y“dq‘ĞzCARBOFF 7“úŒÀ’è‰¿Šif", "y“dq‘ĞzCARBOFF 7“úŒÀ’è‰¿Šif|", "y“dq‘Ğzˆóü”Å‚Ù‚µ‚¢•û‚à‚±‚¿‚ç‚ğw“ü", "y“dq‘ĞzCARBOFF 1“úŒÀ’è‰¿Šig", "y“dq‘ĞzCARBOFF 1“úŒÀ’è‰¿Šiy", "y“dq‘ĞzCARBOFF 1“úŒÀ’è‰¿Šif", "y“dq‘ĞzCARBOFF 3“úŒÀ’è‰¿Ši", "y“dq‘ĞzCARBOFF 7“úŒÀ’è‰¿Šim", "y“dq‘ĞzCARBOFF 7“úŒÀ’è‰¿Šip", "y“dq‘ĞzCARBOFF 1“úŒÀ’è‰¿Šim", "y“dq‘ĞzCARBOFF 1“úŒÀ’è‰¿Šip", "y“dq‘ĞzCARBOFF 3“úŒÀ’è‰¿Ši", "y“dq‘ĞzCARBOFF 3“úŒÀ’è‰¿Šill", "y“dq‘ĞzCARBOFF 3“úŒÀ’è‰¿Šilll"), Operator:=xlFilterValues
        ActiveWindow.SmallScroll Down:=-30
    Columns("F:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:W").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:N").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:T").Select
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Columns("F:F").EntireColumn.AutoFit
    Range("A1").Select
End Sub
