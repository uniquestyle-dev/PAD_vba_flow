Attribute VB_Name = "Module8"
Sub サラ茶_ステータス管理並べ替え_h()
Attribute サラ茶_ステータス管理並べ替え_h.VB_ProcData.VB_Invoke_Func = "h\n14"
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ' シート名を「1」に変更
    On Error Resume Next
    ws.Name = "1"
    On Error GoTo 0
    
    '―――――――――――――――――――――――――
    ' ① 「ご紹介」を削除し、余分なスペースもトリム（D列）
    '―――――――――――――――――――――――――
    Dim c As Range
    For Each c In ws.Range("D2:D100")
        If Not IsError(c.Value) And c.Value <> "" Then
            c.Value = Application.WorksheetFunction.Trim( _
                          Replace(Replace(c.Value, "ご紹介", ""), "追加", "") _
                      )
        End If
    Next c

    '―――――――――――――――――――――――――
    ' ② D列（商品名）の文字列を一括置換
    '―――――――――――――――――――――――――
    Dim rngD As Range
    Set rngD = ws.Range("D2:D100") ' 「商品名」がD列にある前提

    rngD.Replace _
        What:="20％引きの年間コース　【サラシア茶】", _
        Replacement:="15％引きの3ヶ月毎コース　【サラシア茶】", _
        LookAt:=xlPart

    rngD.Replace _
        What:="20％引きの年間コース　【サラシア粒】", _
        Replacement:="15％引きの3ヶ月毎コース　【サラシア粒】", _
        LookAt:=xlPart

    rngD.Replace _
        What:="0割引なしの単品購入　【サラシア茶】", _
        Replacement:="10％引きの1ヶ月毎コース　【サラシア茶】", _
        LookAt:=xlPart

    rngD.Replace _
        What:="0割引なしの単品購入　【サラシア粒】", _
        Replacement:="10％引きの1ヶ月毎コース　【サラシア粒】", _
        LookAt:=xlPart


    '―――――――――――――――――――――――――
    ' ③ F列（定期購入回数）の数値を正規化
    '―――――――――――――――――――――――――
    Dim cell As Range
    For Each cell In ws.Range("F2:F100") ' 「定期購入(自動受注)回数」がF列
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            If cell.Value >= 2 Then
                cell.Value = 2
            ElseIf cell.Value = 0 Then
                cell.Value = 1
            End If
        End If
    Next cell


    '―――――――――――――――――――――――――
    ' ④ ソート処理
    '    順序：D列(昇順) → F列(昇順) → I列(降順) → B列(昇順)
    '―――――――――――――――――――――――――
    With ws.Sort
        .SortFields.Clear

        ' 商品名（D列）を昇順
        .SortFields.Add2 _
            key:=ws.Range("D2:D100"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        ' 定期購入回数（F列）を昇順
        .SortFields.Add2 _
            key:=ws.Range("F2:F100"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        ' 支払い方法（I列）を降順
        .SortFields.Add2 _
            key:=ws.Range("I2:I100"), _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal

        ' 配送先姓（B列）を昇順
        .SortFields.Add2 _
            key:=ws.Range("B2:B100"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal

        ' ソート範囲をA1:I100に指定（ヘッダーあり）
        .SetRange ws.Range("A1:I100")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub


