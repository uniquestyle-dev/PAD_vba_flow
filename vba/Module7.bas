Attribute VB_Name = "Module7"
Sub Shopify_追跡番号登録()
Attribute Shopify_追跡番号登録.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Shopify_追跡番号登録_i Macro
'
' Keyboard Shortcut: Ctrl+i
'
    ' シート名の変更
    Sheets("orders_export_1").Name = "1"
    
    ' R列とA列の範囲を指定し、R列を第一優先に昇順で並べ替える
    Dim lastRow As Long
    With Sheets("1")
        lastRow = .Cells(.Rows.Count, "R").End(xlUp).row
        .Range("A1:R" & lastRow).Sort Key1:=.Range("R1"), Order1:=xlAscending, Header:=xlYes
    End With
    
    ' 不要な列を削除
    Sheets("1").Columns("B:X").Delete Shift:=xlToLeft
    Sheets("1").Columns("C:AZ").Delete Shift:=xlToLeft

    ' 必要な列を追加
    Sheets("1").Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("1").Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("1").Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("1").Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("1").Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' ヘッダーを設定
    With Sheets("1")
        .Cells(1, 1).Value = "Order Number"
        .Cells(1, 2).Value = "Tracking Number"
        .Cells(1, 3).Value = "SKU"
        .Cells(1, 4).Value = "Quantity"
        .Cells(1, 5).Value = "Tracking Company"
        .Cells(1, 6).Value = "Tracking URL"
        .Cells(1, 7).Value = "Billing Name"
        
        ' 最終行を取得（A列の値が入っている行を対象）
        Dim lastRow2 As Long, i As Long
        lastRow2 = .Cells(.Rows.Count, "A").End(xlUp).row
        
        ' 2行目から最終行まで、E列に "Japan Post (JA)" を入力
        For i = 2 To lastRow2
            .Cells(i, 5).Value = "Japan Post (JA)"
        Next i
    End With
    
    ' 最初のセルを選択
    Sheets("1").Range("A1").Select
End Sub

