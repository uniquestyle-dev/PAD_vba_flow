Attribute VB_Name = "Module13"
Option Explicit

Sub Shopify列削除_i()
Attribute Shopify列削除_i.VB_ProcData.VB_Invoke_Func = "i\n14"
    Dim ws As Worksheet
    ' 実行中のブックでアクティブなシートを対象とする
    Set ws = ActiveSheet

    ' 1. B列すべてのセルの書式設定を「数値（小数点以下2桁）」に設定
    ws.Columns("B").NumberFormat = "0"

    ' 2. G列を削除
    ws.Columns("G").Delete

    ' 処理完了メッセージ（不要ならコメントアウト）
    ' MsgBox "B列を数値書式に設定し、G列を削除しました。", vbInformation
    
    ' 今アクティブなブックを保存して閉じる
    With ActiveWorkbook
        .Close SaveChanges:=True
    End With
End Sub

