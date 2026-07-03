Attribute VB_Name = "Module38"
Option Explicit

'==================================================================
' 【一覧表PDF用】Shopify注文CSV → Book_Shipment_List.xlsx へ追記（単独実行マクロ）
'   ・既存の「Shopify変換」(代引き/ラベル/999.csv出力)とは別名・別モジュール想定。
'     既存処理には一切影響しません。必ず新しい標準モジュールに配置すること。
'   ・shopify-list.csv の出力は廃止し、既存の Book_Shipment_List.xlsx に行を追記する。
'     列構成は 本の色 / 名前 / 支払方法（Salaciaマクロが作るヘッダと同一）。
'   ・本の色   : Lineitem name を対応表(変換対応表Shopify)で変換
'   ・名前     : Shipping Name（複数商品注文は注文番号でフィルダウン）
'   ・支払方法 : 「NP後払い…」のみ「コンビニ支払い」に変換、他は空欄。
'                ただし Billing Street / Billing Address1 が Shipping Street / Shipping Address1 と
'                異なる注文（買い主と届け先が別住所）は、支払方法を空欄にする。
'   ・前提: Salaciaマクロが先に Book_Shipment_List.xlsx を作成・リセット済みであること。
'==================================================================
Public Sub Shopify変換一覧表用()
    Dim filePath As String
    Dim wb As Workbook, ws As Worksheet
    Dim qt As QueryTable
    Dim lastRow As Long, lastCol As Long
    Dim i As Long
    Dim idxLineName As Long, idxShippingName As Long
    Dim idxOrderNum As Long, idxPaymentMethod As Long
    Dim idxBillingStreet As Long, idxBillingAddr1 As Long
    Dim idxShippingStreet As Long, idxShippingAddr1 As Long
    Dim fso As Object
    Dim outputFolder As String
    Dim orderNo As String, nm As String, pay As String, li As String
    Dim nameDict As Object, payDict As Object, addrDiffDict As Object

    '─── 1. 最新の orders_export*.csv を自動取得（読み込み部は変更なし）───
    Dim folder As Object, file As Object
    Dim downloadPath As String
    Dim latestDate As Date

    downloadPath = "C:\Users\lenovo\Desktop\ダウンロード\"
    filePath = ""
    latestDate = 0

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(downloadPath)
    For Each file In folder.Files
        If LCase(file.Name) Like "orders_export*.csv" Then
            If file.DateCreated > latestDate Then
                latestDate = file.DateCreated
                filePath = file.Path
            End If
        End If
    Next file

    If filePath = "" Then
        Debug.Print "orders_export.csv が見つかりません"
        Exit Sub
    End If
    Debug.Print "ファイル: " & filePath

    '─── 2. CSV読み込み（UTF-8 / 全列テキスト：読み込み部は変更なし）───
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Sheets(1)

    Set qt = ws.QueryTables.Add("TEXT;" & filePath, ws.Range("A1"))
    With qt
        .TextFilePlatform = 65001
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True

        Dim tArr(0 To 255) As Variant
        For i = 0 To 255: tArr(i) = xlTextFormat: Next i
        .TextFileColumnDataTypes = tArr

        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
    End With

    '─── 3. 必要列インデックス取得 ───
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    idxLineName = 0: idxShippingName = 0: idxOrderNum = 0: idxPaymentMethod = 0
    idxBillingStreet = 0: idxBillingAddr1 = 0
    idxShippingStreet = 0: idxShippingAddr1 = 0
    For i = 1 To lastCol
        Select Case ws.Cells(1, i).Value
            Case "Lineitem name":      idxLineName = i
            Case "Shipping Name":      idxShippingName = i
            Case "Name":               idxOrderNum = i
            Case "Payment Method":     idxPaymentMethod = i
            Case "Billing Street":     idxBillingStreet = i
            Case "Billing Address1":   idxBillingAddr1 = i
            Case "Shipping Street":    idxShippingStreet = i
            Case "Shipping Address1":  idxShippingAddr1 = i
        End Select
    Next i

    '─── 4. 注文番号ごとに 宛名・支払方法・住所差分 を確定（複数商品行への補完用）───
    '     Shopifyは注文の先頭行のみ注文単位項目が入るため、先頭行の値を採用。
    '     addrDiffDict: Billing と Shipping の住所が違う注文を記録（買い主と届け先が別住所）。
    Set nameDict = CreateObject("Scripting.Dictionary")
    Set payDict = CreateObject("Scripting.Dictionary")
    Set addrDiffDict = CreateObject("Scripting.Dictionary")
    Dim bStreet As String, sStreet As String, bAddr1 As String, sAddr1 As String
    Dim diff As Boolean
    For i = 2 To lastRow
        orderNo = Trim(ws.Cells(i, idxOrderNum).Value)
        nm = Trim(ws.Cells(i, idxShippingName).Value)
        If orderNo <> "" And nm <> "" Then
            If Not nameDict.Exists(orderNo) Then
                nameDict.Add orderNo, nm
                payDict.Add orderNo, Trim(ws.Cells(i, idxPaymentMethod).Value)

                ' Billing Street / Address1 と Shipping Street / Address1 を比較
                bStreet = "": sStreet = "": bAddr1 = "": sAddr1 = ""
                If idxBillingStreet > 0 Then bStreet = Trim(ws.Cells(i, idxBillingStreet).Value)
                If idxShippingStreet > 0 Then sStreet = Trim(ws.Cells(i, idxShippingStreet).Value)
                If idxBillingAddr1 > 0 Then bAddr1 = Trim(ws.Cells(i, idxBillingAddr1).Value)
                If idxShippingAddr1 > 0 Then sAddr1 = Trim(ws.Cells(i, idxShippingAddr1).Value)
                diff = (bStreet <> sStreet) Or (bAddr1 <> sAddr1)
                addrDiffDict.Add orderNo, diff
            End If
        End If
    Next i

    '─── 5. Book_Shipment_List.xlsx に追記（列構成: 本の色,名前,支払方法）───
    outputFolder = "C:\Users\lenovo\Desktop\ダウンロード\"

    Dim bookPath As String
    bookPath = outputFolder & "Book_Shipment_List.xlsx"

    Dim wbBook As Workbook, wsBook As Worksheet
    If fso.FileExists(bookPath) Then
        Set wbBook = Workbooks.Open(bookPath)
        Set wsBook = wbBook.Sheets(1)
    Else
        ' 無ければヘッダ付きで新規作成（本来はSalaciaマクロが先に作成する想定）
        Set wbBook = Workbooks.Add(xlWBATWorksheet)
        Set wsBook = wbBook.Sheets(1)
        wsBook.Columns("A:C").NumberFormat = "@"
        wsBook.Range("A1").Value = "本の色"
        wsBook.Range("B1").Value = "名前"
        wsBook.Range("C1").Value = "支払方法"
    End If

    ' 追記開始行（A列の最終行の次）
    Dim bookRow As Long
    bookRow = wsBook.Cells(wsBook.Rows.Count, "A").End(xlUp).row + 1
    If bookRow < 2 Then bookRow = 2

    Dim addedCount As Long
    addedCount = 0
    For i = 2 To lastRow
        li = Trim(ws.Cells(i, idxLineName).Value)
        If li <> "" Then                         ' 商品行のみ
            orderNo = Trim(ws.Cells(i, idxOrderNum).Value)

            nm = ""
            If nameDict.Exists(orderNo) Then nm = nameDict(orderNo)

            pay = ""
            If payDict.Exists(orderNo) Then pay = ConvertPayment(CStr(payDict(orderNo)))
            ' Billing と Shipping の住所が異なる注文は支払方法を空欄にする
            If addrDiffDict.Exists(orderNo) Then
                If addrDiffDict(orderNo) = True Then pay = ""
            End If

            wsBook.Cells(bookRow, 1).Value = ConvertItemNameShopify(li)
            wsBook.Cells(bookRow, 2).Value = nm
            wsBook.Cells(bookRow, 3).Value = pay
            bookRow = bookRow + 1
            addedCount = addedCount + 1
        End If
    Next i

    Application.DisplayAlerts = False
    wbBook.SaveAs fileName:=bookPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbBook.Close SaveChanges:=False

    ws.AutoFilterMode = False
    wb.Close SaveChanges:=False

    Debug.Print "Book_Shipment_List.xlsx に追記: " & addedCount & "件"
    Debug.Print "===== 完了 ====="
End Sub


'----------------------------------------------
' アイテム名変換（変換対応表Shopify.xlsx の内容をハードコード）
'   変換前(Lineitem name) → 変換後
'   ※未定義は変換せずそのまま出力（取りこぼし防止）
'----------------------------------------------
Private Function ConvertItemNameShopify(ByVal itemName As String) As String
    Select Case itemName
        Case "【印刷版】「油が太る」は本当か？":                ConvertItemNameShopify = "白"
        Case "【印刷版】CARBOFF-糖質の吸収をおさえる方法":      ConvertItemNameShopify = "青"
        Case "【紙書籍】Medico-痩せ薬で○kg痩せた話":           ConvertItemNameShopify = "紫"
        Case "【紙書籍】Salasiru-効率がいいサラシア摂取法":      ConvertItemNameShopify = "茶"
        Case Else:                                            ConvertItemNameShopify = itemName
    End Select
End Function


'----------------------------------------------
' 支払方法変換
'   「NP後払い…」だけ「コンビニ支払い」に変換、それ以外は空欄
'   ※表記ゆれ対策として "NP後払い" で前方一致判定
'----------------------------------------------
Private Function ConvertPayment(ByVal p As String) As String
    If Left$(Trim$(p), 5) = "NP後払い" Then
        ConvertPayment = "コンビニ支払い"
    Else
        ConvertPayment = ""
    End If
End Function

