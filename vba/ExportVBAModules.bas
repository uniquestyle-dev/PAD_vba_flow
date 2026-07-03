Attribute VB_Name = "ExportVBAModules"
Option Explicit

'====================================================================
' このマクロが「どこまでやるか」
'   ・PERSONAL.XLSB 内の全モジュールを指定フォルダへテキスト書き出し
'   ・種類別に拡張子を振り分け（.bas / .cls / .frm）
'   ・書き出し前にフォルダ内の旧ファイルを掃除（削除→再生成）
' ここまで。git add / commit / push は【やらない】。
'   → コミットは別途 bat か PAD で実行する想定。
'
' 事前設定（必須）:
'   Excelオプション > トラストセンター > マクロの設定
'     「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェック
'====================================================================

Sub ExportAllModules()
    Dim exportPath As String
    ' ← 出力先（Gitリポジトリ内の任意フォルダ）を指定
    exportPath = "C:\Users\lenovo\マイドライブ\BackupAndRecovery\PAD_vba_flow\vba\"

    ' フォルダ存在チェック
    If Dir(exportPath, vbDirectory) = "" Then
        MsgBox "出力先フォルダが存在しません: " & exportPath, vbExclamation
        Exit Sub
    End If

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject   ' このマクロが入るブック=PERSONAL.XLSB

    ' VBAプロジェクトへのアクセス可否を確認
    On Error GoTo NoAccess
    Dim dummy As Integer
    dummy = vbProj.VBComponents.Count
    On Error GoTo 0

    ' 既存の書き出しファイルを掃除（差分をきれいに保つため）
    Dim f As String
    f = Dir(exportPath & "*.bas"): Do While f <> "": Kill exportPath & f: f = Dir: Loop
    f = Dir(exportPath & "*.cls"): Do While f <> "": Kill exportPath & f: f = Dir: Loop
    f = Dir(exportPath & "*.frm"): Do While f <> "": Kill exportPath & f: f = Dir: Loop
    f = Dir(exportPath & "*.frx"): Do While f <> "": Kill exportPath & f: f = Dir: Loop

    Dim comp As Object, ext As String, cnt As Long
    cnt = 0
    For Each comp In vbProj.VBComponents
        Select Case comp.Type
            Case 1: ext = ".bas"   ' vbext_ct_StdModule   標準モジュール
            Case 2: ext = ".cls"   ' vbext_ct_ClassModule クラス
            Case 3: ext = ".frm"   ' vbext_ct_MSForm      フォーム（.frxも同時出力）
            Case 100: ext = ".cls" ' vbext_ct_Document    ThisWorkbook/シート
            Case Else: ext = ".txt"
        End Select
        ' コードが空のドキュメントモジュールはスキップ
        If comp.Type = 100 And comp.CodeModule.CountOfLines = 0 Then GoTo NextComp
        comp.Export exportPath & comp.Name & ext
        cnt = cnt + 1
NextComp:
    Next comp

    MsgBox cnt & " 個のモジュールを書き出しました。" & vbCrLf & exportPath, vbInformation
    Exit Sub

NoAccess:
    MsgBox "VBAプロジェクトへアクセスできません。" & vbCrLf & _
           "トラストセンターで「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」を有効にしてください。", _
           vbCritical
End Sub
