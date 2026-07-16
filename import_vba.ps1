# ============================================================
# import_vba.ps1
#   リポジトリの .bas / .cls を PERSONAL.XLSB へ反映（インポート）する。
#
#   このスクリプトが「やること／やらないこと」
#     ○ vba/ 内の .bas .cls (.frm) を PERSONAL.XLSB へ取り込み、保存
#     ○ 標準/クラス/フォーム = 既存を削除して Import で置換
#     ○ ThisWorkbook 等のドキュメントモジュール = コード行を差し替え
#     × git add / commit（別途こちらで実行）
#     × エクスポート（PERSONAL.XLSB → .bas。初期化用の別マクロで実施）
#
#   前提:
#     ・Excelを完全に閉じてから実行（起動中は競合するため中断する）
#     ・Excel トラストセンターで
#       「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」を有効化済み
# ============================================================

$ErrorActionPreference = "Stop"

# ==== 設定 ====
$VbaDir = "C:\Users\lenovo\マイドライブ\BackupAndRecovery\PAD_vba_flow\vba"
$PersonalPath = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART\PERSONAL.XLSB"
# ==============

# 定数（VBComponent.Type）
$VBEXT_CT_DOCUMENT = 100

# ---- ドキュメントモジュール用: エクスポートヘッダを除去してコード本体を返す ----
function Get-VbaCodeBody {
    param([string[]]$Lines)
    $body = New-Object System.Collections.Generic.List[string]
    $i = 0; $n = $Lines.Count
    if ($i -lt $n -and $Lines[$i] -match '^\s*VERSION\s')  { $i++ }          # VERSION 行
    if ($i -lt $n -and $Lines[$i] -match '^\s*BEGIN\b') {                    # BEGIN..END ブロック
        $i++
        while ($i -lt $n -and $Lines[$i] -notmatch '^\s*END\b') { $i++ }
        if ($i -lt $n) { $i++ }
    }
    while ($i -lt $n -and $Lines[$i] -match '^\s*Attribute\s') { $i++ }      # 先頭 Attribute 群
    for (; $i -lt $n; $i++) { $body.Add($Lines[$i]) }
    return ($body -join "`r`n")
}

# ---- 事前チェック ----
if (-not (Test-Path $VbaDir))       { throw "vbaフォルダが見つかりません: $VbaDir" }
if (-not (Test-Path $PersonalPath)) { throw "PERSONAL.XLSB が見つかりません: $PersonalPath" }

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    throw "Excelが起動中です。完全に閉じてから再実行してください。"
}

# ---- Excel 起動（イベント抑止でWorkbook_Open等を走らせない）----
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false

$wb = $null
try {
    # XLSTART自動読込で既に開いていれば流用、なければパスから開く
    foreach ($w in $excel.Workbooks) {
        if ($w.Name -ieq "PERSONAL.XLSB") { $wb = $w; break }
    }
    if ($null -eq $wb) { $wb = $excel.Workbooks.Open($PersonalPath) }

    # VBProjectアクセス可否
    try { $null = $wb.VBProject.VBComponents.Count }
    catch {
        throw "VBAプロジェクトへアクセスできません。トラストセンターで『VBAプロジェクトオブジェクトモデルへのアクセスを信頼する』を有効にしてください。"
    }
    $proj = $wb.VBProject

    $files = Get-ChildItem -Path $VbaDir -File |
             Where-Object { $_.Extension -in ".bas", ".cls", ".frm" }

    $replacedDoc = 0; $importedMod = 0
    foreach ($f in $files) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)

        # 同名の既存コンポーネントを探す
        $existing = $null
        foreach ($c in $proj.VBComponents) {
            if ($c.Name -ieq $name) { $existing = $c; break }
        }

        if ($null -ne $existing -and $existing.Type -eq $VBEXT_CT_DOCUMENT) {
            # ThisWorkbook / シート: Import不可のためコード行を差し替え
            # 注: VBEのエクスポートは日本語WindowsではCP932(Shift-JIS)のため932で読む
            $enc  = [System.Text.Encoding]::GetEncoding(932)
            $text = [System.IO.File]::ReadAllText($f.FullName, $enc)
            $lines = $text -split "`r`n|`n"
            $code = Get-VbaCodeBody $lines

            $cm = $existing.CodeModule
            if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
            if ($code.Trim().Length -gt 0) { $cm.AddFromString($code) }
            $replacedDoc++
        }
        else {
            # 標準/クラス/フォーム: 既存を削除して Import（Importが符号化を処理）
            if ($null -ne $existing) { $proj.VBComponents.Remove($existing) }
            $null = $proj.VBComponents.Import($f.FullName)
            $importedMod++
        }
    }

    $wb.Save()
    Write-Host ("インポート完了: 標準/クラス {0} 件, ドキュメント {1} 件を反映しました。" -f $importedMod, $replacedDoc) -ForegroundColor Green
}
finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
