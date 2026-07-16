# 環境情報・固定パス・PAD連携規約

最終確認: 2026-07（会話ログ由来。変更時はこのファイルを更新すること）

## マシン環境

| 項目 | 値 |
|---|---|
| OS / 端末 | Windows（ユーザー名 lenovo） |
| Python | 3.14（`C:\Users\lenovo\AppData\Local\Python\pythoncore-3.14-64\python.exe`、通常は `py -3` で起動） |
| Excel | PERSONAL.XLSB = `C:\Users\lenovo\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB` |
| プリンター | Canon G6000 series |
| M: ドライブ | subst によるマッピング。`M:\BackupAndRecovery\` に実行ファイル類を集約 |
| Googleドライブ | `C:\Users\lenovo\マイドライブ\` はミラーリング（ローカル実体あり） |

## 固定パス

| 用途 | パス |
|---|---|
| 受注CSV等の入出力 | `C:\Users\lenovo\Desktop\ダウンロード\` |
| PADスクリプト置き場 | `C:\Users\lenovo\マイドライブ\BackupAndRecovery\PAD_file\` |
| API系スクリプト+credential | `C:\Users\lenovo\マイドライブ\BackupAndRecovery\PAD_file_API\`（stores_export.py・credential.txt。PADの作業フォルダーに指定） |
| SumatraPDF | `M:\BackupAndRecovery\recovery\SumatraPDF-3.5.2-64\SumatraPDF-3.5.2-64.exe` |
| MyASPインポート | `C:\Users\lenovo\マイドライブ\BackupAndRecovery\PAD_file\myasp_cleaning\import\` |

## PAD作業フォルダ

| 項目 | 値 |
|---|---|
| パス | `C:\Users\lenovo\Desktop\ダウンロード\` |
| 役割 | 受注CSV・出荷CSV等の一時入出力先。PADフロー・VBA・Pythonが共通で使う作業領域 |

運用上の制約：
- **日本語パス問題**: このパスは設計原則 #1「PADから日本語パスを変数渡ししない」の直接の対象。
  `%LatestFile%` 等のPAD変数経由で渡さず、スクリプト／マクロ内にハードコードする。
- **バックアップ対象外**: Googleドライブ同期領域（`M:\BackupAndRecovery\`）の外にあり、
  同期・バックアップの保証がない。設計原則 #10（実行ファイルは `M:\BackupAndRecovery\` に集約）
  の例外にあたる。ここに置くのは処理途中の一時ファイルのみとし、最終成果物や永続データは置かない。

## PAD → VBA（VBScript経由）の標準パターン

```vbscript
On Error Resume Next
Dim xlApp
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.AutomationSecurity = 1
xlApp.Workbooks.Open xlApp.StartupPath & "\PERSONAL.XLSB"
xlApp.Run "PERSONAL.XLSB!CountBookOrders", "<引数(固定パス推奨)>"
xlApp.Quit
Set xlApp = Nothing
```

注意点：
- `%LatestFile[0]%` 等のPAD変数を引数に使うのは日本語パスで不安定。固定パス渡しが確定方針
  （例: `C:\Users\lenovo\Desktop\ダウンロード\orders_export_cod.csv`）。
- `PERSONAL.XLSB!ModuleNN.SubName` はモジュール番号がインポートごとに変わるため、
  可能ならSub名のみ形式＋プロジェクト内Sub名一意を推奨。既存呼び出し形式は変更しない。

## PAD → Python（DOSコマンドの実行）の標準パターン

- コマンド: `py -3 script.py`（解決不可なら絶対パスのpython.exeに置換）
- 作業フォルダー: credential等の置き場（PAD_file_API）に設定＝実行時カレント
- 「タイムアウト後に失敗します」ON（例120秒）
- 分岐は `%CommandExitCode%` のSwitch: 0 / 2 / Else
- 件数抽出: 「テキストの解析」で正規表現 `(?<=ORDER_COUNT=)\d+`（PADは.NET正規表現、後読み可）
- エラー時は `%CommandErrorOutput%` をエラーメール本文へ（ただしエラー判定自体はExitCodeで）

## 印刷コマンド（確定形）

```
モノクロ:
cmd /c M:\BackupAndRecovery\recovery\SumatraPDF-3.5.2-64\SumatraPDF-3.5.2-64.exe -print-to "Canon G6000 series" -print-settings "monochrome" "C:\Users\lenovo\Desktop\ダウンロード\clickpost_merged_%FormattedDateTime%.pdf"

カラー: -print-settings "monochrome" を外す
```

`%FormattedDateTime%` はPAD側で `yyyyMMdd` 書式。Python側の `date_file = today.strftime("%Y%m%d")` と対応。

## エンコーディング規約

| データ | エンコーディング |
|---|---|
| Shopify orders_export*.csv | UTF-8（Excel再保存禁止、コピーはFSO.CopyFile） |
| STORES 管理画面形式CSV / stores_export.py 出力 | cp932 / CRLF / BOMなし |
| リピストクロス order_20*.csv | Shift-JIS（VBAはQueryTable TextFilePlatform=932で読込） |
| ClickPost用CSV | タブ区切り・CP932 |
| VBAからのUTF-8出力 | ADODB.Stream（Charset=UTF-8、BOM付き） |
