---
name: ec-automation
description: 合同会社ユニーク（ロカボ実験室）のEC受注・出荷自動化システム（Shopify / STORES / リピストクロス × PAD + VBA + Python）の改修・保守・新規フロー設計用スキル。受注CSV処理、Book_Shipment_List.xlsx、出荷ラベル印刷（ClickPost / ゆうプリR / SumatraPDF）、追跡番号登録、PERSONAL.XLSBマクロ、PADのDOSコマンド実行・VBScript連携、stores_export.py 等のキーワードが出たら必ずこのスキルを使うこと。VBAマクロやPythonスクリプトの修正依頼、PADフローのエラー相談、新しい出荷処理の設計依頼でも、EC業務に関わる限りこのスキルを参照する。
---

# ec-automation — ロカボ実験室 EC自動化システム

一人運営のEC事業（書籍CARBOFF・サプリのサラシア等）の受注〜出荷〜追跡番号登録を、
PAD（Power Automate Desktop）+ VBA（PERSONAL.XLSB）+ Python + PowerShell で自動化しているシステム。
このスキルは、その改修・保守・拡張時に守るべき環境情報・アーキテクチャ・設計原則を定義する。

## 使い方（Claude向け指示）

1. まず本ファイルの「設計原則（絶対規則）」を確認する。
2. タスクに応じて references/ を読む：
   - 環境・固定パス・PAD連携規約 → `references/environment.md`
   - Book_Shipment_List.xlsx パイプライン → `references/book-shipment-list.md`
   - STORES API（stores_export.py）→ `references/stores-api.md`
   - 印刷（ClickPost / SumatraPDF / ゆうプリR）→ `references/printing.md`
3. 既存の安定運用コードは指示された箇所以外を変更しない。「それ以外は一切変更したくない」が既定の運用方針。
4. 不明な仕様（列名・マクロの現行実装など）は推測で埋めず、該当ファイルのアップロードを依頼する。

## システム全体像

```
[受注取得]
  Shopify    : API取得 → orders_export*.csv（UTF-8）
  STORES     : stores_export.py がAPIで取得 → YYYYMMDD.csv（cp932）
  リピストクロス: 管理画面操作（PAD）。Data Assistant APIはIP固定要件のため不採用
        ↓
[変換・集約]  PERSONAL.XLSB のVBAマクロ群（PADがVBScript経由で起動）
  ・ゆうプリR用CSV変換、ClickPost用CSV生成
  ・Book_Shipment_List.xlsx（本の色/名前/支払方法の3列）への集約
        ↓
[ラベル発行・印刷]
  ClickPost  : Chrome拡張「ECQのクリックポスト一括登録決済印刷v4」でCSV取込→
               一括登録・決済→ラベルPDF保存 → clickpost_step3.py が
               サマリー生成・PDF結合 → SumatraPDF CLIで一括印刷
  ゆうプリR   : PADでGUI操作（追跡番号発番はゆうプリR経由のみ可能。API代替不可）
  Wizプリント : NP後払い請求書（PAD GUI操作）
        ↓
[追跡番号登録・出荷報告]  各プラットフォームへ反映
```

方針：プラットフォームUI変更に強い「API優先アーキテクチャ」へ段階移行中。
STORESはAPI化済み。ロジックはPython/VBA側に寄せ、PADは起動と画面操作だけの薄い層にする。

## 設計原則（絶対規則）

以下は過去の障害・試行錯誤から確定したルール。改修時に必ず適用する。

1. **PADから日本語パスを変数渡ししない。** cmd（CP932）とPAD変数展開の衝突で失敗する。
   ファイルパスはスクリプト/マクロ内にハードコードし、引数なし実行を基本とする。
2. **cmd /c で呼ぶEXEのパスにクォートを使わない構成にする。** EXE本体は
   スペース・日本語を含まないパス（例: `M:\BackupAndRecovery\recovery\...`）に置く。
   引数側のPDFパス等はクォート可。
3. **CSVはExcelで再保存しない。** エンコーディング破壊を防ぐため、コピーは
   `FileSystemObject.CopyFile` を使う（UTF-8/Shift-JISをそのまま保持）。
4. **VBAの導入はVBEへのコピー&ペースト。** .basのインポートは日本語パス
   （「ダウンロード」等）がShift-JIS絡みで化けるリスクがあるため使わない。
5. **PAD単独起動を想定するマクロは末尾に `Application.Quit`。** 入れ忘れると
   Excelが残留する。逆に、他マクロと連続実行される追記系マクロには入れない。
6. **PADのVBScript呼び出し名とSub名を厳密一致させる。**
   `PERSONAL.XLSB!ModuleNN.SubName` 形式はモジュール名まで一致が必要。
   Sub名のみ形式（`PERSONAL.XLSB!SubName`）はプロジェクト内でSub名一意が前提。
7. **PythonのPAD連携は終了コードで分岐。** `CommandErrorOutput` 非空＝エラーと
   判定しない（診断用stderr出力が常に入るため誤検知する）。
   規約: exit 0＝処理対象あり / 2＝0件 / 1等＝エラー。
   件数はstdoutのASCIIセンチネル行 `ORDER_COUNT=N` を正規表現 `(?<=ORDER_COUNT=)\d+` で抽出。
8. **「DOSコマンドの実行」には必ずタイムアウトを設定**（例120秒）。無限待機を防ぐ。
   ハング疑い時は `py -3 -u`（非バッファ）も検討。
9. **文字数カウント・件数検証はPythonで行う**（管理会社フォーム等の上限確認を含む）。
10. **実行ファイル類は `M:\BackupAndRecovery\` 配下に集約**（PC再インストール時の
    復旧性優先。`C:\SumatraPDF\` のような散在配置はしない）。

## 改修時のワークフロー

1. 対象スクリプト/マクロの現物をアップロードしてもらい、現行実装を確認してから差分を作る。
2. 変更は最小差分。既存の読み込み・変換ロジックは指示がない限り触らない。
3. VBAは構文的に完結した .bas ファイルで納品（導入はコピペ前提と明記）。
4. Pythonは `python3 -m py_compile` 等で構文検証してから納品。
5. 検証手順を必ず添える：`--dry-run` → 単発テスト → 本番、の順。
6. 実行順序の依存関係（例: Salaciaマクロが先、CountBookOrdersが先）を変更・追加した
   場合は、PADフロー側の順序変更もセットで案内する。
