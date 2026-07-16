# PAD_vba_flow — リポジトリ運用ルール

合同会社ユニーク（ロカボ実験室）のEC自動化における、VBA（PERSONAL.XLSB）と
PADフローのソース管理リポジトリ。このファイルは全セッションで常時参照される前提の「憲法」。
業務ロジックの詳細規約は `ec-automation` スキルに従う（本ファイルはリポジトリ運用に限定）。

## 構成

- `vba/` … PERSONAL.XLSB の全モジュール（.bas / .cls）。**Gitの正本**。
- `flow/` … PAD各フローのRobinテキスト（*.robin）。稼働フローの復元元ではなく、
  フロー構成をClaude Codeへ伝える**文脈用の写し**。詳細は `pad-flows` スキル参照。
- `import_vba.ps1` … `vba/` を PERSONAL.XLSB へ反映するスクリプト（A方式）。
- `.gitignore` … `*.xlsb` 本体・一時ファイル・`flow/#過去/` を除外。

## 正本と実行環境の非対称性（最重要）

- Gitの正本は `vba/*.bas`。実行環境は `%APPDATA%\Microsoft\Excel\XLSTART\PERSONAL.XLSB`。
- **`.bas` を編集・コミットしただけでは実行環境は変わらない。**
  `import_vba.ps1` の実行で初めて PERSONAL.XLSB に反映される。
- PERSONAL.XLSB 本体はGit管理外。別途タスクスケジューラで日次バックアップ済み。

## 日常ループ

1. `vba/*.bas` を編集。
2. **Excelを完全に閉じる**（起動中はimportが競合し中断する）。
3. import 実行：
   `powershell -ExecutionPolicy Bypass -File "<repo>\import_vba.ps1"`
4. Excelを開き、VBEで反映を確認・動作テスト。
5. コミット：`git add .` → `git commit -m "..."`。区切りで `git push`（＝実質バックアップ）。

## 厳守事項

- import前にExcelを閉じる。
- `import_vba.ps1` は **BOM付きUTF-8** で保存を維持する。Windows PowerShell 5.1 は
  BOM無しをCP932と誤読し日本語が化ける。`pwsh`（PowerShell 7）実行なら不問。
- ドキュメントモジュール（ThisWorkbook 等）は Import 不可のため、スクリプトが
  コード行差し替えで処理している。この分岐を壊さない。
- 指示された箇所以外の安定コードは変更しない（「それ以外は変更しない」が既定方針）。
- 不明な現行仕様は推測で埋めず、現物の確認を求める。

## 既知の不整合（要解消）

- `ec-automation` 設計原則 #4「.basインポートは使わずコピペ」は、本リポジトリの
  A方式（COMインポート）と矛盾する。A方式は対象を `vba/` に限定しエンコーディングを
  制御することで #4 の懸念（日本語DLパスでの化け）を回避しており、実運用で成功済み。
  → **#4 の記述をA方式許容へ更新すること。**

## 関連スキル（.claude/skills/ に配置）

- `ec-automation` … EC自動化の環境・固定パス・設計原則。
- `pad-flows` … `flow/*.robin` のフロー構成マップ。
