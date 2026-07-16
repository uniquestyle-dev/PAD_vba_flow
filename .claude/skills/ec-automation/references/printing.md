# 出荷ラベル印刷系仕様（ClickPost / SumatraPDF / ゆうプリR / Wizプリント）

## ClickPost フロー（4ステップ構成）

```
Step1: 商品別ClickPost用CSV生成（タブ区切り・CP932）
       ファイル名 = 商品名（例: carboff.csv, Medico.csv, salasiru.csv）という命名規約
Step2: ClickPostにCSV登録→決済→ラベルを印刷せずPDF保存
       保存先 = C:\Users\lenovo\Desktop\ダウンロード\、ファイル名 = CSVと同名（.pdf）
       操作はChrome拡張「ECQのクリックポスト一括登録決済印刷v4」経由
       （ClickPost本体のUI変更を拡張側が吸収）
Step3: clickpost_step3.py
       ・各PDFの存在とページ数（＝件数）を集計しサマリーページ生成
       ・商品マスタと実在PDFを突合し0件商品もサマリーに記載（0件は区切りページなし）
       ・1件以上の商品は区切りページ＋ラベルPDFの順で結合
       ・出力: clickpost_merged_YYYYMMDD.pdf（Python側 today.strftime("%Y%m%d")）
       ・PRODUCT_MASTER に "print": True/False フラグあり。False（salacia_big /
         salacia_small）は結合PDF・サマリーページから除外するが、
         サマリーメールの件数には含める
Step4: 出荷報告（ラベル印刷と独立実行可）
```

Step3後の印刷はPADの「DOSコマンドの実行」でSumatraPDF CLIを直接呼ぶ（バッチ経由しない）：

```
cmd /c M:\BackupAndRecovery\recovery\SumatraPDF-3.5.2-64\SumatraPDF-3.5.2-64.exe -print-to "Canon G6000 series" -print-settings "monochrome" "C:\Users\lenovo\Desktop\ダウンロード\clickpost_merged_%FormattedDateTime%.pdf"
```

- 成功条件: EXEパスがスペース・日本語なし（クォート不要）であること。
  cmd /c はEXEパス側にクォートが必要な構成だと解釈に失敗する（過去の障害原因）。
- カラー印刷は `-print-settings "monochrome"` を外す。
- `%FormattedDateTime%` はPADカスタム書式 `yyyyMMdd`。

## ゆうプリR

- 追跡番号の発番は日本郵便システムと通信するゆうプリRのGUIでのみ可能。
  **API・自前PDF生成での代替は不可**（追跡番号が出ない）。PAD GUI操作を継続する領域。
- フロー: CSV取込 → 発番 → 印刷 → 追跡番号出力。

## Wizプリント（NP後払い請求書）

- GUIアプリのためPAD操作。将来的にはNPコネクトプロAPI＋PDF自動生成で置換可能性あり
  （未実装。API利用は連携カート経由でなければ初期・月額費用が発生する点に注意）。

## 障害切り分けの定石

1. PADからの印刷不発 → まず `cmd /c dir <EXEパス> > test.txt 2>&1` でパス可視性を確認。
2. M:ドライブ（subst）はPAD実行コンテキストから見える（確認済み）。
3. クォート問題を疑う場合はバッチを捨てEXE直呼びに単純化して切り分ける。
