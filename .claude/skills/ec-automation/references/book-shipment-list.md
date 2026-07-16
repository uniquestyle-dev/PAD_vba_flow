# Book_Shipment_List.xlsx パイプライン仕様

書籍出荷一覧（PDF）を作るための集約パイプライン。3層構成。

## 列構成（固定）

| 列 | 内容 |
|---|---|
| A | 本の色（青・茶・紫・白・赤・茶のおまけ 等） |
| B | 名前（宛名） |
| C | 支払方法（「代引き」「コンビニ支払い」または空欄。**末尾スペースなしで統一**） |

## 3層アーキテクチャと実行順序（厳守）

```
1. Salaciaマクロ（起点）
   Book_Shipment_List.xlsx を SaveAs で新規作成＝リセットし、Salacia分を書き込む
        ↓
2. 追記マクロ群（順不同・追記専用）
   ・Shopify変換一覧表用 / CountBookOrders系:
       入力 orders_export*.csv（Shopify）
       本の色 = Lineitem name を対応表（変換対応表Shopify）で変換
       名前   = Shipping Name（複数商品注文は注文番号でフィルダウン）
       支払方法 = 用途により2系統が存在:
         - NP後払い系: 「NP後払い…」前方一致 →「コンビニ支払い」、他は空欄
         - 代引き系(CountBookOrders改): 「Cash on Delivery (COD)」→「代引き」、他は空欄
   ・STORES変換一覧表用:
       本の色 = 対応表変換 / 名前 = 氏+名 / 支払方法 = 空欄
   ・Repist赤本コンビニ追記:
       入力 = 最新の order_20*.csv（FileDateTime比較で検出、Shift-JIS QueryTable読込）
       本の色 =「赤」固定 / 名前 = B列（お届け先氏名）/ 支払方法 =「コンビニ支払い」固定
        ↓
3. PDF出力マクロ（Book一覧表PDF出力）
   Book_Shipment_List.xlsx を読み、本の色でソート → ヘッダー緑背景・色別セル塗り分け・
   緑罫線・A4縦 → Book_Shipment_List.pdf 出力（xlsxは SaveChanges:=False）
```

## 重要な規則

- **順序が崩れる、またはSalaciaを挟まず追記マクロを複数回走らせると行が重複・累積する。**
  PADフローの順序は「Salacia → 各追記 → PDF出力」を維持。
- 追記マクロは xlsx が無い場合ヘッダ付きで新規作成するフォールバックを持つが、
  これは保険であり正規動作ではない。
- 追記マクロには `Application.Quit` を入れない（後続マクロが止まる）。
  単独完結マクロには入れる。
- QueryTable読込は全列を文字列扱い（TextFileColumnDataTypes指定）にし、
  「2-5-16」等の番地の日付化を防ぐ。この読込部は流用時に一字一句変えない。

## 関連する周辺仕様

- `CountBookOrders`（別系統・SKU集計版）: 代引き注文をSKU別に集計し
  `books_cod_count.txt` へ出力、末尾でCSVを `orders_export_cod.csv` として
  ダウンロードフォルダにFSO.CopyFileでコピー。
  **依存**: `Shopify代引き追記一覧表用` はこのコピー後ファイル（固定パス）を引数に取るため、
  CountBookOrders が先に実行されている必要がある。
- サラシア（サプリ）側には別途 `Salacia_Shipment_List.xlsx` → PDF の系統が存在
  （wsQty.ExportAsFixedFormat をxlsx保存直後・Close前に挿入する方式で安定運用中）。
  書籍系のBook_Shipment_Listと混同しないこと。
