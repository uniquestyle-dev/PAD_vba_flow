---
name: pad-flows
description: PAD_vba_flowリポジトリのPADフロー構成を参照する。flow/*.robin が何をするフローで、どのVBAマクロ（PERSONAL.XLSB）を呼ぶかを把握したいときに使う。特にVBAマクロを修正する際、そのマクロがどのフローからどの順序で呼ばれるかを確認し、影響範囲や前後関係の取り違えを防ぐ目的で参照する。
---

# pad-flows — PADフロー構成マップ

`flow/*.robin` は稼働中PADフローのテキスト写し（実行はPAD側。ここでは**構成把握のための地図**として使う）。
VBA修正時に「そのマクロがどのフローから、どの順で呼ばれるか」を取り違えないために参照する。

## 使い方

1. 下の一覧で対象フローの目的と、関係しそうなフローの当たりをつける。
2. アクション単位の詳細が要るときだけ、該当 `flow/NN.robin` を開く。Robinは1アクション＝
   複数行で冗長なため、必要なフローだけを読む（全量を読み込まない）。
3. VBAの呼び出し関係を確認するときは、対象 `.robin` 内の `PERSONAL.XLSB` 参照箇所
   （`PERSONAL.XLSB!ModuleNN.SubName` 形式）を探す。ここがマクロとフローの接続点。

## フロー一覧

purpose は退避フォルダ（`flow/#過去/…`）の旧ファイル名からの**推定（要確認）**。
旧名の対応が無い 82・84 は **不明**。VBA呼び出し欄は各 `.robin` を読んで確定済み。

| file | 推定目的（要確認） | 呼び出すVBA（.robinで確認） |
|---|---|---|
| 00-1.robin | 出荷ラベル印刷 — サプリ | `サラシア変換サプリ一覧表`（×2）, `サラシア変換サプリ`, `TransferTrackingNumberCOD` |
| 00-2.robin | 出荷ラベル印刷 — 本 | `CountBookOrders`, `ExportCodNamePdfs`, `Module22.DaibikiConvert`, `Module23.代引きお問い合わせ番号追記`, `Module20.STORES変換`, `Module37.STORES変換一覧表用`, `Module19.Shopify変換`, `Module38.Shopify変換一覧表用`, `Shopify代引き追記一覧表用`, `サラシア変換`, `Module39.サラシア変換一覧表用`, `Book一覧表PDF出力`, `ExportNamePdf`, `TransferTrackingNumberShopify`, `TransferTrackingNumberSalacia`, `TransferTrackingNumberStores` |
| 01.robin | Amazon Pay 関連（amazonpaypass） | なし |
| 02.robin | 定期購入者メアド登録 | なし |
| 03.robin | 単発購入者メアド登録 | なし |
| 04.robin | カゴ落ち販促メール送信 | なし |
| 05.robin | 定期停止者のエキスパタグ変更 | なし |
| 06.robin | 1か月コースの90日→30日 変更 | なし |
| 07.robin | STORES受注CSV取得と通知 | なし |
| 08.robin | LPのPV/UUと購入クリックの転記 | なし |
| 09.robin | 購入ユーザ登録と件数集計 | `Module17.Shopify登録確認_Shift_w`, `Module18.ユーザ登録と確認`, `Module21.リピスト書籍購入ユーザ変換` |
| 80.robin | クリーニング（日刊メルマガ） | なし |
| 82.robin | クリーニング_購入者通信-API | なし |
| 83.robin | クリーニング — 除外（日刊メルマガ） | なし |
| 84.robin | クリーニング-除外_購入者通信 | なし |

## .robin の行番号とPAD GUIステップ番号

`.robin` はRobin DSLのテキスト直列化であり、PAD GUIのステップ番号とは **1:1対応しない**。
VBScript実行アクション1つがスクリプト本文の長さに応じて数十行に展開されるため、
**`.robin` 行番号からPAD GUIステップ番号を推測することはできない。**

- 例: 00-2.robin のブランド結合ループは .robin テキスト行260付近、PAD GUIではステップ106付近。
- ステップ番号が必要な場合はPAD GUIのスクリーンショットで確認する。
- `.robin` 内の行番号を会話に出す場合は「.robin L260」等と明記し、PADステップ番号と混同しない。
- **フロー内の該当箇所を案内するときは、PADの検索窓で見つけやすいキーワードを添える。**
  変数名（`BrandList`、`BaseFile`等）やマクロ名（`STORES変換`等）が有効。
  VBScript内部の文字列はPAD検索でヒットしない場合があるため、アクション直下の変数名を優先する。

## 不具合調査時の注意

`.robin` のロジックだけで原因を推定した場合でも、PAD作業フォルダ
（`C:\Users\lenovo\Desktop\ダウンロード\` — 全フロー共通、`environment.md` に定義済み）
の実ファイルを確認してから結論を出すこと。ロジック上の仮説と実データの突合を省くと、
別の原因（ファイル名の違い、エンコーディング問題、前段マクロの出力先変更等）を見落とすリスクがある。

## 保守

- フローを改修したら、該当 `.robin` を再コピーして更新する（この地図の鮮度＝手動更新に依存）。
- 目的・VBA呼び出し欄を実物で確定したら「要確認／不明」を書き換える。
