# STORES API 連携仕様（stores_export.py）

管理画面のブラウザ操作（PAD）をAPIに置き換えた実装。2026-06に本番運用開始・検証済み。

## stores_export.py（実装済み・稼働中）

| 項目 | 仕様 |
|---|---|
| 役割 | 未発送オーダーを管理画面CSVエクスポート形式（55列）で出力 |
| 取得 | `GET /retail/202211/orders?delivery_status=waiting`、`limit=100` ページング |
| 認証 | 環境変数 `STORES_API_TOKEN` > `credential.txt`（utf-8-sigで開きBOM除去＋strip のみ。過剰なパースはしない — 過去に`=`含みトークンを破壊し403を起こした） |
| 出力先 | スクリプト先頭の `OUTPUT_DIR` 定数（`C:\Users\lenovo\Desktop\ダウンロード`）。無ければ自動作成 |
| ファイル名 | 実行日 `YYYYMMDD.csv` |
| 文字コード | cp932 / CRLF / BOMなし（管理画面CSVと同一） |
| 行モデル | 1明細＝1行（オーダー直下項目は各行に繰り返し）。複数明細も検証済み |
| レート制限 | 429/503 を指数バックオフで再試行 |
| 依存 | 標準ライブラリのみ（urllib, csv, json）。pip不要 |
| User-Agent | 既定の Python-urllib を避け独自UAを付与（サーバー側拒否回避） |

### PAD連携I/F（規約）

- exit 0＝1件以上 / 2＝0件 / 1等＝エラー（403・通信障害）
- stdout に `ORDER_COUNT=N` センチネル行（日本語ログが化けてもASCIIで確実に拾える）
- stderr に credential 指紋（length・フラグのみ、秘密値なし）を常時出力
  → **CommandErrorOutput 非空＝エラー判定は禁止。ExitCodeで判定**

## 出荷報告（次期実装候補・API仕様確認済み）

- `PATCH /retail/202211/orders/{order_id}/deliveries`
  body: `delivery_method_name` / `tracking_number` / `estimated_arrival_date` / `shipped_mail_message`
- `POST /retail/202211/orders/{order_id}/shipped`（body不要、発送完了処理）
- **注意（仕様書明記）**: `order_id`（オーダーID）必須。オーダー番号では更新不可。
  先に `GET /orders` の `id` を取得して使う。

## API仕様の調査方法（再現手順）

公式OpenAPI仕様はGitHubの retail-api-docs リポジトリにある。Redoc HTMLから埋め込み
JSON仕様を抽出し、`test-data/202211/orders.json` のサンプルデータで列マッピングを検証する
方式が確立済み。新エンドポイント対応時も推測でなくこの仕様書を根拠にする。

## リピストクロス（不採用の経緯）

Data Assistant API（有償）は存在するが、**IPホワイトリスト（固定IP）要件のため不採用**。
リピストクロスは当面PADによる管理画面操作を継続。再検討時はこの前提から。

## 新規スクリプト作成時のテンプレート要件

stores_export.py と同じI/F規約に揃える：
標準ライブラリのみ / credential.txt同居フォルダをカレントに / exit 0・2・1 /
ASCIIセンチネル行 / stderr診断出力 / 429・503リトライ / cp932出力（後続がShift-JIS前提の場合）。
納品前に `python3 -m py_compile` で構文検証し、検証手順（--dry-run→単発→本番）を添える。
