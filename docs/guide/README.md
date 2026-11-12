# 使い方の手引き

| 項目 | 値 |
| --- | --- |
| 文書 | XSheet 使い方の手引き（目次） |
| 更新日 | 2026-10-30 |
| 関連 | [../../README.md](../../README.md)（3 行で始める・指定の一覧）・[../02_architecture.md](../02_architecture.md)（設計）・[ADR-0003](../adr/0003-universal-distribution.md) |

## TL;DR
- 最初は README の「3 行で始める」を読む。ここは、そこから先の場面ごとの説明。
- 表示は `view`（`xsheet` か単独版の `XSheet.view`）、読み取りだけなら `xsheet/core` の `openXlsx` か `xsheet` の `read`。
- コードの例は今の API（v0.4.1）に合わせてある。

---

| 文書 | 内容 |
| --- | --- |
| [standalone.md](standalone.md) | 単独版（`xsheet.min.js`・`xsheet.mjs`）：ファイルの一覧・`file://` と http・zip の中身・ライセンスの文・TypeScript |
| [frameworks.md](frameworks.md) | React・Vue での置き方（mount で `view`、unmount で `destroy`。非同期の競合と StrictMode） |
| [worker.md](worker.md) | Web Worker で読む（`xsheet/core`・`ArrayBuffer` の転送）。表示は main thread |
| [csp.md](csp.md) | CSP の指定（`script-src`・`style-src` の nonce・`img-src blob:`）と、`css: 'none'` で `<link>` にする形 |
| [theme.md](theme.md) | ダーク表示（`theme`・`data-xs-theme`）・変数の上書き・Shadow DOM |
