# 変更の記録

この形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に倣う。版は [Semantic Versioning](https://semver.org/lang/ja/) に従う（1.0.0 までは小さい版の上げでも互換を崩すことがある）。

## [0.1.0] - 2026-10-29

最初の版。InterHub（コミット `943e4f0`）の .xlsx の読み取りと表示を、どのプロジェクトからも使えるライブラリにした。

### 追加

- `xsheet`（core）：`openXlsx`・`XlsxBook`（`sheets`・`openSheet`・`openDrawings`・`readImage`・`close`）・`XlsxSheet`（`getRows`・寸法・書式と罫線の表）・`XlsxError`・`DEFAULT_LIMITS`（上限 25 項目）・寸法の換算。DOM・Node 固有の API を使わない。
- `xsheet/render`：`renderWorkbook`（タブ付き・非表示シートの印・開けないファイルの文言）・`renderSheet`（1 シート）。指定は `messages`・`drawings`・`nonce`・`limits`。`destroy()` でリスナ・blob URL・ブックを放す。
- 文言 `MESSAGES_JA`（既定）・`MESSAGES_EN`（23 キー）。
- `xsheet/style.css`：CSS 変数 14 個（明るい既定値は詳細度 0）。スクロール枠・固定の見出し・タブ。
- `xsheet/geom`：図形の形の表（プリセット 176 種）と式の評価。render は図形を描くときに動的 import する。
- CSP：`script-src 'self'`・`style-src` の nonce・`img-src blob:` で動く。
- `examples/`：ファイルを選んで表示する例のページ（明暗・文言・図の切り替え、nonce だけの CSP）。

### InterHub からの違い

- クラス名・変数・属性の接頭辞を `ih-` から `xs-` にした（`xs-*`・`--xs-*`・`data-xs-fmt`）。
- 表示の状態をモジュール変数ではなく表示ごとのインスタンスに持ち、同じページに複数の表示を置けるようにした。
- 形の評価器を描画側から独立させ（`xsheet/geom`）、core が描画に依存する逆転をなくした。

[0.1.0]: https://github.com/TaekangKwon/XSheet/releases/tag/v0.1.0
