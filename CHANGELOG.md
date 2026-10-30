# 変更の記録

この形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に倣う。版は [Semantic Versioning](https://semver.org/lang/ja/) に従う（1.0.0 までは小さい版の上げでも互換を崩すことがある）。

## [Unreleased]

InterHub のグラフ（ADR-0031・ISSUE-101〜111）を取り込んだ（M7）。

### 追加

- グラフの表示：縦棒・横棒（集合・積み上げ・100% 積み上げ）、折れ線（マーカー含む）、円・ドーナツと、タイトル・凡例・項目軸・数値軸・目盛線を、保存された値（`c:numCache`・`c:strCache`）から SVG で描く。描けない種類・失敗は今までどおり種類名の破線枠。グラフは明るい表示・ダークとも Excel の色のまま描く。
- `XlsxBook.readChart(index, path)` とグラフの型（`ChartModel`・`ChartResult` など）。`SheetDrawing.chartPath`。
- 上限 `maxChartSeries`・`maxChartPoints`・`maxChartTextChars`・`maxSheetChartPoints`。
- グラフの描画は遅れて読む部品：バンドラ向けでは別のチャンク、単独版では `xsheet-chart.mjs`（`xsheet.mjs` と同じフォルダに置く。`xsheet.min.js` は中に持つ）。

### 変更

- `XlsxBook` に `readChart` が増えた（`XlsxBook` を自分で実装している場合は足す）。

## [0.2.0] - 2026-10-30

誰でも軽く使える配り方にした（ADR-0003・M6）。`<script>` 1 つ・`import` 1 つ・1 行で表示できる。0.1.0 の公開物はすべて残る。

### 追加

- `view(target, source, options?)`：xlsx をタブ付きで 1 行で表示する。`target` は要素か CSS セレクタ、`source` は File・Blob・ArrayBuffer・型付き配列・Response とその Promise（`fetch(url)` をそのまま渡せる。ライブラリは通信しない）。戻り値の `Viewer` は `sheets`・`showSheet(番号か名前)`・`destroy()`。
  - `lang`（`ja`・`en`・`ko`・`auto`。既定 `auto`）・`theme`（`light`・`dark`・`auto`）・`css`（表示の CSS を自動で入れる。nonce 付き・読み込み済みなら入れない）・`height`（container の高さが 0 なら 480px と注意を 1 回）。
- `read(source, options?)`：DOM なしで開く（`openXlsx` と同じ。受け取れる形は `view` と同じ）。
- 文言 `MESSAGES_KO`（韓国語・23 キー）。
- ダークの値を `style.css` に同梱：`data-xs-theme="dark"`（`auto` は `prefers-color-scheme` に従う）を表示の根か祖先に付ける。
- `renderWorkbook` の戻り値に `sheets`（シートの一覧）。
- 単独版（`dist/` の根）：`xsheet.min.js`（`<script>` 1 つ・大域 `XSheet`・依存と形の表を中に持つ）・`xsheet.mjs`＋`xsheet-geom.mjs`（バンドラ無しの import）・`xsheet.css`・`xsheet.d.ts`・`xsheet.d.mts`・`xsheet.licenses.txt`（中に持つ依存のライセンス）。GitHub Release に zip（`viewer.html` 付き）を添付する。
- `examples/standalone.html`（ビルド不要・`file://` でも開けるビューア）・`examples/esm.html`。
- 開発：InterHub の変更を取り込む道具 `npm run sync:status`・`sync:patch`（`tools/sync/interhub.json`）。`npm run check:standalone`・`npm run release:zip`。

### 変更

- 入口 `xsheet` が、`view`・`read` と core・render の公開物をすべて出すようにした（0.1.0 の `xsheet` の公開物はそのまま残る）。DOM・Node 固有の API を使わない読み取りだけの入口は `xsheet/core`。

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

[0.2.0]: https://github.com/TaekangKwon/XSheet/releases/tag/v0.2.0
[0.1.0]: https://github.com/TaekangKwon/XSheet/releases/tag/v0.1.0
