# 変更の記録

この形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に倣う。版は [Semantic Versioning](https://semver.org/lang/ja/) に従う（1.0.0 までは小さい版の上げでも互換を崩すことがある）。

## [Unreleased]

### 追加

- `findText(sheet, query, options?)`：シートの中の文字を探す（[#28](https://github.com/TaekangKwon/XSheet/issues/28)。[ADR-0007](docs/adr/0007-moobius-viewer-api.md) §1。`xsheet/core`・`xsheet`・単独版の `XSheet.findText`）。表示形式を当てた後の文字（`CellView.text`）で照らし、`{ matches, truncated }` を返す。一致は行・列（1 始まり）と文字の中の位置。大文字と小文字・改行は既定で区別せず（位置は元の文字で返す）、全角と半角は区別する。結合は左上のセルだけ、非表示の行・列は既定で探さない。約 10ms ごとに休むので大きなシートでも画面を止めない。`signal` で中断（`AbortError`）、`maxMatches`（既定 10,000・上限 100,000）で打ち切る。
- `SheetView`・`WorkbookView` の `setHighlights(matches)`：`findText` の一致を表の上で光らせる（[#29](https://github.com/TaekangKwon/XSheet/issues/29)。[ADR-0007](docs/adr/0007-moobius-viewer-api.md) §2）。`null`・空の配列で消す。スクロールで描き直しても残り、シートを替えると消える。一致の片は `span.xs-hit`（`data-xs-hit` に配列の番号）で、セルの文字全体を `span.xs-hits` 1 つで包む（空白を保つ）。色は新しい CSS 変数 `--xs-hit-bg`（既定 `rgba(255, 255, 0, 0.4)`）・`--xs-hit-current-bg`（既定 `#ffa500`。ダークの値もある）。背景だけで、行の高さ・列の幅は変えない。`###` で出す数値のセルはセル全体を光らせる。

## [0.4.1] - 2026-11-12

曜日の書式（`aaa`・`aaaa`）の修正を先に出す版（[ADR-0007](docs/adr/0007-moobius-viewer-api.md) §9）。表とシートタブの見た目を Excel に近づけた。公開 API は変えていない。

### 追加

- GitHub Pages のデモ（https://taekangkwon.github.io/XSheet/。[ADR-0005](docs/adr/0005-pages-demo.md)）。例のページを `npm run pages:build` でビルドし、`main` への push で Actions が出す。例のページのサンプルはページからの相対で読むようにした。

### 修正

- 曜日の書式（`aaa`＝「月」・`aaaa`＝「月曜日」）を Excel と同じに出すようにした（[#25](https://github.com/TaekangKwon/XSheet/issues/25)。[ADR-0007](docs/adr/0007-moobius-viewer-api.md) §6）。今までは `m/d(aaa)` が `11/11(aaa)` と文字のまま出ていた。曜日は日付の通し番号から計算する（1900 年基準・1904 年基準。時間帯のずれは入らない）。Excel と同じく、a が 5 個以上なら長い形、大文字の `AAA` も曜日、引用符・`\`・`[...]` の中と `AM/PM`・`A/P` の a は曜日にしない。日付の範囲の外（Excel が #### にする値）は空にする。セルの文字（`CellView.text`）とグラフの軸の文字の両方に効く。Excel 16 で作った `fixtures/excel/ref-weekday.xlsx` と照合した。

### 変更

- 表とシートタブを Excel に近い見た目にした（[#23](https://github.com/TaekangKwon/XSheet/issues/23)。CSS だけで、DOM・公開 API は変えない）。
  - 表：セルの地を白（`--xs-cell-bg` `#ffffff`）にし、グリッド線（`--xs-grid` `#dcdfe4`）と見出しの区切り（`--xs-border` `#c4c8cf`）を少し濃くした。見出しは太字にせず、行番号は中央に置く。**`renderSheet` だけを使う画面も色が変わる。** 元の色に戻すには 3 つの変数を `#f5f6f8`・`#e5e8ee`・`#d0d0d0` で上書きする。ダークの値は変えていない。
  - タブ列：表に付けて境に線を引く。タブは角の無い枠を縦線で区切り、選んだタブは表と同じ地・強調色の太字と棒で示す。◀ ▶ は左にまとめる。非表示のシートは斜体・破線ではなく薄い文字にした。

## [0.4.0] - 2026-11-05

リポジトリを公開して最初の版（[ADR-0004](docs/adr/0004-public-release.md)）。シートタブの位置を選べるようにした。0.3.0 の公開物はすべて残る。

### 追加

- シートタブの位置を選ぶ指定 `tabPosition`（`'top'`・`'bottom'`。既定 `'top'`）を `renderWorkbook`・`view` に足した（[#17](https://github.com/TaekangKwon/XSheet/issues/17)）。`'bottom'` ではタブ列を表の下（DOM でもシートの表示の後ろ）に置き、根に `xs-tabs-bottom` を付ける。

### 変更

- ライセンスを MIT にした（`LICENSE`。`package.json` の `private` を外し、`files` と単独版の zip に `LICENSE` を足した）。
- README を英語にし、日本語（`README.ja.md`）・韓国語（`README.ko.md`）の訳を置いた。
- `SECURITY.md` と CI（GitHub Actions）を足した。

## [0.3.0] - 2026-10-30

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

[Unreleased]: https://github.com/TaekangKwon/XSheet/compare/v0.4.1...HEAD
[0.4.1]: https://github.com/TaekangKwon/XSheet/releases/tag/v0.4.1
[0.4.0]: https://github.com/TaekangKwon/XSheet/releases/tag/v0.4.0
[0.3.0]: https://github.com/TaekangKwon/XSheet/releases/tag/v0.3.0
[0.2.0]: https://github.com/TaekangKwon/XSheet/releases/tag/v0.2.0
[0.1.0]: https://github.com/TaekangKwon/XSheet/releases/tag/v0.1.0
