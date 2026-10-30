import { l as XlsxError, p as colName, w as DEFAULT_LIMITS } from "./chunks/color-sme5jHOH.js";
import { a as ptToPx, c as isPaperFill, i as defaultColWidthToPx, n as colWidthToPx, o as fontDefaultRowPx, r as createGeometry, s as lookupMdw, t as openXlsx } from "./chunks/core-CGDQdCgf.js";
import { a as MESSAGES_KO, i as MESSAGES_JA, n as renderSheet, o as resolveMessages, r as MESSAGES_EN, t as renderWorkbook } from "./chunks/render-C4lxLF3M.js";
// src/facade/source.ts
var isBufferLike = (v) => typeof v === "object" && v !== null && typeof v.arrayBuffer === "function";
var isArrayBuffer = (v) => Object.prototype.toString.call(v) === "[object ArrayBuffer]";
/**
* 入れ物を Uint8Array にする。
* @throws TypeError 受けない形（URL の文字列など）・失敗した `Response`（`ok` が false）
*/
async function toSourceBytes(source) {
	const value = await source;
	if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
	if (isArrayBuffer(value)) return new Uint8Array(value);
	if (isBufferLike(value)) {
		if (value.ok === false) throw new TypeError(`XSheet: the response failed (HTTP ${String(value.status)})`);
		return new Uint8Array(await value.arrayBuffer());
	}
	if (typeof value === "string") throw new TypeError("XSheet: a URL string is not accepted (XSheet never fetches). Pass fetch(url) or a File/Blob/ArrayBuffer instead");
	throw new TypeError("XSheet: the source must be a File, Blob, ArrayBuffer, typed array or Response");
}
// src/style.css?inline
var style_default = "/*\n * XSheet の表示の規則（`xsheet/style.css`）。\n * 正: InterHub src_V1/src/styles/global.css（943e4f0）の .xlsx の部分と、表示が使う共用の規則\n *     （`.ih-data-view`・`.ih-data-body--table`・`.ih-data-banner`・`.ih-dv-*`・`.ih-toolbar`・`.ih-toggle`）。\n *     接頭辞は `xs-`（D6）。対応は docs/03_source-map.md の「M3-1 で写した結果」。\n *\n * - 変数は明るい既定値だけを持つ（docs/02_architecture.md §6）。ダーク表示は使う側が `--xs-*` を上書きする。\n *   既定値は詳細度 0 の `:where()` に置くため、使う側の `:root`・テーマの属性・コンテナのどこで上書きしても勝つ。\n * - Excel の色（塗り・文字の色・図形の色）は変数ではなく、表示の根の `<style>`（書式の CSS）とインラインの値で入る。\n * - `container` は高さが決まっていて、表示はその中でスクロールする（`.xs-body` がスクロール枠）。\n */\n\n:where(:root, :host) {\n  /* 文字の色（塗りの無いセル・自動の文字・図形の色の省略） */\n  --xs-fg: #1a1a1a;\n  /* 面の地（タブ・ボタン） */\n  --xs-bg: #ffffff;\n  /* セルの地。白の塗り（「紙」）と、はみ出す文字が罫線を覆う色・二重線の隙間もこれ。書式のために別の変数を増やさない */\n  --xs-cell-bg: #f5f6f8;\n  /* 列記号・行番号の地。図の破線枠・図形の塗りの省略の色の元 */\n  --xs-head-bg: #eef0f4;\n  /* グリッド線（セルの右と下） */\n  --xs-grid: #e5e8ee;\n  /* 見出し行の下の線・タブとボタンの枠 */\n  --xs-border: #d0d0d0;\n  /* 破線枠・文字の枠の線 */\n  --xs-muted: #6b7280;\n  /* 行番号・空の表示の文字 */\n  --xs-faint: #8a92a2;\n  /* 選ばれているタブ・ボタンの強調 */\n  --xs-accent: #3b82f6;\n  /* 通知・エラーの帯 */\n  --xs-warn: #b45309;\n  --xs-warn-bg: #fef3c7;\n  /* 自動・黒の罫線の墨色（ダークは明るい灰色にする） */\n  --xs-border-ink: #000000;\n  /* 図の下地（ダーク表示だけ白にする）。-display は下地の svg の表示（明るい表示は none＝描画・レイアウトの負荷を省く） */\n  --xs-backdrop: transparent;\n  --xs-backdrop-display: none;\n}\n\n/* --- 表示の根 ---\n   `renderSheet` は `div.xs-sheet`、`renderWorkbook` は `div.xs-workbook`（タブ列＋`div.xs-sheet`）を container の子に足す。\n   どちらも container の高さいっぱいに広がり、表はスクロール枠（.xs-body）の中で動く。 */\n.xs-sheet,\n.xs-workbook {\n  box-sizing: border-box;\n  display: flex;\n  flex-direction: column;\n  height: 100%;\n  min-height: 0;\n  color: var(--xs-fg);\n  font-family: system-ui, sans-serif;\n}\n\n.xs-workbook > .xs-sheet {\n  flex: 1 1 0;\n  height: auto;\n}\n\n/* 本文枠（.xs-body）が仮想化の窓。描き直しでスクロール位置が動かないようスクロールアンカリングを切る。\n   面の色は、文字がはみ出す部分の罫線を覆う色（.xs-text）と必ず同じ（--xs-cell-bg）。 */\n.xs-body {\n  flex: 1 1 auto;\n  min-height: 0;\n  overflow: auto;\n  overflow-anchor: none;\n  background: var(--xs-cell-bg);\n}\n\n/* 図の通知（1 行）・開けないファイル・シートを開けないときの帯。通知は無ければ hidden。 */\n.xs-notice,\n.xs-banner {\n  flex: 0 0 auto;\n  margin-bottom: 12px;\n  padding: 8px 12px;\n  border-radius: 6px;\n  background: var(--xs-warn-bg);\n  color: var(--xs-warn);\n}\n\n.xs-notice[hidden] {\n  display: none;\n}\n\n/* シートが 0 枚・セルも図も無いシート */\n.xs-empty {\n  padding: 8px 0;\n  color: var(--xs-faint);\n  font-style: italic;\n}\n\n/* --- 表 ---\n   表は Excel の寸法（core の geometry）で固定する。\n   - 列幅は <colgroup> の <col width>、行の高さは <tr height> とセル内側の要素（.xs-cell）の height で固定する。\n   - 罫線は右と下だけ（border-collapse: separate・border-spacing: 0）。セルの寸法が罫線を含んだ整数 px になり、\n     行の位置が geometry の rowTop と 1px もずれない。td の左右 padding は 0（列が 6px など狭くても幅が崩れない）。\n   - 縦位置は下揃え（Excel の既定）。セルは折り返さず、はみ出しは内側の要素で隠す。\n   - 書式ごとの規則（`.xs-sheet[data-xs-fmt] td.xs-xf-<番号>`）は表示の根の `<style>` にあり、ここの規則より詳細度が強い。 */\n.xs-grid {\n  width: auto;\n  table-layout: fixed;\n  border-collapse: separate;\n  border-spacing: 0;\n  font-size: 13px;\n}\n\n.xs-grid th,\n.xs-grid td {\n  box-sizing: border-box;\n  padding: 0;\n  border: 0 solid var(--xs-grid);\n  border-width: 0 1px 1px 0;\n  text-align: left;\n  vertical-align: bottom;\n  white-space: nowrap;\n}\n\n.xs-grid th {\n  overflow: hidden;\n  background: var(--xs-head-bg);\n}\n\n/* td は overflow: visible（文字がはみ出す入れ物 .is-spill が右の空セルの上へ出られるように）。切るのは内側の要素。 */\n.xs-grid td {\n  overflow: visible;\n}\n\n/* 列記号の見出し行。高さは src/render/grid.ts の GRID_HEADER_HEIGHT と同じ値。 */\n.xs-grid thead th {\n  height: 24px;\n  line-height: 22px;\n  text-align: center;\n  vertical-align: middle;\n  border-bottom-color: var(--xs-border);\n}\n\n/* 行番号の列（左上の角を含む） */\n.xs-grid .xs-rownum {\n  position: sticky;\n  left: 0;\n  color: var(--xs-faint);\n  text-align: right;\n  background: var(--xs-head-bg);\n}\n\n/* スクロール中も見出し行・行番号の列を固定し、対応関係を見失わないようにする。\n   重ね順：セル・罫線の層・図の層（z 0）＜ 行番号（z 1）＜ 見出し行（z 2）＜ 左上の角（z 3）。 */\n.xs-body thead th {\n  position: sticky;\n  top: 0;\n  z-index: 2;\n}\n\n.xs-body tbody .xs-rownum {\n  z-index: 1;\n}\n\n.xs-body thead th.xs-rownum {\n  z-index: 3;\n}\n\n/* セル（と行番号）の中身。高さは JS が行高から出して指定する（行の高さを固定し、内容で伸ばさない）。 */\n.xs-cell {\n  display: flex;\n  align-items: flex-end;\n  box-sizing: border-box;\n  padding: 0 3px 1px;\n  overflow: hidden;\n  white-space: nowrap;\n}\n\n/* 文字のはみ出し：幅は JS が右の空セルの列幅の合計で指定する。行番号・見出し（z-index 1 以上）の下に入るよう z-index 0。\n   文字を包む span だけがシートの面の色を持ち、文字の範囲の縦罫線を隠す（文字の届かない空セルの罫線は見える）。 */\n.xs-cell.is-spill {\n  position: relative;\n  z-index: 0;\n}\n\n.xs-cell.is-spill .xs-text {\n  display: flex;\n  align-items: flex-end;\n  align-self: stretch;\n  margin: 0 -3px -1px -3px;\n  padding: 0 3px 1px 3px;\n  background: var(--xs-cell-bg);\n}\n\n/* 白以外の塗りが絡むはみ出し（grid-table.ts が is-over-fill を付ける）。\n   文字の下に面の色を敷かない（敷くと右の空セルの塗りを隠す）。各セルは自分の塗りを保つ。\n   「自動」色の文字が塗りの状態の違う列をまたぐときは、grid-table.ts（inkSpill）が span に列ごとの区間の\n   linear-gradient＋background-clip: text＋color: transparent をインラインで足す（塗りの上は黒、塗りなしは --xs-fg）。 */\n.xs-cell.is-spill.is-over-fill .xs-text {\n  background: transparent;\n}\n\n/* 型ごとの既定の配置（Excel の「標準」）：数値・日付は右、真偽・エラーは中央、文字列は左。\n   safe：入れ物より文字が広いとき（フォントの差で ### が溢れる等）は左から並べて右を切る（左が切れて途中の桁だけ見えるのを避ける）。\n   safe に対応しないブラウザは 1 行目の指定になる。 */\n.xs-cell.is-right {\n  justify-content: flex-end;\n  justify-content: safe flex-end;\n}\n\n.xs-cell.is-center {\n  justify-content: center;\n  justify-content: safe center;\n}\n\n.xs-grid .xs-rownum .xs-cell {\n  align-items: center;\n  justify-content: flex-end;\n  padding: 0 6px;\n}\n\n/* 非表示の行・列は畳む（Excel どおり。見出し・行番号も一緒に畳む）。 */\n.xs-grid .xs-collapsed {\n  visibility: collapse;\n}\n\n/* 仮想化で描かない行の高さを埋める行 */\n.xs-grid .xs-spacer td {\n  padding: 0;\n  border: 0;\n}\n\n/* --- 図の層 ---\n   スクロール枠（.xs-body）の中の wrapper（position: relative）に、表・罫線の層・図の層を並べる。\n   層の原点は A1 の左上（JS が left/top に行番号の列の幅・見出し行の高さを指定する）。\n   z-index: 0 の位置指定のため、セルの背景・文字の上、列見出し（sticky・z 2）・行番号（sticky・z 1）の下を通る。\n   図にクリック操作は持たない（pointer-events: none で、セルの選択・コピーを妨げない）。 */\n.xs-wrap {\n  position: relative;\n}\n\n.xs-layer {\n  position: absolute;\n  z-index: 0;\n  /* 重ね順の文脈を層が作る（z-index: 0 の位置指定と同じだが明示）。下地（z-index: -1）は層の中で全図の下・セルの上に描かれる。 */\n  isolation: isolate;\n  overflow: hidden;\n  pointer-events: none;\n}\n\n/* 図の要素には transform・opacity・z-index を付けない（重ね順の文脈を作らない。回転は内側の要素・下地の svg だけ）。 */\n.xs-figure {\n  position: absolute;\n}\n\n/* 罫線の層（border-layer.ts）。表の直後・図の層の直前の svg。\n   z-index: 0 の位置指定で、DOM の順にセル（はみ出しの文字を含む）の上・図の層の下・sticky の見出しと行番号（z 1・2）の下を通る。\n   原点・大きさは JS が style に指定する（描いた線の外接矩形）。クリック操作は持たない（pointer-events: none）。\n   色は変数 --xs-border-ink（自動・黒の罫線の墨色）と --xs-cell-bg（二重線の隙間）だけ。\n   明示色は JS が CSSOM で stroke に入れる（要素の style は下の規則に勝つ）。テーマの切り替えで描き直さない。 */\n.xs-borders {\n  position: absolute;\n  z-index: 0;\n  overflow: visible;\n  pointer-events: none;\n}\n\n.xs-border-path {\n  fill: none;\n}\n\n.xs-border-path.is-ink {\n  stroke: var(--xs-border-ink);\n}\n\n.xs-border-path.is-gap {\n  stroke: var(--xs-cell-bg);\n}\n\n/* ダーク表示の図の下地（図形・線・代替の文字枠の直前の兄弟 svg。image-layer.ts）。\n   色は変数 --xs-backdrop だけ（明るい表示は透明・ダークは白。要素の style に色を入れないため、テーマの切り替えで描き直さない）。\n   z-index: -1 で、層の中の全図（画像・線・図形・文字）の下に描く。縁取り（太い白い線）が svg の外へ出ても切らない。\n   明るい表示は透明の下地を描く意味が無いため --xs-backdrop-display を none にする（形のある図形 1,000 個の計測で、透明のままでも描画の負荷が増えたため）。 */\n.xs-backdrop {\n  position: absolute;\n  z-index: -1;\n  display: var(--xs-backdrop-display, block);\n  overflow: visible;\n}\n\n.xs-backdrop-part {\n  fill: none;\n  stroke: none;\n}\n\n.xs-backdrop-part.is-fill {\n  fill: var(--xs-backdrop);\n}\n\n.xs-backdrop-part.is-stroke {\n  stroke: var(--xs-backdrop);\n}\n\n/* 切り抜き・回転・反転は無視し、アンカーの枠いっぱいに伸ばす（既知の差異）。 */\n.xs-img {\n  display: block;\n  width: 100%;\n  height: 100%;\n}\n\n/* 画像以外・描けない画像・壊れた画像の種類名入りの破線枠。色は変数だけ（ダーク表示でも読める）。 */\n.xs-frame {\n  box-sizing: border-box;\n  width: 100%;\n  height: 100%;\n  padding: 2px 4px;\n  overflow: hidden;\n  border: 1px dashed var(--xs-muted);\n  background: color-mix(in srgb, var(--xs-head-bg) 75%, transparent);\n  color: var(--xs-fg);\n  font-size: 12px;\n  line-height: 1.3;\n  white-space: nowrap;\n  text-overflow: ellipsis;\n}\n\n/* 文字のある図形・テキストボックス。段落（div）・ラン（span）は JS が文字の大きさ・太さ等を style に入れる。\n   縦位置は flex（column の主軸＝ブロック方向）、揃えは text-align の論理値のため、縦書きでもそのまま合う。\n   枠線・塗りは Excel の有無に合わせてクラスで切り替える（色は変数だけ）。枠線は outline（余白・文字の位置を動かさない）。 */\n.xs-shape-text {\n  box-sizing: border-box;\n  display: flex;\n  flex-direction: column;\n  overflow: hidden;\n  color: var(--xs-fg);\n  line-height: 1.2;\n  overflow-wrap: break-word;\n}\n\n.xs-shape-text.has-outline {\n  outline: 1px dashed var(--xs-muted);\n  outline-offset: -1px;\n}\n\n.xs-shape-text.has-fill {\n  background: color-mix(in srgb, var(--xs-head-bg) 75%, transparent);\n}\n\n/* 線・矢印（線 1 本につき svg 1 つ）。外接矩形が矢じり・太さを含むため overflow は既定のまま。\n   色は Excel の色を要素の style（fill・stroke）に入れる。色が無い（省略）ときだけ currentColor 経由で --xs-fg。 */\n.xs-line {\n  display: block;\n  color: var(--xs-fg);\n}\n\n/* 図形の形（図形 1 つにつき svg 1 つ＋文字の枠。image-layer.ts）。\n   色は Excel の色を要素の style（fill・stroke）に入れる。以下は色が無い（省略）ときだけ効く（枠線＝--xs-fg、塗り＝--xs-head-bg の 75%）。\n   svg は外接矩形いっぱい。miter の角が外接矩形から少し出ても切らない（overflow: visible。層が外側で切る）。 */\n.xs-shape-svg {\n  position: absolute;\n  top: 0;\n  left: 0;\n  display: block;\n  overflow: visible;\n  color: var(--xs-fg);\n}\n\n.xs-shape-path {\n  fill: none;\n  stroke: none;\n}\n\n.xs-shape-path.has-fill {\n  fill: color-mix(in srgb, var(--xs-head-bg) 75%, transparent);\n}\n\n.xs-shape-path.has-stroke {\n  stroke: currentColor;\n}\n\n/* 面の濃淡（path の fill が darken・darkenLess・lighten・lightenLess）：塗りに文字色（暗く）・面の地（明るく）の薄い膜を重ねた色。 */\n.xs-shape-path.has-fill.is-darken {\n  fill: color-mix(\n    in srgb,\n    color-mix(in srgb, var(--xs-fg) 15%, var(--xs-head-bg)) 75%,\n    transparent\n  );\n}\n\n.xs-shape-path.has-fill.is-darken-less {\n  fill: color-mix(in srgb, color-mix(in srgb, var(--xs-fg) 8%, var(--xs-head-bg)) 75%, transparent);\n}\n\n.xs-shape-path.has-fill.is-lighten {\n  fill: color-mix(\n    in srgb,\n    color-mix(in srgb, var(--xs-bg) 30%, var(--xs-head-bg)) 75%,\n    transparent\n  );\n}\n\n.xs-shape-path.has-fill.is-lighten-less {\n  fill: color-mix(\n    in srgb,\n    color-mix(in srgb, var(--xs-bg) 15%, var(--xs-head-bg)) 75%,\n    transparent\n  );\n}\n\n/* 回転前の矩形（JS が left/top/width/height と rotate() を指定）。文字は図形の矩形で切る。\n   文字領域（.xs-shape-text）は box の中に絶対配置し、領域では切らない（楕円の内接矩形からはみ出した文字も、矩形の中なら見える）。 */\n.xs-shape-box {\n  position: absolute;\n  overflow: hidden;\n  color: var(--xs-fg);\n}\n\n.xs-shape-box > .xs-shape-text {\n  position: absolute;\n  overflow: visible;\n}\n\n/* --- シートのタブ（renderWorkbook）---\n   Excel のシート見出しのように 1 行に並べ、入りきらない分は横スクロール。\n   表示範囲の左右に ◀ ▶ ボタンを置く（JS が端・あふれに応じて disabled・hidden にする）。 */\n.xs-sheet-strip {\n  display: flex;\n  flex: 0 0 auto;\n  align-items: stretch;\n  gap: 4px;\n  margin-bottom: 16px;\n}\n\n.xs-sheet-tabs {\n  position: relative; /* タブの offsetLeft をスクロール領域の中の座標にする */\n  display: flex;\n  flex: 1 1 auto;\n  gap: 8px;\n  min-width: 0;\n  overflow-x: auto;\n  overflow-y: hidden;\n  scrollbar-width: none;\n}\n\n.xs-sheet-tabs::-webkit-scrollbar {\n  display: none;\n}\n\n.xs-sheet-tab {\n  flex-shrink: 0;\n  max-width: 240px;\n  padding: 6px 12px;\n  overflow: hidden;\n  border: 1px solid var(--xs-border);\n  border-radius: 6px;\n  background: var(--xs-bg);\n  color: var(--xs-fg);\n  cursor: pointer;\n  font: inherit;\n  font-size: 13px;\n  white-space: nowrap;\n  text-overflow: ellipsis;\n}\n\n.xs-sheet-tab.is-active {\n  border-color: var(--xs-accent);\n  background: var(--xs-accent);\n  color: #fff;\n}\n\n/* 非表示のシートのタブ（⌀ 付き）は区別できるよう斜体・破線にする。 */\n.xs-sheet-tab.is-hidden-sheet {\n  font-style: italic;\n  border-style: dashed;\n}\n\n.xs-sheet-nav {\n  flex: 0 0 auto;\n  padding: 0 8px;\n  border: 1px solid var(--xs-border);\n  border-radius: 6px;\n  background: var(--xs-bg);\n  color: var(--xs-fg);\n  cursor: pointer;\n  font-size: 11px;\n  line-height: 1;\n}\n\n.xs-sheet-nav:hover:not(:disabled) {\n  border-color: var(--xs-accent);\n  color: var(--xs-accent);\n}\n\n.xs-sheet-nav:disabled {\n  opacity: 0.35;\n  cursor: default;\n}\n\n.xs-sheet-nav[hidden] {\n  display: none;\n}\n\n/* --- ダーク表示（M6-2・ADR-0003 §2）---\n   `data-xs-theme=\"dark\"` を表示の根か、その祖先（`<html>` など）に付ける。`view` の theme は表示の根 `.xs-viewer` に付ける。\n   `auto` は OS・ブラウザの設定（prefers-color-scheme）に従う。値を変えたいときは、同じ要素か内側で `--xs-*` を上書きする。 */\n[data-xs-theme='dark'] {\n  color-scheme: dark;\n  --xs-fg: #e6e6e6;\n  --xs-bg: #1e1e1e;\n  --xs-cell-bg: #0f1115;\n  --xs-head-bg: #1f232c;\n  --xs-grid: #232733;\n  --xs-border: #3a3a3a;\n  --xs-muted: #9ca3af;\n  --xs-faint: #6b7386;\n  --xs-accent: #60a5fa;\n  --xs-warn: #fbbf24;\n  --xs-warn-bg: rgba(251, 191, 36, 0.14);\n  --xs-border-ink: #d4d4d4;\n  --xs-backdrop: #ffffff;\n  --xs-backdrop-display: block;\n}\n\n@media (prefers-color-scheme: dark) {\n  [data-xs-theme='auto'] {\n    color-scheme: dark;\n    --xs-fg: #e6e6e6;\n    --xs-bg: #1e1e1e;\n    --xs-cell-bg: #0f1115;\n    --xs-head-bg: #1f232c;\n    --xs-grid: #232733;\n    --xs-border: #3a3a3a;\n    --xs-muted: #9ca3af;\n    --xs-faint: #6b7386;\n    --xs-accent: #60a5fa;\n    --xs-warn: #fbbf24;\n    --xs-warn-bg: rgba(251, 191, 36, 0.14);\n    --xs-border-ink: #d4d4d4;\n    --xs-backdrop: #ffffff;\n    --xs-backdrop-display: block;\n  }\n}\n\n/* --- `view` の表示の根（M6-2）---\n   container の高さいっぱいに広がる。container の高さが 0 のときは `view` が既定の高さ（480px）を CSSOM で付ける。 */\n.xs-viewer {\n  box-sizing: border-box;\n  height: 100%;\n  min-height: 0;\n}\n";
// src/facade/style.ts
/**
* `view` の表示の CSS・テーマ・高さ（M6-2）。
* - CSS：`style.css` の文面を `<style data-xsheet-style>` で 1 回だけ入れる。`<link>`・import で読み込み済みなら入れない
*   （試しの要素に `.xs-workbook` の規則（display: flex）が効くかで見る。変数は継承され、使う側が `:root` に書くこともあるため、変数では見ない）。
*   container が Shadow DOM の中なら、その ShadowRoot に入れる（文書の CSS は Shadow DOM の中に届かない）。
*   書式の `<style>` と同じく nonce を付ける（CSP の要件は今と変わらない）。
* - テーマ：表示の根に `data-xs-theme` を付ける。ダークの値は style.css が持つ。
* - 高さ：container の高さが 0 のときだけ、表示の根に既定の高さを付けて console.warn を 1 回出す（P5）。
*   高さは表示の根（`.xs-viewer`）を隠して測る（前の表示の高さを container の高さと取り違えない）。
*   container が描かれていない（display: none の中）ときは、描かれたときに ResizeObserver で測る。
*   高さは CSSOM（`el.style.height`）で付ける（style 属性の文字列は CSP の unsafe-inline が要るため使わない）。
*/
/** container の高さが 0 のときに付ける高さ */
var DEFAULT_HEIGHT = "480px";
/** 自動で入れた `<style>` の印 */
var STYLE_MARK = "data-xsheet-style";
var warnedHeight = false;
/** CSS を入れる先：container が Shadow DOM の中ならその ShadowRoot、そうでなければ文書の head */
function styleHost(container) {
	const root = container.getRootNode();
	if (root.nodeType === 11 && "host" in root) return root;
	const doc = container.ownerDocument;
	return doc.head ?? doc.documentElement;
}
/** 表示の CSS が container に効いているか（試しの要素を足して、すぐ外す。同期なので描画されない） */
function styleApplied(container) {
	const view = container.ownerDocument.defaultView;
	if (!view) return false;
	const probe = container.ownerDocument.createElement("div");
	probe.className = "xs-workbook";
	container.append(probe);
	const display = view.getComputedStyle(probe).display;
	probe.remove();
	return display === "flex";
}
/** 表示の CSS が無ければ入れる（読み込み済み・入れ済みなら何もしない） */
function ensureStyle(container, nonce) {
	const host = styleHost(container);
	if (host.querySelector(`style[${STYLE_MARK}]`) !== null) return;
	if (host === container.ownerDocument.head || host === container.ownerDocument.documentElement) {
		if (styleApplied(container)) return;
	}
	const style = container.ownerDocument.createElement("style");
	style.setAttribute(STYLE_MARK, "");
	if (nonce !== void 0) style.setAttribute("nonce", nonce);
	style.textContent = style_default;
	host.append(style);
}
/** テーマを表示の根に付ける（`light` は何も付けず、使う側の `--xs-*` の上書きがそのまま効く） */
function applyTheme(root, theme) {
	if (theme === "dark" || theme === "auto") root.setAttribute("data-xs-theme", theme);
}
/** container の高さ（中の表示の根を隠して測る） */
function containerHeight(container) {
	const roots = [...container.children].filter((el) => el.classList.contains("xs-viewer") && "style" in el);
	const saved = roots.map((el) => el.style.display);
	roots.forEach((el) => el.style.display = "none");
	const height = container.clientHeight;
	roots.forEach((el, i) => el.style.display = saved[i]);
	return height;
}
/** 高さが無いときの既定の高さを付ける */
function useDefaultHeight(root) {
	root.style.height = DEFAULT_HEIGHT;
	if (!warnedHeight) {
		warnedHeight = true;
		console.warn(`XSheet: the container has no height, so the viewer uses ${DEFAULT_HEIGHT}. Give the container a height (or pass the height option) to change it`);
	}
}
/**
* 表示の根の高さを決める。
* - 数値は px、文字列は CSS の長さとしてそのまま付ける。
* - `auto`（既定）は container の高さが 0 のときだけ {@link DEFAULT_HEIGHT} を付ける（それ以外は style.css の 100%）。
*   container が描かれていないときは、描かれたときに決める。
* @returns 後始末（描かれるのを待つ ResizeObserver を外す）
*/
function applyHeight(container, root, height) {
	const none = () => void 0;
	if (typeof height === "number") {
		root.style.height = `${height}px`;
		return none;
	}
	if (height !== void 0 && height !== "auto") {
		root.style.height = height;
		return none;
	}
	const decide = () => {
		if (containerHeight(container) === 0) useDefaultHeight(root);
	};
	const Observer = container.ownerDocument.defaultView?.ResizeObserver;
	if (container.getClientRects().length > 0 || Observer === void 0) {
		decide();
		return none;
	}
	const observer = new Observer(() => {
		if (container.getClientRects().length === 0) return;
		observer.disconnect();
		decide();
	});
	observer.observe(container);
	return () => observer.disconnect();
}
// src/facade/view.ts
/**
* 1 行で使える入口 `view`・`read`（M6-2）。core（`openXlsx`）と render（`renderWorkbook`）の上の薄い層で、
* 写したファイル（`src/core/`・`src/render/`）の作りには立ち入らない。
* - `view`：xlsx をタブ付きで出す。表示の根 `div.xs-viewer` を container の子に足し、その中に `renderWorkbook` の表示を置く。
*   言語・テーマ・CSS・高さを語で選べる。開けないファイルは今までどおり例外にせず、帯（`div.xs-banner`）に文言を出す。
* - `read`：DOM なしで開く（`openXlsx` と同じ。受け取れる形を `view` と揃える）。
*/
var TABLES = {
	ja: MESSAGES_JA,
	en: MESSAGES_EN,
	ko: MESSAGES_KO
};
/** 言語の名前（`ko-KR` など）から表を選ぶ。合わなければ null */
function tableFor(tag) {
	const primary = (tag ?? "").trim().toLowerCase().split(/[-_]/)[0];
	return primary === "ja" || primary === "en" || primary === "ko" ? TABLES[primary] : null;
}
/**
* 使う文言の表（`lang` で選んだ表に `messages` を重ねる）。
* `auto` は `<html lang>` があればそれだけで決め（合わなければ英語）、無いときだけブラウザの言語を見る。
* 型を通らない呼び出し（`<script>` から `'ko-KR'`・`'zh'` など）も、名前から選ぶか `auto` と同じにする。
*/
function pickMessages(doc, options) {
	const lang = options.lang ?? "auto";
	let base = lang === "auto" || typeof lang !== "string" ? null : tableFor(lang);
	if (base === null) {
		const pageLang = doc.documentElement.lang.trim();
		const nav = doc.defaultView?.navigator;
		base = pageLang !== "" ? tableFor(pageLang) ?? MESSAGES_EN : tableFor(nav?.languages?.[0]) ?? tableFor(nav?.language) ?? MESSAGES_EN;
	}
	return resolveMessages(options.messages, base);
}
/** 同じ container の今の表示（新しい `view` が前の表示を外す。読み込み中に次の `view` が来たら、前の読み込みは捨てる） */
var current = /* @__PURE__ */ new WeakMap();
/** 捨てた読み込みの戻り値（表示は無い） */
var DISCARDED = Object.freeze({
	sheets: Object.freeze([]),
	showSheet: () => Promise.reject(/* @__PURE__ */ new RangeError("XSheet: the viewer was replaced or destroyed")),
	destroy: () => void 0
});
/** 要素か CSS セレクタを要素にする */
function resolveTarget(target) {
	if (typeof target !== "string") {
		if (typeof target !== "object" || target === null || target.nodeType !== 1) throw new TypeError("XSheet: the target must be an element or a CSS selector");
		return target;
	}
	const el = document.querySelector(target);
	if (el === null) throw new TypeError(`XSheet: no element matches the selector ${target}`);
	return el;
}
/**
* xlsx をタブ付きで表示する。
* 同じ container に続けて呼ぶと、前の表示を外してから出す（読み込み中の前の呼び出しは、表示の無い Viewer を返す）。
* @param target 表示を置く要素か CSS セレクタ（子として `div.xs-viewer` を足す）
* @param source File・Blob・ArrayBuffer・型付き配列・Response、またはその Promise（`fetch(url)` をそのまま渡せる）
* @throws TypeError 要素が見つからない・受けない形の source・失敗した Response
*/
async function view(target, source, options = {}) {
	const container = resolveTarget(target);
	const token = {};
	current.get(container)?.viewer?.destroy();
	current.set(container, {
		token,
		viewer: null
	});
	const isCurrent = () => current.get(container)?.token === token;
	let bytes;
	try {
		bytes = await toSourceBytes(source);
	} catch (e) {
		if (isCurrent()) current.delete(container);
		throw e;
	}
	if (!isCurrent()) return DISCARDED;
	if (options.css !== "none") ensureStyle(container, options.nonce);
	const root = container.ownerDocument.createElement("div");
	root.className = "xs-viewer";
	applyTheme(root, options.theme);
	const disposeHeight = applyHeight(container, root, options.height);
	container.append(root);
	let workbook;
	try {
		workbook = await renderWorkbook(root, bytes, {
			messages: pickMessages(container.ownerDocument, options),
			drawings: options.drawings,
			nonce: options.nonce,
			limits: options.limits
		});
	} catch (e) {
		disposeHeight();
		root.remove();
		if (isCurrent()) current.delete(container);
		throw e;
	}
	const wb = workbook;
	if (!isCurrent()) {
		disposeHeight();
		wb.destroy();
		root.remove();
		return DISCARDED;
	}
	const destroy = () => {
		disposeHeight();
		wb.destroy();
		root.remove();
		if (current.get(container)?.viewer === viewer) current.delete(container);
	};
	const viewer = {
		sheets: wb.sheets,
		showSheet(sheet) {
			const info = wb.sheets.find((s) => (typeof sheet === "number" ? s.index : s.name) === sheet);
			if (info === void 0) return Promise.reject(/* @__PURE__ */ new RangeError("XSheet: no such sheet"));
			return wb.showSheet(info.index);
		},
		destroy
	};
	current.set(container, {
		token,
		viewer
	});
	return viewer;
}
/**
* xlsx を DOM なしで開く（`openXlsx` と同じ。閉じるのは呼び出し側の `close()`）。
* @throws TypeError 受けない形の source・失敗した Response
* @throws XlsxError 開けないファイル（`kind` で種類が分かる）
*/
async function read(source, options = {}) {
	return openXlsx(await toSourceBytes(source), { limits: options.limits });
}
export { DEFAULT_LIMITS, MESSAGES_EN, MESSAGES_JA, MESSAGES_KO, XlsxError, colName, colWidthToPx, createGeometry, defaultColWidthToPx, fontDefaultRowPx, isPaperFill, lookupMdw, openXlsx, ptToPx, read, renderSheet, renderWorkbook, view };
