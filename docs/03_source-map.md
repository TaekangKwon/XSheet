# 03. 元のコードとの対応 — InterHub から写すもの・直すこと

| 項目 | 値 |
| --- | --- |
| 文書 | XSheet 立ち上げ資料 03 |
| 元 | `C:\GitHub\InterHub` main `943e4f0`（以下のパスは `src_V1/src/` 基準。`tools/` だけ `src_V1/` 基準） |
| 更新日 | 2026-10-28 |
| 関連 | [02_architecture.md](02_architecture.md)・[04_roadmap.md](04_roadmap.md) |

## TL;DR
- core は `lib/xlsx/` をほぼそのまま写す。geom は描画側から core と並ぶ層へ移す。render は `lib/viewers/xlsx-*.ts` を写し、`spreadsheet.ts` から .xlsx の部分だけを抜き出す。
- 直すのは 6 点：層の逆転・`t()`・`ih-` のクラスと変数・モジュール変数の状態・InterHub の共用部品・型だけの import。
- InterHub 側は読むだけで変更しない。写した時点のコミット（`943e4f0`）を新リポジトリの最初のコミットに記録する。

---

## 1. core（`lib/xlsx/` → `src/core/`）

| ファイル | 行数 | 役割 | 写すときの注意 |
| --- | --- | --- | --- |
| `index.ts` | 237 | `openXlsx`・`XlsxBook` の組み立て・テーマを 1 回だけ読む | 公開の範囲を決める（[02](02_architecture.md) §2） |
| `types.ts` | 599 | 公開型 | `../viewers/xlsx-geom` の import を `../geom` へ |
| `limits.ts` | 196 | 上限と既定値 | 根拠のコメントの InterHub 固有の語（「非破壊の原則」など）を一般の語に |
| `open.ts` | 149 | 先頭の印・jszip・展開の大きさの上限 | そのまま |
| `workbook.ts` | 141 | workbook.xml・rels・テーマのパス | そのまま |
| `sheet.ts` | 773 | 行の索引・`getRows`・結合・寸法・罫線と書式の番号・範囲 | そのまま |
| `shared-strings.ts` | 68 | 共有文字列 | そのまま |
| `styles.ts` | 73 | styles.xml の分割 | そのまま |
| `format.ts` | 161 | 数値の表示形式（ssf） | そのまま |
| `xml.ts` | 196 | 正規表現の XML 小道具 | そのまま |
| `geometry.ts` | 231 | 列幅・行高 → px の座標 | そのまま |
| `font-metrics.ts` | 210 | 既定フォントの最大桁幅（MDW）の表 | そのまま |
| `borders.ts` | 143 | 罫線の表 | そのまま |
| `cell-color.ts` | 170 | セルの色（rgb・theme・indexed・tint） | そのまま |
| `cell-format.ts` | 325 | 書式の表（fonts・fills・cellXfs） | そのまま |
| `color.ts` | 420 | 色の要素と変換 | そのまま |
| `theme.ts` | 102 | テーマの 12 色 | そのまま |
| `drawing.ts` | 608 | drawing の読み取り・アンカー | そのまま |
| `drawing-text.ts`・`drawing-shape.ts`・`drawing-group.ts`・`drawing-color.ts` | 330・322・92・183 | 図形の文字・形と線・グループ・色 | そのまま |
| `drawing-geom.ts` | 283 | prstGeom・avLst・custGeom | `../viewers/xlsx-geom` の import を `../geom` へ |
| `image.ts` | 254 | 画像の読み取り・形式の判定・予算 | そのまま |
| `make-xlsx.ts`・`make-image.ts` | 811・197 | 合成の xlsx・画像を作る（テストだけが使う） | `src/testing/` へ。公開しない |

## 2. geom（→ `src/geom/`・`tools/geom/`）

| 元 | 先 | 注意 |
| --- | --- | --- |
| `lib/viewers/xlsx-geom.ts`（522 行） | `src/geom/evaluate.ts` | 17 演算子・path 6 命令・文字領域 |
| `lib/viewers/xlsx-preset-geom.json`（171 KB・176 種） | `src/geom/preset-geom.json` | render から動的 import（図形が無いシートでは読まない） |
| `tools/geom/presetShapeDefinitions.xml`・`build-preset-geom.mjs`・`.d.mts`・`.gitattributes` | `tools/geom/` | `npm run geom:build` で JSON を作り直せること |
| `tools/geom/NOTICE` | ルートの `NOTICE` | Ecma の著作権表示。配布物に必ず含める |

## 3. render（`lib/viewers/` → `src/render/`）

| ファイル | 行数 | 役割 | 注意 |
| --- | --- | --- | --- |
| `xlsx-grid.ts` | 659 | 範囲・行と列の仮想化・はみ出し・`###`・余白の書式 | `ih-` → `xv-` |
| `xlsx-borders.ts` | 465 | 罫線の線分・辺の優先・path | そのまま |
| `xlsx-border-layer.ts` | 128 | 罫線の SVG 層 | `ih-` → `xv-` |
| `xlsx-format-css.ts` | 165 | 書式のクラスとブック単位の CSS | `ih-xf-`・`data-ih-xlsx-fmt`・`.ih-spreadsheet` → `xv-` 系。`nonce` を付けられるように（M3） |
| `xlsx-drawings.ts` | 327 | アンカー → px・グループの写像・見えている図 | そのまま |
| `xlsx-image-store.ts` | 147 | 画像の blob URL と後始末 | そのまま |
| `xlsx-image-layer.ts` | 778 | 図の DOM 層・破線枠・下地 | `t()` → `messages`、`ih-` → `xv-` |
| `xlsx-shapes.ts` | 912 | 図形と線の SVG・矢じり・外接矩形 | そのまま |
| `spreadsheet.ts`（1,076 行のうち .xlsx の部分） | — | 表と層の組み立て・スクロール・後始末 | `renderXlsxView`（551 行〜）・`showSheet`（813 行〜）・`draw`（707 行〜）・`buildGridTable`（409 行〜）・`dispose`（907 行〜）を `renderSheet` へ。`.xls`（SheetJS）の経路・`fitTableBodyToWindow`・InterHub のバナーとツールバーは写さない |
| `styles/global.css` の `.ih-spreadsheet`・`.ih-xlsx-*` | 31〜1442 行に散在（74 か所） | 表・層・図・罫線・下地の規則と変数 | `src/style.css` へ。変数は [02](02_architecture.md) §6 |
| `locales/ja.ts`・`en.ts` の `spreadsheet.*` | 22 キー（ja.ts 275〜302 行） | 文言 | `src/render/messages.ts` の既定値に |

## 4. 直す 6 点

| # | 直すこと | 場所 | 方法 |
| --- | --- | --- | --- |
| 1 | 層の逆転 | `drawing-geom.ts`・`types.ts` が `viewers/xlsx-geom` を import | geom を独立の層にする（§2） |
| 2 | InterHub の文言関数 | `spreadsheet.ts`・`xlsx-image-layer.ts` の `t()` | `RenderOptions.messages` を引数で回す |
| 3 | 接頭辞 | `ih-xlsx-*`・`ih-xf-N`・`data-ih-xlsx-fmt`・`--ih-*` | `xv-` 系へ一括で置き換え、テストの期待値も合わせる |
| 4 | モジュール変数の状態 | `spreadsheet.ts` の `disposeXlsxView`・`renderGeneration` | `SheetView` のインスタンスへ |
| 5 | InterHub の共用部品 | `table-fit.ts`（`fitTableBodyToWindow`）・`ih-data-view`・`ih-data-banner`・`ih-dv-rownum`・`ih-sheet-tabs` | render に要るものだけ `xv-` で作り直す。枠の高さは使う側の仕事 |
| 6 | 型だけの import | 全体 | `verbatimModuleSyntax: true` にして `import type` へ直す（xrag-Viewer はこの設定） |

## 5. テスト（写す）

| 元 | 件数 | 先 |
| --- | --- | --- |
| `lib/xlsx/*.test.ts` | 36 ファイル | `tests/core/` |
| `lib/viewers/xlsx-*.test.ts`・`spreadsheet.*.render.test.ts` | 26＋5 ファイル | `tests/render/`（`spreadsheet.*` は `renderSheet` を呼ぶ形に書き直す） |
| `lib/xlsx/__fixtures__/excel/` | 22 ファイル | `fixtures/excel/`（[05_quality.md](05_quality.md) §2） |

- テストは vitest＋jsdom。`node:fs` で fixtures を読むテストがある（Node で走る前提）。
- InterHub の版：vitest 3.2・jsdom 29・TypeScript 5.9・vite 6.4。新しい版（xrag は vitest 5・jsdom 30・TS 7・vite 8）へ上げるときの落ちは M1・M2 で直す。
