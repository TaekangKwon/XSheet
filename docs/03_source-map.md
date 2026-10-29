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

### M1-2 で写した結果（2026-10-28）

- 元はすべて InterHub `943e4f0` から取り出した。本体 24・テストの道具 2（`src/testing/`）・テスト 32（`tests/core/`）・fixtures 23（ルートの `fixtures/excel/`）。
  - fixtures の xlsx・JSON・`.ps1`・`.d.mts` は元の blob と一致（git のハッシュで確認）。`.mjs` の 5 本は Prettier で整えた分だけ違う。
  - `git archive` は `core.autocrlf` で改行を CRLF に変えて出す。原本のバイトが要るものは `git cat-file blob` で取り出す。
- 直したのは次だけ（Prettier の整形と import のパスを除くと、元との差はこれで全部）：
  - `drawing-geom.ts`：`GUIDE_ARITY` を `../geom/defs` から引いて出し直す（同じ表を 2 か所に持たない）。
  - `limits.ts`：`MAX_GEOM_GUIDES`・`MAX_GEOM_COMMANDS` を `../geom/defs` から引いて出し直す（「定数の置き場は limits.ts の 1 か所」の唯一の例外。冒頭の説明にも書いた）。
  - `color.ts` の `hslToRgb`：ESLint 10 の `no-useless-assignment` に当たる初期値（どの分岐でも代入し直す `let r = 0` など）を、型の宣言だけにした（動きは同じ）。
  - 冒頭の出典（`// 正: docs/...`）を `// 正: InterHub docs/...` に、`viewers/xlsx-geom.ts` を指すコメントを geom の `evaluate.ts` に直した。
- `types.ts` は geom を import していなかった（コメントで触れていただけ）。§1 の表の注意は不要だった。
- `import type`：元が `import type` で書かれており、`verbatimModuleSyntax` のままの型検査で直すところは無かった。
- テスト：`./x` → `../../src/core/x`、`./make-xlsx` → `../../src/testing/make-xlsx`、`../viewers/xlsx-geom` → `../../src/geom/evaluate`、`./__fixtures__/excel/` → `../../fixtures/excel/`。`vi.mock` のパスも同じく直した。
  - `tests/geom/preset-geom.test.ts` の `LINE_PRESETS` の照合を `it.todo` から戻した。
- **Node 固有の API の止め方を足した**（[ADR-0001](adr/0001-library-decision.md) 追記）。jszip の型が `@types/node` を引き込み、`tsconfig.core.json` では `process`・`Buffer` が通ってしまう（DOM は止まる）。ESLint で `src/core/**`・`src/geom/**` の Node の大域変数と `node:*`・組み込みモジュールの import を止める。
- `fixtures/excel/compare-ref-geoms.mjs` のハーネスは InterHub の描画を指したまま（M2 で向け直す。[05_quality.md](05_quality.md) §2）。

## 2. geom（→ `src/geom/`・`tools/geom/`）

| 元 | 先 | 注意 |
| --- | --- | --- |
| `lib/viewers/xlsx-geom.ts`（522 行） | `src/geom/evaluate.ts` | 17 演算子・path 6 命令・文字領域 |
| `lib/viewers/xlsx-preset-geom.json`（171 KB・176 種） | `src/geom/preset-geom.json` | render から動的 import（図形が無いシートでは読まない） |
| `tools/geom/presetShapeDefinitions.xml`・`build-preset-geom.mjs`・`.d.mts`・`.gitattributes` | `tools/geom/` | `npm run geom:build` で JSON を作り直せること |
| `tools/geom/NOTICE` | ルートの `NOTICE` | Ecma の著作権表示。配布物に必ず含める |

### M1-1 で写した結果（2026-10-28）

- 元はすべて `git show 943e4f0:<パス>` で取り出した（XML・NOTICE・JSON はバイト単位で一致を確認）。元の変換器と元の NOTICE からの再生成も、元の JSON とバイト単位で一致した。
- **`src/geom/defs.ts` を新しく作った**。評価器が core と render から引いていたものを geom に移した（直す 6 点の 1）。
  - `GUIDE_ARITY`（元は core の `drawing-geom.ts`）。M1-2 で `drawing-geom.ts` はここから引く。
  - `MAX_GEOM_GUIDES`・`MAX_GEOM_COMMANDS`（元は core の `limits.ts`）。M1-2 で `limits.ts` はここから引いて `DEFAULT_LIMITS` に入れる（「定数の置き場は limits.ts の 1 か所」の例外）。
  - `Pt`・`PathSeg`（元は render の `xlsx-shapes.ts`）。M2-1 で `xlsx-shapes.ts` はここから引く。
  - `defs.ts` は形の表を読まない。core が geom から引いても 171 KB の表がバンドルに入らないようにするため。
- 変換器：既定の出力を `src/geom/preset-geom.json`、NOTICE をルートの `NOTICE` にした。InterHub の `public/vendor/` への NOTICE の写し（`--notice-out`）はやめた（`package.json` の `files` で NOTICE を配る）。
- NOTICE：冒頭の日本語の説明（パス・同梱先）だけを直した。Ecma の著作権表示の本文は変えていない。
- 再生成した JSON は、176 形の定義が元と同じ。`_source`・`_notice` の中のパスだけが違う。
- テスト：`xlsx-geom.test.ts` → `tests/geom/evaluate.test.ts`、`xlsx-preset-geom.test.ts` → `tests/geom/preset-geom.test.ts`。
  - `public/vendor` の写しのテストは、「NOTICE が `files` に入る」テストに替えた。
  - core の `LINE_PRESETS` との照合は、core が無いので `it.todo` にした。M1-2 で戻す。
- 元のコードは Prettier で整っていない。写したファイルは 1 ファイルずつ `prettier --write <ファイル>` で整えた（中身は変えない。リポジトリ全体への `--write` はしない）。

## 3. render（`lib/viewers/` → `src/render/`）

| ファイル | 行数 | 役割 | 注意 |
| --- | --- | --- | --- |
| `xlsx-grid.ts` | 659 | 範囲・行と列の仮想化・はみ出し・`###`・余白の書式 | `ih-` → `xs-` |
| `xlsx-borders.ts` | 465 | 罫線の線分・辺の優先・path | そのまま |
| `xlsx-border-layer.ts` | 128 | 罫線の SVG 層 | `ih-` → `xs-` |
| `xlsx-format-css.ts` | 165 | 書式のクラスとブック単位の CSS | `ih-xf-`・`data-ih-xlsx-fmt`・`.ih-spreadsheet` → `xs-` 系。`nonce` を付けられるように（M3） |
| `xlsx-drawings.ts` | 327 | アンカー → px・グループの写像・見えている図 | そのまま |
| `xlsx-image-store.ts` | 147 | 画像の blob URL と後始末 | そのまま |
| `xlsx-image-layer.ts` | 778 | 図の DOM 層・破線枠・下地 | `t()` → `messages`、`ih-` → `xs-` |
| `xlsx-shapes.ts` | 912 | 図形と線の SVG・矢じり・外接矩形 | そのまま |
| `spreadsheet.ts`（1,076 行のうち .xlsx の部分） | — | 表と層の組み立て・スクロール・後始末 | `renderXlsxView`（551 行〜）・`showSheet`（813 行〜）・`draw`（707 行〜）・`buildGridTable`（409 行〜）・`dispose`（907 行〜）を `renderSheet` へ。`.xls`（SheetJS）の経路・`fitTableBodyToWindow`・InterHub のバナーとツールバーは写さない |
| `styles/global.css` の `.ih-spreadsheet`・`.ih-xlsx-*` | 31〜1442 行に散在（74 か所） | 表・層・図・罫線・下地の規則と変数 | `src/style.css` へ。変数は [02](02_architecture.md) §6 |
| `locales/ja.ts`・`en.ts` の `spreadsheet.*`・`data.noData` | 23 キー（ja.ts 275〜302 行・248 行） | 文言 | `src/render/messages.ts` の既定値に |

### M2-1 で写した結果（2026-10-28）

- 元はすべて `git cat-file blob 943e4f0:<パス>` で取り出し、1 ファイルずつ Prettier で整えた。
- ファイル名から `xlsx-` を外した（`src/render/` の中はすべて xlsx の描画のため）。テストも同じ（`xlsx-grid.test.ts` → `tests/render/grid.test.ts`）。
- 接頭辞（D6）は次のとおり置き換えた。`ih-` は src・テストとも 0 件。

  | InterHub | XSheet |
  | --- | --- |
  | `ih-xlsx-*`（`ih-xlsx-figure`・`ih-xlsx-frame` など 17 種） | `xs-*`（`xs-figure`・`xs-frame` など） |
  | `ih-xf-N` | `xs-xf-N` |
  | `data-ih-xlsx-fmt` | `data-xs-fmt` |
  | `.ih-spreadsheet`（表の根） | `.xs-sheet` |
  | `--ih-fg`・`--ih-xlsx-cell-bg`・`--ih-xlsx-border-ink`・`--ih-xlsx-backdrop` | `--xs-fg`・`--xs-cell-bg`・`--xs-border-ink`・`--xs-backdrop` |

- import：core は `../core/index`（型を含む）と `../core/internal`（`applyFillMode`）から、geom は `../geom/index` から引く。同じ入口からの import は 1 文にまとめた。
- `shapes.ts` の `Pt`・`PathSeg` は geom の `defs.ts` の型を引いて出し直す（定義は同じ。M1-1 の予定どおり）。
- 文言：`image-layer.ts` の `t()` は、仮置きの `src/render/i18n.ts`（キーをそのまま返す。InterHub の試験と同じ状態）から引く。M2-2 で `messages` に替えて消す。
- コメントの `global.css` は `style.css`（M3-1 で作る）に直した。
- 整形した元に同じ置き換えをかけて比べ、差は上の import・`Pt`・`PathSeg`・`t()` だけであることを確かめた。
- テスト：17 ファイル（681 件）を写した。vitest の render は node 環境にし、DOM が要るファイルは元どおり先頭の `// @vitest-environment jsdom` で jsdom にする（jsdom では fixtures を `new URL(…, import.meta.url)` で読めないため）。
- まだ写さないテスト：
  - `xlsx-format.perf.test.ts`：`spreadsheet.ts` の `buildGridTable` を使う。M2-3 で写す。
  - `xlsx-format-css-static.test.ts`・`xlsx-backdrop-css.test.ts`・`xlsx-borders-css.test.ts`：`global.css` の文面を読む。M3-1（`style.css`）で写す。

### M2-3 で写した結果（2026-10-28）

- 元は `spreadsheet.ts`（`943e4f0`）の .xlsx の部分。`.xls`（SheetJS）の経路・`fitTableBodyToWindow`・タブ列・`renderSpreadsheet` の振り分けは写さない。

  | 元（spreadsheet.ts の行） | 先 |
  | --- | --- |
  | `spacerRow`・`spillOut`・`inkSpill`・`cellBox`・`buildGridTable`（304〜543） | `src/render/grid-table.ts`（`buildGridTable` は計測テストのため export） |
  | `xlsxErrorMessage`（286〜301）・`renderXlsxView` の中の `showSheet`・`draw`・`dispose` ほか（551〜948 のうちタブ以外） | `src/render/sheet-view.ts`（内部の `createSheetView`） |
  | — | `src/render/render-sheet.ts`（公開の `renderSheet`・`RenderOptions`・`SheetView`。`index.ts` から出す） |

- クラス名：`ih-xlsx-grid`→`xs-grid`、`ih-xlsx-collapsed`→`xs-collapsed`、`ih-dv-rownum`→`xs-rownum`、`ih-xlsx-spacer`→`xs-spacer`、`ih-xlsx-cell`→`xs-cell`、`ih-xlsx-text`→`xs-text`、`ih-xlsx-wrap`→`xs-wrap`、`ih-xlsx-body`（と共用の `ih-data-body ih-data-body--table`）→`xs-body`、`ih-xlsx-notice`（と共用の `ih-data-banner`）→`xs-notice`、`ih-data-banner`→`xs-banner`、`ih-dv-empty`→`xs-empty`、表示の根 `ih-data-view ih-spreadsheet`→`xs-sheet`。状態のクラス（`is-spill`・`is-over-fill`・`is-right`・`is-center`）はそのまま。
- 決めたこと：
  - 状態は表示のインスタンスに持つ（`renderGeneration`・`disposeXlsxView` は無い）。`formatSeq` だけモジュールの連番。新しい `showSheet`・`destroy` は `openSeq` を進め、読み込み中の結果を捨てる。
  - 根は自前の `div.xs-sheet`（`container` の子）。子は `div.xs-notice`・`div.xs-body`・`<style>`（`nonce` を付けられる）。タブ列は持たない（M2-4）。
  - 枠の大きさは `clientWidth`・`clientHeight` だけで測る（`getComputedStyle(...).maxHeight` と `refit` はやめた）。`window` の `resize` の代わりにスクロール枠の `ResizeObserver`（無い環境ではガード）。スクロールと大きさの変化は rAF 1 回にまとめる。取り付け後の rAF の描き直し（`initFrame`）は残した。
  - `destroy`：層の `dispose`（blob URL の revoke）を最初に行い、rAF の取り消し・リスナと ResizeObserver の解除・DOM を空にして根を外す。何度呼んでもよい。ブックは閉じない（`renderWorkbook` が閉じる）。
  - `renderSheet` はブックに無い番号を `RangeError` にする（DOM に何も足さない）。`openSheet` の例外は種類ごとの文言の `div.xs-banner` を枠の中に出して解決する。
  - `drawings: false` は `openDrawings` を呼ばず、図の層は空のまま（要素は残す）。
- 形の表の動的 import（M2-3）：`shapes.ts` は geom を型だけ引き、`loadGeom()`（約束を共有・失敗したら捨てて次に読み直す）で `import('../geom/index')` する。読む前・読めないときの `renderShape` は代替（評価できない形と同じ）。表示は形のある図形（`geom` を持つ図）があるシートを開く時だけ読み、読めなければ通知 `libLoadFailed` を出す（文言を「図形の形の表を読み込めなかったため…」に替えた）。ESLint で `src/render/**` の `../geom` を型だけに限った（`@typescript-eslint/no-restricted-imports` の `allowTypeImports`。動的 import は対象外）。render のテストは `tests/render/setup.ts`（vitest の `setupFiles`）で先に読む。
- InterHub との振る舞いの違い：
  - 別のファイルを開いても前の表示を後始末しない（表示ごとに独立。壊すのは使う側の `destroy`）。
  - 通知の出入りで枠の高さを合わせ直さない（`refit` が無い。高さは CSS のレイアウトで決まる）。
  - `libLoadFailed` は xlsx の読み込み失敗ではなく、形の表の読み込み失敗の通知。
- テスト（`tests/render/`）：
  - `spreadsheet.render.test.ts` → `render-sheet.test.ts`（23 件）、`spreadsheet.borders.render.test.ts` → `render-sheet.borders.test.ts`（5 件）、`spreadsheet.cols.render.test.ts` → `render-sheet.cols.test.ts`（6 件）、`spreadsheet.format.render.test.ts` → `render-sheet.format.test.ts`（11 件）、`spreadsheet.image.render.test.ts` → `render-sheet.image.test.ts`（25 件）、`xlsx-format.perf.test.ts` → `format.perf.test.ts`（2 件）。共通の部品は `render-sheet.helpers.ts`。
  - 書き直し：`renderSpreadsheet(buffer, name)` → `renderSheet(container, book, index)`。キーの文字列の期待値は `MESSAGES_JA` の文言に。`cleanupSpreadsheetRendering` → `destroy`（「次のファイルを開くと前が消える」は明示の `destroy` に）。`vi.mock('../xlsx')` → 通知・画像を差し替えたブックを渡す。暗号化などの `openXlsx` の失敗は、文言の対応（`xlsxErrorMessage`）の試験にした（表示は M2-4）。
  - 時間の待ち：並行して走るほかのテストで rAF・画像の読み込みが遅れても落ちないよう、`wait` は先に次のフレームを待ち、画像のテストは読み込み中の画像が無くなるまで待つ。
  - 新しく足した：`render-sheet.lifecycle.test.ts`（2 つの表示を同時に作って壊し、blob URL・リスナ・ResizeObserver・rAF の漏れが 0・根が外れ・ブックが使えること。同じ表示でシートを替えても `<style>` が同じこと）、`render-sheet.geom.test.ts`（geom を読む前の代替・読めないときの通知）、ほか `RangeError`・根の構造・`nonce`・`drawings: false`・読み込み中の `destroy`。
  - 写さないテスト：`.xls` の 4 件（TD-84-12・TD-82-6 の 2 件・TD-86-15）。
  - M2-4（`renderWorkbook`）で写すテスト：
    - `spreadsheet.render.test.ts`：「TD-82-2: シートタブを作り、非表示シートは ⌀ 付きで最初の表示シートを出す」・「シートタブ列は ◀ ▶ ボタンとスクロール領域を持ち…」・「シートタブは title にシート名全文を持つ…」・TD-82-2 の後半（シートが 0 枚なら `noSheets`）・TD-82-4（`openXlsx` の失敗をバナーにする表示）・TD-82-5 の 1 つ目（読み込み中の離脱で空の要素）・「シートタブ列の ResizeObserver 通知」の 2 件。
    - `spreadsheet.borders.render.test.ts`：TD-98-11 の 2 つ目の後半（シートを切り替えると罫線を作り直す）。
    - `spreadsheet.format.render.test.ts`：TD-100-14（シートを切り替えても `<style>` を作り直さない）。
    - `spreadsheet.image.render.test.ts`：TD-86-11 の 2 つ目（タブ切替で revoke・戻ると読み直す）・TD-86-12 の 1 つ目（素早いタブ切替）・TD-86-13 の over-budget の後半（タブ切替で通知が消える）・TD-88-20 の 2 つ目の前半（別のシートで消え、戻ると再表示）。

### M2-4 で写した結果（2026-10-28）

- 元は `spreadsheet.ts`（`943e4f0`）の `makeToggleButton`・`SheetTabStrip`・`buildSheetTabStrip`（112〜254）と、`renderXlsxView` のうちブックを開く・シートが無い・タブの作成とクリック・選択中の印・最初のシート（555〜583・815〜817・889〜903・939）。

  | 元 | 先 |
  | --- | --- |
  | `makeToggleButton`・`buildSheetTabStrip` | `src/render/sheet-tabs.ts`（`makeSheetTab`・`buildSheetTabStrip`） |
  | `renderXlsxView` のブック・タブの部分 | `src/render/render-workbook.ts`（公開の `renderWorkbook`・`WorkbookOptions`・`WorkbookView`。`index.ts` から出す） |

- クラス名：`ih-toolbar ih-sheet-strip`→`xs-sheet-strip`、`ih-sheet-tabs`→`xs-sheet-tabs`、`ih-sheet-nav`→`xs-sheet-nav`、`ih-toggle`（タブ）→`xs-sheet-tab`、表示の根→`xs-workbook`（中に `xs-sheet-strip` と M2-3 の `xs-sheet`）。`is-active`・`is-hidden-sheet` はそのまま。
- 決めたこと：
  - シートの表示は 1 冊に 1 つ。`createSheetView` に `div.xs-workbook` を親として渡し、タブ列の後ろに `div.xs-sheet` を足させる（`createSheetView` は変えていない）。
  - 開けないファイルは例外にせず `div.xs-banner`（種類ごとの文言）、シートが 0 枚は `div.xs-empty`（`noSheets`）で解決する。どちらの `showSheet` も `RangeError`、`destroy` は根を外すだけ。
  - `destroy` の順：タブのクリック・タブ列（◀▶ のクリック・scroll・wheel・ResizeObserver・rAF）→ シートの表示（blob URL の revoke）→ `book.close()` → 根を外す。`openXlsx` の途中の `destroy` は公開の API からは呼べない（約束が解決していない）ため判定しない。
  - ◀▶・タブのクリックは `onclick` ではなく `addEventListener` にし、`destroy` で外す（リスナの漏れを数えて確かめるため）。
- InterHub との違い：
  - 表示中のシートの押し直しは、`activeIndex` が同じなら何もしない。InterHub は空のシート（`sheet` が null）なら開き直していた。開けなかったシートは `activeIndex` が null なので、今までどおり開き直す。
  - `WorkbookView.showSheet(index)` を公開した（タブを押したのと同じ。無い番号は `RangeError`）。`WorkbookOptions.limits` は `openXlsx` に渡す。
- M3-1 の `style.css` に要るもの（タブ列）：`.xs-sheet-strip { display: flex }`、`.xs-sheet-tabs { position: relative; display: flex; flex: 1 1 auto; min-width: 0; overflow-x: auto; overflow-y: hidden; scrollbar-width: none }`（`position: relative` は、タブの `offsetLeft` をスクロール領域の中の座標にするのに要る）、`.xs-sheet-tab { flex-shrink: 0; white-space: nowrap; max-width: 240px; overflow: hidden; text-overflow: ellipsis }`、`.xs-sheet-tab.is-hidden-sheet`（斜体・破線）、`.xs-sheet-nav[hidden] { display: none }`。`div.xs-workbook` は縦の flex にし、`div.xs-sheet` を残りの高さ（`flex: 1; min-height: 0`）にする。元は InterHub `global.css` の 690〜711・1417〜1470 行。
- テスト（`tests/render/`）：
  - `render-workbook.test.ts`（15 件）：タブ（⌀・`is-active`・最初の表示シート・全部非表示なら先頭）・根の構造・◀▶ とスクロール領域・文言の差し替え・title・シート名は textContent・`showSheet`（`RangeError`・押し直し）・`destroy` の後・シートが 0 枚・開けないファイルの文言（暗号化・zip でない・Strict・壊れた zip・`limits` の上限超え）・ArrayBuffer・タブ列の ResizeObserver の 2 件（シートの表示もスクロール枠を観察するため、`.xs-sheet-tabs` を観察するものだけを呼ぶ）。
  - `render-workbook.switch.test.ts`（9 件）：TD-98-11・TD-100-14・TD-86-11・TD-86-12・TD-86-13・TD-88-20 のタブの部分、TD-82-5 の書き直し 2 件、後始末の試験 1 件。`openXlsx` を包んで開いたブックを捕まえ（`vi.mock('../../src/core/index')`）、図の通知・画像の結果を差し替える。
  - TD-82-5（読み込み中の離脱で空の要素を返す）は、公開の API では約束の解決前に壊せないため、「画像の読み込み中に `destroy`」と「タブの切替の途中で `destroy`」に書き直した。
  - 後始末の試験：タブを替えてから `destroy` すると、blob URL はすべて revoke、リスナ（`EventTarget` の add・remove を数える）・ResizeObserver（タブ列とスクロール枠の 2 つ）・予約中の rAF が 0、根が外れ、捕まえたブックの `openSheet` が例外になる（閉じた）。
- 計測（`*.perf.test.ts`）：ほかのテストと同時に走ると時間がぶれて落ちることがあった（M2-3 の後の全体の実行で `borders.perf`）。vitest の project `perf` に分け、ほかが終わってから（`sequence.groupOrder: 1`）1 ファイルずつ（`fileParallelism: false`）走らせる。計測の中身・閾値は変えていない。

## 4. 直す 6 点

| # | 直すこと | 場所 | 方法 |
| --- | --- | --- | --- |
| 1 | 層の逆転 | `drawing-geom.ts`・`types.ts` が `viewers/xlsx-geom` を import（`943e4f0` で確かめると、本体ではなく評価器 `xlsx-geom.ts` が core の `drawing-geom.ts`・`limits.ts` と render の `xlsx-shapes.ts` を import していた。core から評価器を引くのはテストだけ） | geom を独立の層にする（§2。M1-1 で `defs.ts` を作った） |
| 2 | InterHub の文言関数 | `spreadsheet.ts`・`xlsx-image-layer.ts` の `t()` | `RenderOptions.messages` を引数で回す |
| 3 | 接頭辞 | `ih-xlsx-*`・`ih-xf-N`・`data-ih-xlsx-fmt`・`--ih-*` | `xs-` 系へ一括で置き換え、テストの期待値も合わせる |
| 4 | モジュール変数の状態 | `spreadsheet.ts` の `disposeXlsxView`・`renderGeneration` | `SheetView` のインスタンスへ |
| 5 | InterHub の共用部品 | `table-fit.ts`（`fitTableBodyToWindow`）・`ih-data-view`・`ih-data-banner`・`ih-dv-rownum`・`ih-sheet-tabs` | render に要るものだけ `xs-` で作り直す。枠の高さは使う側の仕事 |
| 6 | 型だけの import | 全体 | `verbatimModuleSyntax: true` にして `import type` へ直す（xrag-Viewer はこの設定） |

## 5. テスト（写す）

| 元 | 件数 | 先 |
| --- | --- | --- |
| `lib/xlsx/*.test.ts` | 32 ファイル（当初の見積もりは 36） | `tests/core/` |
| `lib/viewers/xlsx-*.test.ts`・`spreadsheet.*.render.test.ts` | 22（geom の 2 を含む。当初の見積もりは 26）＋5 ファイル | `tests/render/`（`spreadsheet.*` は `renderSheet` を呼ぶ形に書き直す） |
| `lib/xlsx/__fixtures__/excel/` | 23 ファイル | `fixtures/excel/`（[05_quality.md](05_quality.md) §2） |

- テストは vitest＋jsdom。`node:fs` で fixtures を読むテストがある（Node で走る前提）。
- InterHub の版：vitest 3.2・jsdom 29・TypeScript 5.9・vite 6.4。新しい版（xrag は vitest 5・jsdom 30・TS 7・vite 8）へ上げるときの落ちは M1・M2 で直す。
