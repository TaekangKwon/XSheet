# 05. 品質 — テスト・Excel との照合・性能・既知の制限

| 項目 | 値 |
| --- | --- |
| 文書 | XSheet 立ち上げ資料 05 |
| 更新日 | 2026-10-28 |
| 関連 | [01_requirements.md](01_requirements.md)・[04_roadmap.md](04_roadmap.md) |

## TL;DR
- 正しさは 3 段で確かめる：①合成の xlsx を使う vitest（約 2,000 件を引き継ぐ）②Excel で作った実ファイルと Excel 自身の値との照合 ③実機の目視。
- 性能の目安は InterHub の実測：100 万セルで 0.5〜0.7 秒・図形 1,000 個で 200ms 超のフレーム 0・全罫線 10,000 セルで罫線層 10〜22ms。
- 暫定のまま引き継ぐ項目（罫線 E1〜E5・書式 F1〜F10）と既知の制限を、v0.1.0 の README に書く。

---

## 1. テスト

| 群 | 何を確かめるか | 主なファイル（InterHub の名前） |
| --- | --- | --- |
| 入口・安全 | 先頭の印・暗号化・壊れた zip・上限・線形時間 | `open`・`workbook`・`safety`・`linear-time` |
| セル・寸法 | 表示文字列・結合・行の索引・列幅と行高の換算 | `cells`・`rows`・`layout`・`geometry` |
| 罫線・書式・色 | 罫線の表・書式の表・セルの色・色の変換・テーマ | `borders`・`sheet.borders`・`cell-format`・`sheet.formats`・`cell-color`・`color`・`theme` |
| 図の読み取り | drawing・グループ・線・図形の文字・色・形・上限 | `drawing*` |
| 画像 | 形式の判定・予算・壊れた画像 | `image`・`image-read` |
| 実ファイルとの照合 | Excel で作ったファイルの値と Excel の値 | `ref-*` |
| 描画の部品 | 仮想化・罫線の線分・書式の CSS・図の配置・図形の SVG・下地 | `xlsx-*` |
| DOM の結合 | 表・列・罫線・書式・画像が DOM に正しく出る | `spreadsheet.*.render` |
| 性能 | 罫線・書式の構築時間 | `xlsx-borders.perf`・`xlsx-format.perf` |

- 合成の xlsx は `src/testing/make-xlsx.ts`（シート・セル・行と列・書式・罫線・図・テーマを指定して zip を作る）。新しいテストもこれで作る。
- 外部のファイルを持ち込まない。fixtures は自作の合成か、Excel で自作したものだけ。

## 2. Excel との照合（`fixtures/excel/`）

| ファイル | 役割 |
| --- | --- |
| `ref-calibri.xlsx`・`ref-calibri-stdw12.xlsx`・`ref-yugothic.xlsx`・`ref-malgun.xlsx` | 既定フォント別の列幅・行高・図の位置の基準 |
| `ref-drawings.xlsx` | 図の一覧 |
| `ref-shapes.xlsx`＋`.expected.json` | 図形の種類・件数・回転と反転（133 図形） |
| `ref-geoms.xlsx`＋`.expected.json` | 図形の形と枠線（225 図形） |
| `ref-colors.xlsx`＋`.expected.json` | 図形・線・文字の色（128 図形） |
| `ref-geoms.compare.json`・`ref-flips.compare.json` | 画像の輪郭の比較の基準値 |
| `showcase.xlsx`＋`make-showcase.ps1` | README の画面（`docs/images/screenshot.png`）と例のページの見本。照合の基準ではない。XSheet がまだ描かないもの（セルの配置など）は使わない。画面の撮り方：`npm run examples` を動かし、Edge のヘッドレス（`--window-size=1280,720 --force-device-scale-factor=2 --screenshot`）で `http://localhost:5173/examples/?sample=showcase.xlsx` を撮って、上の操作の帯と空の行を切り落とす |
| `make-ref-*.ps1`（4 本）・`make-ref-*.extract.mjs`（3 本） | Excel 16 の COM でファイルと基準値を作る（Windows＋Excel が要る・手動） |
| `geom-compare.mjs`・`.d.mts`・`compare-ref-geoms.mjs` | 輪郭の比較器と、Playwright で描画を撮って比べる道具（手動）。描画の組み立て（ハーネス）は XSheet の `renderWorkbook` を Vite で束ねる。手順：① `make-ref-geom-pngs.ps1 -Dir <dir>`（Excel が要る）② `npm run geom:compare -- <dir> [--only p108,p142]` ③ 出た `ref-*.compare.json` を `fixtures/excel/` へ写して `ref-geoms.compare.test.ts` |

- 照合の結果（InterHub）：
  - 列幅・行高・図の位置は差 0px。
  - 線は 1px 以内。
  - 図形の形は中心線の p95 ≤ 1.5px が 191/207。超えた 16 形は原因が分かっている既知の差。
  - 色は最大差 1。
- COM を使うときの注意：
  - 画面が 125% のとき、COM の 0.6pt＝1px。
  - `.ps1` は UTF-8 BOM 付きで保存する。
  - 実行は `powershell -ExecutionPolicy Bypass -File …`。
  - PNG の書き出しでは `Visible = $true` にする。クリップボードを上書きする。

## 3. 暫定の項目（Excel で未確認）

| 区分 | 項目 | 元 |
| --- | --- | --- |
| 罫線 | E1〜E5（隣り合う辺の優先・線のパターンの長さなど） | InterHub ADR-0029 |
| 書式 | F1〜F10（模様の濃さの混色・グラデーションの代表色・`###` の数・ダーク表示の描き分けなど） | InterHub ADR-0030 |

- M7 で Excel 16 の COM を使って確かめ、結果を ADR とテストに反映する。

## 4. 既知の制限

- 範囲外の明示的な `<c s=0/>` では、列の書式が消えない。
- 列の書式があるシートでは、`getRows` が列数ぶんの配列を作る（メモリが増える）。
- 色がグラデーションのはみ出し文字では、下線などの装飾が透明になることがある。
- フォントの差で、列幅が数 px、図形の文字の折り返しが Excel とずれることがある。
- ダーク表示では、図の白い下地の下のセルが見えない。
- グラフは G1（縦棒・横棒・折れ線・円・ドーナツ）だけ描き、ほかの種類と SmartArt は種類名入りの破線枠。軸の自動の範囲・配置は Excel の近似で、文字幅の推定の差などが残る（InterHub ADR-0031 §決定 16）。

## 5. 性能の測り方

| 対象 | 方法 | 目安（InterHub） |
| --- | --- | --- |
| 開く速さ | 大きい合成ファイル（100 万セル）で `openXlsx`＋`openSheet`＋最初の `getRows` | 0.5〜0.7 秒 |
| スクロール | Playwright で 3 往復し、rAF の間隔を記録 | 200ms 超のフレーム 0（図形 1,000＋線 1,000：最大 81〜98ms） |
| 罫線層 | 10,000 セル全罫線で層の構築 | 10〜22ms |
| 非表示行 | 非表示の多いシートのスクロール | 60ms（DOM に出さない） |

- 画面の自動計測は、ソースを変えずに `renderSheet` を esbuild でバンドルしたページを作り、Playwright で開いて測る（InterHub で使った方法）。
