/**
 * `.xlsx` グリッドの table を組む（InterHub `spreadsheet.ts` の `buildGridTable` とその部品・M2-3）。
 * 寸法・範囲の計算は DOM 非依存の `grid.ts`、書式の CSS は `format-css.ts` にあり、ここは DOM を組むだけ。
 * ユーザーのファイル内容（セル値）は信頼できない入力として扱い、`innerHTML` は使わず
 * `document.createElement` + `textContent`（元の値は `title`）だけで入れる。
 */
import type { CellFormat, Geometry } from '../core/index';
import { type ColRange, type GridRow } from './grid';
/**
 * グリッドの table を Excel の寸法（geometry）で固定して組む（計測テスト `format.perf.test.ts` のため export）：
 * `colgroup`（列幅）・見出し行（列記号）・上の余白・行（Excel の行番号＋セル）・下の余白。
 * - 列幅は `<col>` の `width`、行の高さは `<tr>` の `height` とセルの内側の要素の高さで固定する（`table-layout: fixed`）。
 * - 非表示の列は幅 0 とし、`<col>` に `xs-collapsed`（`visibility: collapse`）を付けて畳む（ADR-0026 §決定 5）。
 *   非表示の行は `rows` に含まれない（呼び出し側で除く）。結合セルの `rowSpan` は描く行の数。
 * - セル値は `textContent` のみで入れる。結合で隠れるセル（null）は td を作らない。td は 1 セル 1 個（罫線を残す）。
 *   文字列のはみ出しは、内側の要素を右の空セルの列幅の合計まで広げ、文字の範囲だけ背景色で下の罫線を覆う。
 *   セルの幅（結合を含む）が 0 なら内側の要素を作らない（畳んだ列の文字を出さない）。
 * - 配置は型ごとの既定（数値・日付は右、真偽・エラーは中央、文字列は左）。収まらない数値・日付は `###` にし、元の値を title に入れる
 *   （大きな文字のセルは推定幅に文字サイズの比を掛ける）。
 * - 書式（塗り・文字の装飾）のあるセルの td には、className に `xs-xf-<番号>` を 1 語足すだけ（規則は表示の根の `<style>`。style 属性は書かない）。
 *   塗り（白以外）が絡むはみ出しの入れ物には `is-over-fill` を付ける（ADR-0030 §決定 9）。
 * - 列も仮想化する：描く列は `range`（可視範囲の左右 1 画面分）だけ。描かない左右の列は、幅 `colLeft` の差の
 *   余白の列（`<col>` 1 本＋各行に空の td 1 個）で置き換え、表の全幅は geometry の totalWidth のまま保つ。
 *   範囲の端で切れる結合セルは `clipRowToColumns` が `colSpan` を縮める。
 *
 * @param colName 列番号（1 始まり）→ 列記号
 * @param cols 表の列数（データの列数より広いことがある）
 * @param rows 描く行（見える行だけ。行番号の昇順）
 * @param geo 寸法の換算結果（`createGeometry`）
 * @param rowCount 表の行数（データの行数より広いことがある）
 * @param mdw 最大桁幅 MDW（`SheetLayout.mdw`。収まらない数値を `###` にする推定幅に使う）
 * @param range 描く列の範囲（`first`〜`last`・1 始まり）。表の全列を描くときは 1〜`cols`
 * @param formats 書式の表（`XlsxSheet.formats`。`rows[i].formats` の値が指す）
 */
export declare function buildGridTable(colName: (c: number) => string, cols: number, rows: GridRow[], geo: Geometry, rowCount: number, mdw: number, range: Pick<ColRange, 'first' | 'last'>, formats: readonly CellFormat[]): HTMLTableElement;
