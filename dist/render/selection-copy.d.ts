/**
 * 選択をコピーする文字（XSheet の追加。M7-26・Issue #57）。DOM を使わない。
 * - 見えている文字（`CellView.text`。表示形式を当てた後）を、タブと `\r\n` で区切る（Excel のコピーと同じ形。最後の行にも `\r\n`）。
 * - タブ・改行・`"` を含むセルは `"` で囲み、`"` は `""` にする。結合したセルは左上に文字、残りは空。
 * - 非表示（高さ 0・幅 0）の行・列は入れない。行・列の全体は、値のある範囲（`rowCount`・`colCount`）までに切る。
 * - セルの数が上限を超えたら null（コピーしない）。
 * `selection.ts` と同じく遅れて読む部品の中にあり、表示の部品を import しない。
 */
import type { CellRange, Geometry, XlsxSheet } from '../core/index';
import type { Selection } from './selection';
/** コピーするセルの数の上限（見える行 × 見える列） */
export declare const MAX_COPY_CELLS = 1000000;
/** コピーする範囲。行の全体は値のある最後の列まで、列の全体は値のある最後の行まで（空なら r2 < r1 か c2 < c1） */
export declare function copyRange(sel: Selection, used: {
    rows: number;
    cols: number;
}): CellRange;
/** セル 1 つの文字（タブ・改行・`"` を含めば `"` で囲む） */
export declare function quoteCell(text: string): string;
/**
 * 範囲の文字。見える行 × 見える列が `maxCells` を超えたら null。
 * 結合の左上が範囲（と読む区切り）より上にあると、`getRows` は区切りの最初の行に左上の文字を置くので、
 * 1 行前から読んで捨てる（範囲の中の続きの行は空のまま）。
 */
export declare function rangeText(sheet: Pick<XlsxSheet, 'getRows' | 'rowCount'>, range: CellRange, geo: Pick<Geometry, 'rowHeight' | 'colWidth'>, maxCells?: number): string | null;
