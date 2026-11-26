// 正: docs/adr/0011-selection.md §5
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
export const MAX_COPY_CELLS = 1_000_000;

/** 行を読む単位 */
const READ_CHUNK = 512;

/** コピーする範囲。行の全体は値のある最後の列まで、列の全体は値のある最後の行まで（空なら r2 < r1 か c2 < c1） */
export function copyRange(sel: Selection, used: { rows: number; cols: number }): CellRange {
  const { r1, c1, r2, c2 } = sel.range;
  return {
    r1,
    c1,
    r2: sel.type === 'column' ? Math.min(r2, used.rows) : r2,
    c2: sel.type === 'row' ? Math.min(c2, used.cols) : c2,
  };
}

/** セル 1 つの文字（タブ・改行・`"` を含めば `"` で囲む） */
export function quoteCell(text: string): string {
  return /[\t\r\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

/**
 * 範囲の文字。見える行 × 見える列が `maxCells` を超えたら null。
 * 結合の左上が範囲（と読む区切り）より上にあると、`getRows` は区切りの最初の行に左上の文字を置くので、
 * 1 行前から読んで捨てる（範囲の中の続きの行は空のまま）。
 */
export function rangeText(
  sheet: Pick<XlsxSheet, 'getRows' | 'rowCount'>,
  range: CellRange,
  geo: Pick<Geometry, 'rowHeight' | 'colWidth'>,
  maxCells = MAX_COPY_CELLS,
): string | null {
  const cols: number[] = [];
  for (let c = range.c1; c <= range.c2; c++) if (geo.colWidth(c) > 0) cols.push(c);
  let shownRows = 0;
  for (let r = range.r1; r <= range.r2; r++) {
    if (geo.rowHeight(r) > 0 && ++shownRows * cols.length > maxCells) return null;
  }
  if (shownRows === 0 || cols.length === 0) return '';
  const lines: string[] = [];
  for (let a = range.r1; a <= range.r2; a += READ_CHUNK) {
    const b = Math.min(range.r2, a + READ_CHUNK - 1);
    const from = Math.max(1, a - 1);
    const data = a <= sheet.rowCount ? sheet.getRows(from, Math.min(b, sheet.rowCount)) : [];
    for (let r = a; r <= b; r++) {
      if (geo.rowHeight(r) === 0) continue;
      const cells = data[r - from]?.cells;
      lines.push(cols.map((c) => quoteCell(cells?.[c - 1]?.text ?? '')).join('\t'));
    }
  }
  return lines.join('\r\n') + '\r\n';
}
