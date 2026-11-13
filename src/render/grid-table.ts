// 正: InterHub docs/plan/adr/0026-xlsx-image-overlay.md（§決定 4・5）・0030-xlsx-cell-format.md（§決定 1・9）
/**
 * `.xlsx` グリッドの table を組む（InterHub `spreadsheet.ts` の `buildGridTable` とその部品・M2-3）。
 * 寸法・範囲の計算は DOM 非依存の `grid.ts`、書式の CSS は `format-css.ts` にあり、ここは DOM を組むだけ。
 * ユーザーのファイル内容（セル値）は信頼できない入力として扱い、`innerHTML` は使わず
 * `document.createElement` + `textContent`（元の値は `title`）だけで入れる。
 */
import type { CellFormat, Geometry } from '../core/index';
import { formatClassName, spillInkGradient } from './format-css';
// XSheet: 検索の一致のハイライト（ADR-0007 §2）
import { decorateCell, type GridHits } from './highlight';
import {
  cellAlignClass,
  cellInnerHeight,
  clipRowToColumns,
  colSpacerWidths,
  fitNumberText,
  rowNumWidth,
  spacerHeights,
  spanHeight,
  spillColumns,
  spillOverFill,
  spillWidth,
  type ColRange,
  type GridRow,
} from './grid';

/** 上下の描かない行の高さを持つ空の行（行仮想化の余白）。 */
function spacerRow(height: number, cols: number): HTMLTableRowElement {
  const tr = document.createElement('tr');
  tr.className = 'xs-spacer';
  tr.setAttribute('aria-hidden', 'true');
  const px = `${Math.max(0, height)}px`;
  tr.style.height = px;
  const td = document.createElement('td');
  td.colSpan = cols + 1;
  td.style.height = px;
  tr.append(td);
  return tr;
}

/**
 * 文字がはみ出す入れ物にする：幅を右の空セルの列幅の合計まで広げ（`overflow: hidden` で値のあるセルの手前で切る）、
 * 文字を span で包む（`style.css` の `.is-spill .xs-text` が背景色を持ち、文字の範囲だけ下の罫線を覆う。
 * 塗りが絡むときは呼び出し側が `is-over-fill` を足し、その背景を外す）。
 * td 自体は `overflow: visible` のため、広げた入れ物が右の空セルの上へ出る。
 */
function spillOut(box: HTMLDivElement, text: string, width: number): void {
  const label = document.createElement('span');
  label.className = 'xs-text';
  label.textContent = text;
  box.classList.add('is-spill');
  box.style.width = `${width}px`;
  box.replaceChildren(label);
}

/**
 * 塗りの絡むはみ出し（`is-over-fill`）の文字色を、列の面に合わせて塗り分ける（ADR-0030 §決定 9）。
 * 元の文字が「自動」色で、またぐ列の面が塗りあり（黒）と塗りなし（画面の文字色）に分かれるときだけ、span に
 * 列ごとの区間のグラデーション＋`background-clip: text`＋透明な文字色をインラインで付ける（はみ出しはまれなので少数）。
 * 明示色・面が同じ・span の無い入れ物には何もしない（単色のまま）。
 *
 * @param k 描く列の中での元の位置（`formats` の添字）
 * @param c 元の列（0 始まり。geometry は c + 1）
 * @param columns `spillColumns` の結果
 */
function inkSpill(
  box: HTMLElement,
  table: readonly CellFormat[],
  formats: readonly (number | undefined)[] | undefined,
  k: number,
  geo: Geometry,
  c: number,
  columns: number,
): void {
  const label = box.firstElementChild;
  if (!(label instanceof HTMLElement)) {
    return;
  }
  const widths: number[] = [];
  for (let i = 0; i < columns; i++) {
    widths.push(geo.colWidth(c + 1 + i));
  }
  const image = spillInkGradient(table, formats, k, widths);
  if (image === null) {
    return;
  }
  label.style.setProperty('background-image', image);
  label.style.setProperty('-webkit-background-clip', 'text');
  label.style.setProperty('background-clip', 'text');
  label.style.setProperty('color', 'transparent');
}

/**
 * 内側の要素（高さ固定・はみ出しを隠す）を持つ td の中身を作る。
 * 行の高さを固定するため、文字は高さ固定の `div` の中に入れる（`style.css` の `.xs-cell`）。
 * セル値は `textContent` のみで入れる。
 */
function cellBox(text: string, height: number): HTMLDivElement {
  const box = document.createElement('div');
  box.className = 'xs-cell';
  box.style.height = `${height}px`;
  box.textContent = text;
  return box;
}

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
 * @param hits 検索の一致のハイライト（XSheet: ADR-0007 §2。無ければ null）
 */
export function buildGridTable(
  colName: (c: number) => string,
  cols: number,
  rows: GridRow[],
  geo: Geometry,
  rowCount: number,
  mdw: number,
  range: Pick<ColRange, 'first' | 'last'>,
  formats: readonly CellFormat[],
  hits: GridHits | null = null,
): HTMLTableElement {
  const table = document.createElement('table');
  table.className = 'xs-grid';
  table.setAttribute('aria-rowcount', String(rowCount + 1));
  const numWidth = rowNumWidth(rowCount);
  table.style.width = `${numWidth + geo.totalWidth}px`;

  const { first, last } = range;
  const pad = colSpacerWidths(geo, first, last);
  // 幅 0 の余白（先頭・末尾の列が非表示など）は作らない（行の td と数を揃える）
  const leftPad = first > 1 && pad.left > 0 ? pad.left : 0;
  const rightPad = last < cols && pad.right > 0 ? pad.right : 0;
  /** DOM 上の列数（行番号の列＋余白の列＋描く列） */
  const domCols = 1 + (leftPad > 0 ? 1 : 0) + (last - first + 1) + (rightPad > 0 ? 1 : 0);
  const padCol = (width: number): HTMLTableColElement => {
    const col = document.createElement('col');
    col.style.width = `${width}px`;
    return col;
  };

  const group = document.createElement('colgroup');
  const numCol = document.createElement('col');
  numCol.style.width = `${numWidth}px`;
  group.append(numCol);
  if (leftPad > 0) {
    group.append(padCol(leftPad));
  }
  for (let c = first; c <= last; c++) {
    const col = document.createElement('col');
    const w = geo.colWidth(c);
    col.style.width = `${w}px`;
    if (w === 0) {
      col.className = 'xs-collapsed';
    }
    group.append(col);
  }
  if (rightPad > 0) {
    group.append(padCol(rightPad));
  }
  table.append(group);

  const head = table.createTHead().insertRow();
  const corner = document.createElement('th');
  corner.className = 'xs-rownum';
  head.append(corner);
  if (leftPad > 0) {
    head.append(document.createElement('th'));
  }
  for (let c = first; c <= last; c++) {
    const th = document.createElement('th');
    th.textContent = colName(c);
    head.append(th);
  }
  if (rightPad > 0) {
    head.append(document.createElement('th'));
  }

  const rowPad = spacerHeights(geo, rows[0]?.r ?? 1, rows[rows.length - 1]?.r ?? 0);
  const tbody = table.createTBody();
  tbody.append(spacerRow(rowPad.top, domCols - 1));
  rows.forEach((source, i) => {
    const row = clipRowToColumns(source, first, last);
    const n = last - first + 1;
    const tr = document.createElement('tr');
    tr.setAttribute('aria-rowindex', String(row.r + 1));
    const h = geo.rowHeight(row.r);
    tr.style.height = `${h}px`;
    const rowHead = document.createElement('td');
    rowHead.className = 'xs-rownum';
    rowHead.append(cellBox(String(row.r), cellInnerHeight(h)));
    tr.append(rowHead);
    if (leftPad > 0) {
      tr.append(document.createElement('td'));
    }
    for (let k = 0; k < n; k++) {
      const c = first - 1 + k; // 列（0 始まり。geometry は c + 1）
      const cell = row.cells[k];
      if (cell === null) {
        continue; // 結合で隠れるセル
      }
      const td = document.createElement('td');
      const fi = row.formats?.[k];
      const fmt = fi === undefined ? undefined : formats[fi];
      if (fi !== undefined && fmt) {
        td.className = formatClassName(fi); // 値の無い塗りだけのセルにも付く
      }
      if (cell) {
        const colSpan = Math.max(1, cell.colSpan);
        const width = geo.colLeft(c + 1 + colSpan) - geo.colLeft(c + 1);
        if (cell.text !== '' && width > 0) {
          const shown = fitNumberText(cell.text, cell.kind, width, mdw, fmt?.sizeRatio ?? 1);
          const box = cellBox(shown, cellInnerHeight(spanHeight(geo, rows, i, cell.rowSpan)));
          const align = cellAlignClass(cell.kind);
          if (align !== '') {
            box.classList.add(align);
          }
          if (shown !== cell.text) {
            box.title = cell.text; // 元の値（ホバーで見える）
          }
          const columns = spillColumns(row.cells, k, n); // 描く範囲の右端で切る
          if (columns > 1) {
            spillOut(box, cell.text, spillWidth(geo, c, columns));
            if (spillOverFill(formats, row.formats, k, columns)) {
              box.classList.add('is-over-fill'); // 文字の下に面の色を敷かない（塗りを隠さない）
              inkSpill(box, formats, row.formats, k, geo, c, columns);
            }
          }
          // XSheet: 検索の一致を光らせる口（highlight.ts・ADR-0007 §2）
          if (hits) {
            decorateCell(box, cell.text, shown, hits.at(row.r, c + 1), hits.current);
          }
          td.append(box);
        }
        if (colSpan > 1) {
          td.colSpan = colSpan;
        }
        if (cell.rowSpan > 1) {
          td.rowSpan = cell.rowSpan;
        }
      }
      tr.append(td);
    }
    if (rightPad > 0) {
      tr.append(document.createElement('td'));
    }
    tbody.append(tr);
  });
  tbody.append(spacerRow(rowPad.bottom, domCols - 1));
  return table;
}
