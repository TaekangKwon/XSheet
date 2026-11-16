// 正: docs/adr/0008-align-wrap-spill.md §4
/**
 * 最後の列からはみ出す文字の分だけ、表の列数を広げる（M7-11・Moobius の依頼 R8）。
 * - 行ごとに「値のある最後のセル」が文字列ではみ出せるときだけ、`列の左端 + 文字の幅の見積もり` を含む列を数える。
 * - 文字の幅は DOM で測らず、文字数から多めに見積もる（多めに出た分は空の列が増えるだけで、文字は切れない）。
 * - 調べるのは描いた行と移動先の行だけ（シートを開くときに全行を読まない）。広げた列数は表示（`sheet-view.ts`）が持つ。
 * DOM 非依存の純関数。
 */
import type { CellFormat, Geometry } from '../core/index';
import type { GridRow } from './grid';

/** 表の文字の大きさ（px・倍率 1）。`style.css` の `.xs-grid` の `font-size` と同じ値 */
const GRID_FONT_PX = 13;
/** 全角の文字の幅（em） */
const FULL_EM = 1;
/** 全角でない文字の幅（em）。多めに出す（英字の平均は 0.55em 前後） */
const HALF_EM = 0.7;
/** 太字・斜体で広がる分 */
const STYLE_FACTOR = 1.1;
/** セルの左右の余白（`style.css` の `.xs-cell` の padding）と端数 */
const PAD_PX = 8;
/** 見積もる文字数の上限（長い文字列でも 1 行あたりの計算を抑える。超えた分は全角として足す） */
const MAX_SCAN_CHARS = 4096;

/** 全角として数える文字か（ハングル字母より後ろ。半角カナ FF61〜FF9F・半角記号 FFE8〜FFEE を除く） */
function isFullWidth(cp: number): boolean {
  if (cp < 0x1100) return false;
  if (cp >= 0xff61 && cp <= 0xff9f) return false;
  if (cp >= 0xffe8 && cp <= 0xffee) return false;
  return true;
}

/**
 * 文字の幅の見積もり（px・倍率 1）。多めに出す。
 * @param text 表示する文字（1 行。改行は空白として数える）
 * @param format セルの書式（文字の大きさの比・太字・斜体）
 */
export function estimateCellTextWidth(text: string, format?: CellFormat): number {
  let em = 0;
  let n = 0;
  for (const ch of text) {
    if (n >= MAX_SCAN_CHARS) {
      em += (text.length - n) * FULL_EM; // 残りは全角として多めに足す（UTF-16 の長さで数える）
      break;
    }
    em += isFullWidth(ch.codePointAt(0)!) ? FULL_EM : HALF_EM;
    n += ch.length;
  }
  const ratio = format?.sizeRatio ?? 1;
  const style = format?.bold || format?.italic ? STYLE_FACTOR : 1;
  return em * GRID_FONT_PX * ratio * style + PAD_PX;
}

/** 値のあるセルか（`grid.ts` の空のセルの判定の逆。null＝結合で隠れるセルも値ありとして止める） */
function hasValue(cell: GridRow['cells'][number]): boolean {
  return (
    cell !== undefined &&
    (cell === null ||
      cell.text !== '' ||
      cell.kind !== undefined ||
      cell.colSpan !== 1 ||
      cell.rowSpan !== 1 ||
      cell.merged === true)
  );
}

/**
 * 行の文字がはみ出して届く列（1 始まり）の最大。はみ出す文字の無い行だけなら 0。
 * 行の「値のある最後のセル」だけを見る（それより左のセルは、右に値があるので最後の列を越えない）。
 *
 * @param geo 倍率 1 の寸法（表の範囲より右の列も既定の幅で延びるもの）
 * @param rows 調べる行（`GridRow`。`cells[i]` は列 i + 1）
 * @param formats 書式の表（`XlsxSheet.formats`）
 */
export function spillReachCols(
  geo: Pick<Geometry, 'colLeft' | 'colAt' | 'colWidth'>,
  rows: readonly Pick<GridRow, 'cells' | 'formats'>[],
  formats: readonly CellFormat[],
): number {
  let max = 0;
  for (const row of rows) {
    const cells = row.cells;
    let i = cells.length - 1;
    while (i >= 0 && !hasValue(cells[i])) i--;
    const cell = i >= 0 ? cells[i] : undefined;
    if (
      !cell ||
      cell.kind !== 's' ||
      cell.text === '' ||
      cell.colSpan > 1 ||
      cell.rowSpan > 1 ||
      cell.merged === true ||
      geo.colWidth(i + 1) === 0
    ) {
      continue;
    }
    const f = row.formats?.[i];
    const width = estimateCellTextWidth(cell.text, f === undefined ? undefined : formats[f]);
    const reach = geo.colAt(geo.colLeft(i + 1) + width);
    if (reach > max) max = reach;
  }
  return max;
}
