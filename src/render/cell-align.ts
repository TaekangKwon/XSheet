// 正: docs/adr/0008-align-wrap-spill.md §2・§3
/**
 * セルの配置と折り返しの描き方（M7-13・Moobius の依頼 R9）。
 * - 入れ物（`div.xs-cell`）に配置のクラスを付ける（{@link alignClasses}）。規則は `style.css` の `is-h-*`・`is-v-*`・`is-wrap`。
 * - 右寄せの文字列は左へ、中央の文字列は両側へはみ出す（{@link spillSides}・{@link spillAround}）。左寄せ・標準の右へのはみ出しは
 *   今の `grid-table.ts`（`spillColumns`・`spillOut`）のまま。折り返すセルははみ出さない。
 * - 位置の判定（`navigate.ts`）も {@link spillSides} で、左右どちらから届くはみ出しかを引く。
 * `grid-table.ts`（InterHub の写し）には `// XSheet:` の印で呼び出しの口だけを足す。
 * セル値は `textContent` だけで入れる（信頼できない入力）。
 */
import type { CellFormat, CellKind, Geometry } from '../core/index';
import { spillOverFill, type GridCell, type GridRow } from './grid';

/** 文字がはみ出して使える列数の上限（`grid.ts` の `MAX_SPILL_COLS` と同じ） */
const MAX_SPILL_COLS = 1000;
/** 列の右端の罫線（`grid.ts` の `BORDER` と同じ） */
const BORDER = 1;
/** 入れ物の左右の余白（`style.css` の `.xs-cell` の padding） */
const PAD = 3;

/** 横の配置（標準は型で決めたもの） */
export type HAlign = 'left' | 'center' | 'right';

/**
 * 文字列（型 `s`）か。配置の指定で `safe` を付けない・折り返す・左右へはみ出すのは文字列だけ。
 * 型の無いセル（結合セルの左上。リーダーは型を付けない）も文字列として扱う（結合の中で折り返す。はみ出しは結合なので起きない）。
 */
const isText = (kind: CellKind | undefined): boolean => kind === 's' || kind === undefined;

/**
 * 実際の横の配置。指定（`hAlign`）が無ければ型で決める（数値・日付は右、真偽・エラーは中央、ほかは左）。
 */
export function resolveHAlign(kind: CellKind | undefined, format?: CellFormat): HAlign {
  if (format?.hAlign !== undefined) return format.hAlign;
  if (kind === 'n' || kind === 'd') return 'right';
  if (kind === 'b' || kind === 'e') return 'center';
  return 'left';
}

/** 折り返すセルか（文字列で `wrap` の指定がある） */
export function isWrapped(kind: CellKind | undefined, format?: CellFormat): boolean {
  return format?.wrap === true && isText(kind);
}

/**
 * 入れ物に足すクラス。指定の無いセルは空（型ごとの既定は `grid.ts` の `cellAlignClass`）。
 * - 横：文字列は `is-h-center`・`is-h-right`（`safe` を付けない）。文字列でないセル（数値など）は今の `is-center`・`is-right`
 *   （`safe`：収まらない `###` は左から）。左は付けない。
 * - 縦：`is-v-top`・`is-v-center`。折り返し：`is-wrap`（文字列だけ）。
 */
export function alignClasses(kind: CellKind | undefined, format?: CellFormat): string[] {
  const out: string[] = [];
  const h = format?.hAlign;
  if (h === 'center' || h === 'right') {
    out.push(isText(kind) ? `is-h-${h}` : `is-${h}`);
  }
  if (format?.vAlign !== undefined) out.push(`is-v-${format.vAlign}`);
  if (isWrapped(kind, format)) out.push('is-wrap');
  return out;
}

/** 空のセルか（`grid.ts` の判定と同じ。null＝結合で隠れるセルは空ではない） */
function isEmptyCell(cell: GridCell | null | undefined): boolean {
  return (
    cell === undefined ||
    (cell !== null &&
      cell.text === '' &&
      cell.kind === undefined &&
      cell.colSpan === 1 &&
      cell.rowSpan === 1 &&
      cell.merged !== true)
  );
}

/** 左右へはみ出す列数（自分の列を含まない） */
export interface SpillSides {
  left: number;
  right: number;
}

/**
 * 文字列が左右へはみ出して使える列数。左寄せ・標準は右だけ、右寄せは左だけ、中央は両側。
 * はみ出すのは型が文字列で値のある、結合の無い、折り返さないセルだけ。値のあるセル・結合で隠れるセル・結合の手前で止める。
 *
 * @param cells 行のセル（`GridRow.cells`）
 * @param k 対象の位置（`cells` の添字）
 * @param format 対象のセルの書式
 * @param lo 左へ数える下限（含む。描く列の左端）
 * @param hi 右へ数える上限（含まない。描く列の右端の次・表の列数）
 */
export function spillSides(
  cells: readonly (GridCell | null | undefined)[],
  k: number,
  format: CellFormat | undefined,
  lo = 0,
  hi: number = cells.length,
): SpillSides {
  const none = { left: 0, right: 0 };
  const cell = cells[k];
  if (
    cell === undefined ||
    cell === null ||
    cell.merged === true ||
    cell.colSpan > 1 ||
    cell.rowSpan > 1 ||
    cell.text === '' ||
    cell.kind !== 's' || // 型の無いセルははみ出さない（`grid.ts` の spillColumns と同じ）
    isWrapped(cell.kind, format)
  ) {
    return none;
  }
  const h = resolveHAlign(cell.kind, format);
  let left = 0;
  let right = 0;
  if (h !== 'left') {
    while (left < MAX_SPILL_COLS - 1 && k - left - 1 >= lo && isEmptyCell(cells[k - left - 1])) {
      left++;
    }
  }
  if (h !== 'right') {
    while (right < MAX_SPILL_COLS - 1 && k + right + 1 < hi && isEmptyCell(cells[k + right + 1])) {
      right++;
    }
  }
  return { left, right };
}

/**
 * 右寄せ・中央の文字列のはみ出しを描く（`grid-table.ts` から呼ぶ）。左寄せ・標準・折り返しのセルは何もせず false
 * （左寄せ・標準は呼び出し側の今の右へのはみ出しで描き、折り返しははみ出さない）。
 * - 入れ物を `margin-left` で左へ広げ、幅を左右の列幅の分だけ足す。中央は左右の余白の差で、文字の中心を元のセルの中心に合わせる。
 * - 文字を `span.xs-text` で包む（`style.css` の `.is-spill .xs-text` が文字の範囲だけ罫線を覆う）。
 * - 白以外の塗りが絡むなら `is-over-fill`（文字の下に面の色を敷かない）。列ごとの文字色の塗り分けはしない（ADR-0008 §3）。
 *
 * @param box セルの入れ物（`cellBox` の結果。文字は直下のテキスト）
 * @param row 描く行（範囲で切ったもの。`cells[k]` が対象）
 * @param k 対象の位置（`row.cells` の添字）
 * @param n 描く列の数（右へ数える上限）
 * @param geo 寸法
 * @param c 対象の列（0 始まり。geometry は c + 1）
 * @param table 書式の表（`XlsxSheet.formats`）
 * @param text セルの文字
 * @returns 呼び出し側の右へのはみ出しを止めるなら true（右寄せ・中央・折り返し）
 */
export function spillAround(
  box: HTMLDivElement,
  row: GridRow,
  k: number,
  n: number,
  geo: Geometry,
  c: number,
  table: readonly CellFormat[],
  text: string,
): boolean {
  const cell = row.cells[k];
  const fi = row.formats?.[k];
  const format = fi === undefined ? undefined : table[fi];
  if (!cell) return false;
  if (isWrapped(cell.kind, format)) return true;
  const h = resolveHAlign(cell.kind, format);
  if (!isText(cell.kind) || h === 'left') return false;
  const { left, right } = spillSides(row.cells, k, format, 0, n);
  const own = geo.colWidth(c + 1);
  if ((left === 0 && right === 0) || own === 0) return true;
  const leftW = geo.colLeft(c + 1) - geo.colLeft(c + 1 - left);
  const rightW = geo.colLeft(c + 2 + right) - geo.colLeft(c + 2);
  const label = document.createElement('span');
  label.className = 'xs-text';
  label.textContent = text;
  box.classList.add('is-spill');
  box.style.width = `${leftW + own + rightW - BORDER}px`;
  if (leftW > 0) box.style.marginLeft = `${-leftW}px`;
  if (h === 'center') {
    // 中央：内容の中心を元のセルの中心へ（左右の余白の差＝広げた幅の差）
    if (leftW > rightW) box.style.paddingLeft = `${PAD + leftW - rightW}px`;
    else if (rightW > leftW) box.style.paddingRight = `${PAD + rightW - leftW}px`;
  }
  box.replaceChildren(label);
  if (spillOverFill(table, row.formats, k - left, left + 1 + right)) {
    box.classList.add('is-over-fill');
  }
  return true;
}
