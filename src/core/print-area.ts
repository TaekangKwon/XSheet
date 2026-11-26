// 正: docs/adr/0010-scroll-element-and-print.md §3
/**
 * 印刷範囲（`workbook.xml` の `<definedName name="_xlnm.Print_Area" localSheetId="n">`）の読み取り（M7-22）。
 * - `workbook.ts`（InterHub の写し）は、読んだ workbook.xml を {@link readPrintAreas} に渡すだけ。
 * - 中身は `'Sheet 1'!$A$1:$H$50,'Sheet 1'!$J$1:$K$5` の形（シート名の引用・カンマで並ぶ複数の範囲）。
 *   列だけ（`$A:$C`）・行だけ（`$1:$5`）・1 セル（`$B$2`）も読む。`#REF!` などの読めない範囲は捨てる。
 * - シート名の部分は照らさない（Excel は `localSheetId` のシートを書く）。
 * - 範囲の数は 1 シート `maxPrintAreas` まで。超えた分は読まない。
 * DOM・Node 固有の API を使わない。
 */
import { MAX_COLS, MAX_ROWS } from './limits';
import type { CellRange } from './types';
import { colOfRef, decodeEntities, forEachElement, parseAttrs } from './xml';

/** 印刷範囲の定義の名前（大小を区別しない。Excel は `_xlnm.Print_Area` と書く） */
const PRINT_AREA = '_xlnm.print_area';

/**
 * 範囲 1 つの文字（シート名を除いた部分）を読む。`$` は無くてもよい。
 * `A1:C3`・`A1`・`A:C`・`1:3`。読めなければ null
 */
export function parseAreaRef(ref: string): CellRange | null {
  const s = ref.replace(/\$/g, '').trim().toUpperCase();
  const cell = /^([A-Z]{1,3})(\d{1,7})(?::([A-Z]{1,3})(\d{1,7}))?$/.exec(s);
  if (cell) {
    const c1 = colOfRef(cell[1]!);
    const r1 = parseInt(cell[2]!, 10);
    const c2 = cell[3] === undefined ? c1 : colOfRef(cell[3]);
    const r2 = cell[4] === undefined ? r1 : parseInt(cell[4], 10);
    return ordered(r1, c1, r2, c2);
  }
  const cols = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(s);
  if (cols) return ordered(1, colOfRef(cols[1]!), MAX_ROWS, colOfRef(cols[2]!));
  const rows = /^(\d{1,7}):(\d{1,7})$/.exec(s);
  if (rows) return ordered(parseInt(rows[1]!, 10), 1, parseInt(rows[2]!, 10), MAX_COLS);
  return null;
}

/** 範囲の向きをそろえる（`C3:A1` → `A1:C3`）。範囲外の行・列を含めば null */
function ordered(r1: number, c1: number, r2: number, c2: number): CellRange | null {
  const ok = (r: number, c: number) => r >= 1 && r <= MAX_ROWS && c >= 1 && c <= MAX_COLS;
  if (!ok(r1, c1) || !ok(r2, c2)) return null;
  return {
    r1: Math.min(r1, r2),
    c1: Math.min(c1, c2),
    r2: Math.max(r1, r2),
    c2: Math.max(c1, c2),
  };
}

/**
 * 定義の中身（実体参照を戻した後）をカンマで分ける。シート名の引用（`'a,b'!A1`・`''` は `'`）の中のカンマでは分けない。
 * 走査は 1 文字ずつの線形
 */
export function splitAreas(formula: string): string[] {
  const out: string[] = [];
  let start = 0;
  let quoted = false;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === "'") quoted = !quoted;
    else if (ch === ',' && !quoted) {
      out.push(formula.slice(start, i));
      start = i + 1;
    }
  }
  out.push(formula.slice(start));
  return out;
}

/** `'Sheet 1'!$A$1:$B$2` のシート名の部分を除く（最後の `!` より後。引用の外の `!` だけを見る） */
function stripSheet(part: string): string {
  let quoted = false;
  let bang = -1;
  for (let i = 0; i < part.length; i++) {
    const ch = part[i];
    if (ch === "'") quoted = !quoted;
    else if (ch === '!' && !quoted) bang = i;
  }
  return bang < 0 ? part : part.slice(bang + 1);
}

/**
 * workbook.xml の印刷範囲を、シートの位置（0 始まり・`<sheets>` の順＝`localSheetId`）ごとに返す。
 * 同じシートの定義が 2 つあれば先の物を使う。範囲が 1 つも読めないシートは入れない。
 */
export function readPrintAreas(
  wbXml: string,
  sheetCount: number,
  maxPrintAreas: number,
): Map<number, readonly CellRange[]> {
  const out = new Map<number, readonly CellRange[]>();
  forEachElement(wbXml, 'definedName', (open, body) => {
    if (body === null) return;
    const a = parseAttrs(open);
    if ((a.name ?? '').toLowerCase() !== PRINT_AREA) return;
    const id = Number(a.localSheetId);
    if (!Number.isInteger(id) || id < 0 || id >= sheetCount || out.has(id)) return;
    const ranges: CellRange[] = [];
    for (const part of splitAreas(decodeEntities(body))) {
      if (ranges.length >= maxPrintAreas) break;
      const r = parseAreaRef(stripSheet(part));
      if (r !== null) ranges.push(r);
    }
    if (ranges.length > 0) out.set(id, ranges);
  });
  return out;
}
