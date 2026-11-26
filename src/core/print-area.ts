// 正: docs/adr/0010-scroll-element-and-print.md §3
/**
 * 印刷範囲（`workbook.xml` の `<definedName name="_xlnm.Print_Area" localSheetId="n">`）の読み取り（M7-22）。
 * - `workbook.ts`（InterHub の写し）は、読んだ workbook.xml を {@link readPrintAreas} に渡すだけ。
 * - 中身は `'Sheet 1'!$A$1:$H$50,'Sheet 1'!$J$1:$K$5` の形（シート名の引用・カンマで並ぶ複数の範囲）。
 *   列だけ（`$A:$C`）・行だけ（`$1:$5`）・1 セル（`$B$2`）も読む。`#REF!` などの読めない範囲は捨てる。
 * - シート名の部分は照らさない（Excel は `localSheetId` のシートを書く）。
 * - 範囲の数は 1 シート `maxPrintAreas` まで。超えた分は読まない。
 * core の大きさ（単独版の上限）に効くので、処理は短く保つ。DOM・Node 固有の API を使わない。
 */
import { MAX_COLS, MAX_ROWS } from './limits';
import type { CellRange } from './types';
import { colOfRef, decodeEntities, forEachElement, parseAttrs } from './xml';

/**
 * 範囲 1 つの文字（シート名を除いた部分）を読む。`$` は無くてもよい。
 * `A1:C3`・`A1`・`A:C`（行の全体）・`1:3`（列の全体）。読めない・範囲外は null
 */
export function parseAreaRef(ref: string): CellRange | null {
  const m = /^([A-Z]{0,3})(\d{0,7})(?::([A-Z]{0,3})(\d{0,7}))?$/.exec(
    ref.replace(/\$/g, '').trim().toUpperCase(),
  );
  if (!m) return null;
  const [, a = '', b = '', c = a, d = b] = m;
  // 両端の形（列の有無・行の有無）がそろい、列か行の少なくとも一方がある
  if (!(a || b) || !a !== !c || !b !== !d) return null;
  const r1 = b ? +b : 1;
  const r2 = d ? +d : MAX_ROWS;
  const c1 = a ? colOfRef(a) : 1;
  const c2 = c ? colOfRef(c) : MAX_COLS;
  if (Math.min(r1, r2, c1, c2) < 1 || Math.max(r1, r2) > MAX_ROWS) return null;
  return {
    r1: Math.min(r1, r2),
    c1: Math.min(c1, c2),
    r2: Math.max(r1, r2),
    c2: Math.max(c1, c2),
  };
}

/**
 * 定義の中身（実体参照を戻した後）を、範囲ごとの文字（シート名の部分を除く）に分ける。
 * シート名の引用（`'a,b'!A1`・`''` は `'`）の中のカンマ・`!` では分けない。走査は 1 文字ずつの線形
 */
export function splitAreas(formula: string): string[] {
  const out: string[] = [];
  let start = 0;
  let quoted = false;
  for (let i = 0; i <= formula.length; i++) {
    const ch = formula[i];
    if (ch === "'") quoted = !quoted;
    else if (quoted) continue;
    else if (ch === '!') start = i + 1;
    else if (ch === ',' || ch === undefined) {
      out.push(formula.slice(start, i));
      start = i + 1;
    }
  }
  return out;
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
    const a = parseAttrs(open);
    const id = Number(a.localSheetId);
    if (
      body === null ||
      a.name?.toLowerCase() !== '_xlnm.print_area' ||
      !(id >= 0 && id < sheetCount && id % 1 === 0) ||
      out.has(id)
    )
      return;
    const ranges = splitAreas(decodeEntities(body))
      .map(parseAreaRef)
      .filter((r) => r !== null)
      .slice(0, maxPrintAreas);
    if (ranges.length > 0) out.set(id, ranges);
  });
  return out;
}
