// 正: InterHub docs/plan/adr/0025-xlsx-own-reader.md
/**
 * シート 1 枚（sheetN.xml）の読み取り：行の位置の索引と、範囲の行だけの展開・結合セル。
 * - 開くときは `<row>` の位置の索引と、最後の行・列、結合セルだけを作る。セルは展開しない。
 * - 展開（getRows）は呼ばれた範囲の行だけを作って返し、展開した行を保持しない。
 * - 寸法（`<cols>`・`<sheetFormatPr>`・行の `ht`／`hidden`）は `layout` に疎な形で持つ（px への換算は geometry.ts）。
 *   行の寸法は、行の位置の索引を作る走査で同時に拾う（XML を 2 度読まない）。
 * - 名前空間の接頭辞付き XML（`<x:c>`）・`r` 属性の省略・`t="d"`・inlineStr・真偽・エラー・数式の保存済みの値を読む。
 */
import type { BorderTable } from './borders';
import { type FormatTable, isPaperFill } from './cell-format';
import type { CellFormatter } from './format';
import { isoToSerial } from './format';
import {
  MAX_CELL_TEXT,
  MAX_COLS,
  MAX_COL_WIDTH,
  MAX_ROWS,
  MAX_ROW_PT,
  type XlsxLimits,
} from './limits';
import { type SharedStrings, richTextToPlain } from './shared-strings';
import {
  type CellKind,
  type CellView,
  type ColLayout,
  type RowLayout,
  type SheetLayout,
  type SheetRow,
  XlsxError,
  type XlsxSheet,
  emptyLayout,
} from './types';
import {
  type Attrs,
  NS,
  colOfRef,
  decodeEntities,
  decodeText,
  findElement,
  forEachElement,
  isTagDelimiter,
  looksStrict,
  parseAttrs,
  parseCellRef,
} from './xml';

/** シートを読むのに要るもの（ブック全体で 1 つ） */
export interface SheetContext {
  strings: SharedStrings;
  formatter: CellFormatter;
  date1904: boolean;
  limits: XlsxLimits;
  /** 最大桁幅 MDW（px。styles.xml の先頭の font から。寸法の換算用） */
  mdw: number;
  /** 先頭の font の既定の行高（px・96dpi。実測表に無ければ null。寸法の換算用） */
  fontRowPx: number | null;
  /** 罫線の表（ブックで 1 つ。styles.xml の cellXfs の位置 → 罫線の番号。ADR-0029） */
  borderTable: BorderTable;
  /** 書式の表（ブックで 1 つ。styles.xml の cellXfs の位置 → 書式の番号。ADR-0030） */
  formatTable: FormatTable;
}

/** 開いたシート（getRows のほか、メモリを放す release を持つ） */
export interface ReleasableSheet extends XlsxSheet {
  release(): void;
}

/** 空のシート（グラフシートなど、セルを持たないもの） */
export function createEmptySheet(): ReleasableSheet {
  return {
    rowCount: 0,
    colCount: 0,
    layout: emptyLayout(),
    borders: [],
    formats: [],
    rowFormat: () => undefined,
    colFormat: () => -1,
    getRows: () => [],
    release() {},
  };
}

// セル本体の中の <v>・<is> の開きタグ（接頭辞を [1] に取る）。中身は閉じを indexOf で探して切り出す（bodyOf）
const V_OPEN = new RegExp(String.raw`<(${NS})v(?=[\s>/])[^<>]*>`);
const IS_OPEN = new RegExp(String.raw`<(${NS})is(?=[\s>/])[^<>]*>`);

/**
 * 断片の中で最初の要素（開きタグの正規表現 open で探す）の中身を返す。自己閉じ・閉じタグの無い要素・要素が無いときは undefined。
 * 閉じは `[\s\S]*?` ではなく `indexOf` で探すので、閉じの無い要素を並べた細工でも線形時間。
 */
function bodyOf(xml: string, open: RegExp, name: string): string | undefined {
  const m = open.exec(xml);
  if (m === null || m[0].endsWith('/>')) return undefined;
  const start = m.index + m[0].length;
  const close = xml.indexOf(`</${m[1] ?? ''}${name}>`, start);
  return close < 0 ? undefined : xml.slice(start, close);
}
const MERGE_RE = new RegExp(String.raw`<${NS}mergeCell\b[^<>]*?\sref\s*=\s*["']([^"']+)["']`, 'g');
const ROW_NUM_RE = /\sr\s*=\s*["'](\d+)["']/;
const CELL_COL_RE = /\sr\s*=\s*["']([A-Z]{1,3})\d*["']/;
const HT_RE = /\sht\s*=\s*["']([^"']*)["']/;
const HIDDEN_RE = /\shidden\s*=\s*["']([^"']*)["']/;
const FORMAT_PR_RE = new RegExp(String.raw`<${NS}sheetFormatPr\b[^<>]*>`);
const CUSTOM_RE = /\scustomHeight\s*=\s*["']([^"']*)["']/;
const COL_RE = new RegExp(String.raw`<${NS}col\b[^<>]*>`, 'g');
const CELL_S_RE = /\ss\s*=\s*["'](\d+)["']/;
const CUSTOM_FORMAT_RE = /\scustomFormat\s*=\s*["']([^"']*)["']/;
const SHEET_DATA_RE = new RegExp(String.raw`<(${NS})sheetData\b([^<>]*?)(/?)>`);

/** 結合セル 1 つ（1 始まり・両端を含む）。値の左上は (r1, c1) */
interface Merge {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/**
 * シートの XML から、シートを開く。
 * @throws XlsxError Strict 形式は unsupported、`<sheetData>` の閉じが無ければ broken、行・結合が多すぎれば too-large
 */
export function createSheet(xml: string, ctx: SheetContext): ReleasableSheet {
  if (looksStrict(xml)) throw new XlsxError('unsupported', 'strict');

  const sd = SHEET_DATA_RE.exec(xml);
  // 行の位置の索引（rowNums・starts・ends は同じ長さで、行番号の昇順）
  let rowNums: Int32Array = new Int32Array(0);
  let starts: Int32Array = new Int32Array(0);
  let ends: Int32Array = new Int32Array(0);
  let rowLayouts: RowLayout[] = [];
  // <sheetFormatPr>・<cols> は <sheetData> より前にある（無ければ全体から探す）
  const layoutBase = readLayoutHead(sd ? xml.slice(0, sd.index) : xml, ctx.mdw, ctx.fontRowPx);
  let dataEnd = 0;
  let prefix = '';
  if (sd) {
    prefix = sd[1] ?? '';
    const selfClosing = sd[3] === '/';
    const dataStart = sd.index + sd[0].length;
    if (selfClosing) {
      dataEnd = dataStart;
    } else {
      const close = xml.indexOf(`</${prefix}sheetData>`, dataStart);
      if (close < 0) throw new XlsxError('broken', 'sheetData');
      dataEnd = close;
      ({ rowNums, starts, ends, rowLayouts } = indexRows(
        xml,
        prefix,
        dataStart,
        dataEnd,
        layoutBase,
      ));
    }
  }

  let source: string | null = xml;
  const closeCell = `</${prefix}c>`;
  const openCell = `<${prefix}c`;
  const n = rowNums.length;

  // ---- セル 1 つを文字にする ----
  const cellText = (attrs: Attrs, body: string): string => {
    const t = attrs.t;
    let text: string;
    if (t === 'inlineStr') {
      const is = bodyOf(body, IS_OPEN, 'is');
      text = is === undefined ? '' : richTextToPlain(is);
    } else {
      const raw = bodyOf(body, V_OPEN, 'v');
      if (raw === undefined) return '';
      if (t === 's') {
        const idx = Number(raw);
        text = Number.isInteger(idx) ? ctx.strings.get(idx) : '';
      } else if (t === 'str') {
        text = decodeText(raw);
      } else if (t === 'b') {
        const v = raw.trim();
        text = v === '1' || v === 'true' ? 'TRUE' : 'FALSE';
      } else if (t === 'e') {
        text = decodeEntities(raw);
      } else if (t === 'd') {
        const iso = decodeEntities(raw);
        const parsed = isoToSerial(iso, ctx.date1904);
        text = parsed ? ctx.formatter.date(parsed.serial, xfOf(attrs), parsed.hasTime) : iso;
      } else {
        const v = raw.trim();
        if (v === '') return '';
        const num = Number(v);
        text = Number.isFinite(num) ? ctx.formatter.number(num, xfOf(attrs)) : decodeEntities(raw);
      }
    }
    return text.length > MAX_CELL_TEXT ? text.slice(0, MAX_CELL_TEXT) : text;
  };

  /** `<c t>` からセルの型を決める（t 無し・`n`・未知の値は数値） */
  const kindOf = (t: string | undefined): CellKind => {
    switch (t) {
      case 's':
      case 'str':
      case 'inlineStr':
        return 's';
      case 'b':
        return 'b';
      case 'e':
        return 'e';
      case 'd':
        return 'd';
      default:
        return 'n';
    }
  };

  /** セルの中身に数式（`<f>`）があるか */
  const hasFormula = (body: string): boolean => {
    const tag = `<${prefix}f`;
    for (let i = body.indexOf(tag); i >= 0; i = body.indexOf(tag, i + 1)) {
      if (isTagDelimiter(body[i + tag.length])) return true;
    }
    return false;
  };

  /**
   * 文字が空のセルが「値のあるセル」か（Excel は右隣への文字のはみ出しを止める）。次の 2 つだけ：
   * - 数式の結果が空文字（`t="str"` で `<f>` と `<v>` がある）
   * - 空の文字列を指す共有文字列（`t="s"` で番号が範囲内）
   * それ以外（値の要素が無い・数式の無い `t="str"` の空の `<v>`・空の inlineStr・文字列以外）は空のセル。
   */
  const isEmptyString = (attrs: Attrs, body: string): boolean => {
    const t = attrs.t;
    if (t === 's') {
      const raw = bodyOf(body, V_OPEN, 'v');
      if (raw === undefined) return false;
      const idx = Number(raw);
      return Number.isInteger(idx) && idx >= 0 && idx < ctx.strings.count;
    }
    if (t === 'str') return bodyOf(body, V_OPEN, 'v') !== undefined && hasFormula(body);
    return false;
  };

  /** 行の断片の中のセルを順に渡す。`r` が無い・不正なら、直前のセルの次の列とみなす（列を数えて補う） */
  const eachCell = (seg: string, cb: (col: number, attrs: Attrs, body: string) => void): void => {
    let prevCol = 0;
    // セル 1 つ（`<c r="A1" t="s"><v>0</v></c>` と `<c r="A1" s="3"/>`）。閉じは indexOf で探す（forEachElement）
    forEachElement(seg, 'c', (open, body) => {
      const attrs = parseAttrs(open);
      const ref = attrs.r;
      // r があるのに範囲外・不正な番地のセルは読まない。r が無いセルだけ、直前のセルの次の列とみなす
      const col = ref !== undefined ? colOfRef(ref) : prevCol + 1;
      if (col === 0 || col > MAX_COLS) return;
      prevCol = col;
      if (body === null) return; // 値のない（書式だけの）セル
      cb(col, attrs, body);
    });
  };

  /** i 番目の行（索引の位置）を cells へ展開する。cells の長さより右の列は捨てる */
  const expandRow = (i: number, cells: (CellView | null | undefined)[]): void => {
    if (source === null) return;
    const seg = source.slice(starts[i]!, ends[i]!);
    eachCell(seg, (col, attrs, body) => {
      if (col > cells.length) return;
      const text = cellText(attrs, body);
      if (text !== '') cells[col - 1] = { text, kind: kindOf(attrs.t), colSpan: 1, rowSpan: 1 };
      // 空文字の文字列（数式の結果が ""）は、見た目は空のまま、値のあるセルとして残す（右隣へのはみ出しを止める）
      else if (isEmptyString(attrs, body))
        cells[col - 1] = { text, kind: 's', colSpan: 1, rowSpan: 1 };
    });
  };

  // ---- 罫線（ADR-0029）・書式（ADR-0030）：値の無い書式だけのセルも、セルの開きタグの `s` から番号を引く ----
  const xfBorder = ctx.borderTable.xfBorder;
  const xfFormat = ctx.formatTable.xfFormat;
  const formatList = ctx.formatTable.formats;
  /** 罫線を持つ書式が 1 つでもあるか（無ければ、罫線のための走査をしない） */
  const hasBorders = ctx.borderTable.borders.length > 0;
  /** 見える書式が 1 つでもあるか（無ければ、書式のための走査も行・列の書式の解決もしない） */
  const hasFormats = formatList.length > 0;
  /** 書式の番号 → 紙でない塗りがあるか（表の範囲を広げる対象。装飾だけ・不透明な白の塗り＝紙では広げない） */
  const formatFilled = Uint8Array.from(formatList, (f) =>
    f.fill !== undefined && !isPaperFill(f.fill) ? 1 : 0,
  );
  const hasFilled = formatFilled.includes(1);
  /**
   * xf の番号 → 表の範囲を広げるか（罫線があるか、紙でない塗りの書式か）。
   * 行の断片にこれらの xf を参照する `s` が 1 つも無ければ、その行はセルの要素単位の走査をしない（lastStyledCol）。
   */
  const xfExtends = new Uint8Array(Math.max(xfBorder.length, xfFormat.length));
  for (let x = 0; x < xfExtends.length; x++) {
    const f = xfFormat[x] ?? -1;
    if ((xfBorder[x] ?? -1) >= 0 || (f >= 0 && formatFilled[f] === 1)) xfExtends[x] = 1;
  }
  /** `s` が無いセルは xf 0 の書式を引く。xf 0 が紙でない塗りのときだけ、`s` の有無で行を絞れない */
  const xf0Filled = (xfFormat[0] ?? -1) >= 0 && formatFilled[xfFormat[0]!] === 1;
  const sAttrRe = new RegExp(CELL_S_RE.source, 'g');
  /** 断片に、表の範囲を広げる xf を参照する `s` があるか（無ければ lastStyledCol の走査を省ける） */
  const mayExtend = (seg: string): boolean => {
    if (xf0Filled) return true;
    sAttrRe.lastIndex = 0;
    for (let m = sAttrRe.exec(seg); m !== null; m = sAttrRe.exec(seg))
      if (xfExtends[Number(m[1])] === 1) return true;
    return false;
  };
  const cellTagRe = new RegExp(String.raw`<${prefix}c(?=[\s>/])[^<>]*>`, 'g');

  /**
   * 行の断片の中のセルの開きタグを順に渡す（値のあるセルも自己閉じの書式だけのセルも。閉じは探さない）。
   * 列は `r` 属性、無い・不正なら直前の次（eachCell と同じ。範囲外の列は渡さない）。
   * 渡すのは、セルの要素 1 つにつき 1 回。border は罫線の番号（`s` が無い・cellXfs の範囲外・罫線なしは -1）、
   * format は書式の番号（`s` が無ければ xf 0。cellXfs の範囲外・書式なしは -1）。1 回の走査で両方を引く。
   */
  const eachStyledCell = (
    seg: string,
    cb: (col: number, border: number, format: number) => void,
  ): void => {
    let prevCol = 0;
    cellTagRe.lastIndex = 0;
    for (let m = cellTagRe.exec(seg); m !== null; m = cellTagRe.exec(seg)) {
      const tag = m[0];
      const ref = CELL_COL_RE.exec(tag)?.[1];
      const col = ref !== undefined ? colOfRef(ref) : prevCol + 1;
      if (col === 0 || col > MAX_COLS) continue;
      prevCol = col;
      const sm = CELL_S_RE.exec(tag);
      const xf = sm === null ? 0 : Number(sm[1]);
      cb(col, sm === null ? -1 : (xfBorder[xf] ?? -1), xfFormat[xf] ?? -1);
    }
  };

  /** i 番目の行の、罫線か塗りを持つ最後のセルの列（無ければ 0） */
  const lastStyledCol = (i: number): number => {
    let max = 0;
    const seg = source!.slice(starts[i]!, ends[i]!);
    if (!mayExtend(seg)) return 0;
    eachStyledCell(seg, (col, b, f) => {
      if (col > max && (b >= 0 || (f >= 0 && formatFilled[f] === 1))) max = col;
    });
    return max;
  };

  /**
   * 行の断片の開きタグから、行の書式（`customFormat="1"` の行の `s`）の書式の番号を返す。
   * `customFormat` が無い・偽なら undefined（行の書式なし）。`s` が無い・範囲外・書式なしの xf は -1（セルの要素の無い位置を書式なしにする）。
   */
  const rowFormatOf = (seg: string): number | undefined => {
    const gt = seg.indexOf('>');
    const tag = gt < 0 ? seg : seg.slice(0, gt + 1);
    if (!boolAttr(CUSTOM_FORMAT_RE.exec(tag)?.[1])) return undefined;
    const sm = CELL_S_RE.exec(tag);
    return xfFormat[sm === null ? 0 : Number(sm[1])] ?? -1;
  };

  // ---- 最後の行・列 ----
  /** i 番目の行に値のセルがあるか（まず `</c>` の有無で安く確かめ、あるときだけ実際に読む） */
  const rowHasText = (i: number): boolean => {
    const seg = source!.slice(starts[i]!, ends[i]!);
    if (seg.indexOf(closeCell) < 0) return false;
    let found = false;
    eachCell(seg, (_c, attrs, body) => {
      if (!found && cellText(attrs, body) !== '') found = true;
    });
    return found;
  };

  /** i 番目の行の、値のある最後のセルの列（無ければ 0）。多くの行は最後のセルの r 属性だけで決まる */
  const lastValueCol = (i: number): number => {
    const seg = source!.slice(starts[i]!, ends[i]!);
    const closeAt = seg.lastIndexOf(closeCell);
    if (closeAt < 0) return 0;
    let openAt = seg.lastIndexOf(openCell, closeAt);
    while (openAt >= 0 && !isTagDelimiter(seg[openAt + openCell.length]))
      openAt = seg.lastIndexOf(openCell, openAt - 1);
    if (openAt >= 0) {
      const gt = seg.indexOf('>', openAt);
      const m = CELL_COL_RE.exec(seg.slice(openAt, gt < 0 ? undefined : gt));
      if (m) {
        const col = colOfRef(m[1]!);
        if (col > 0) return col;
      }
    }
    // r が無い・不正なとき：行を読んで数える
    let max = 0;
    eachCell(seg, (col, attrs, body) => {
      if (col > max && cellText(attrs, body) !== '') max = col;
    });
    return max;
  };

  // 値のある最後の行（末尾の、書式だけで値の無い行は数えない）
  let lastRowIdx = -1;
  for (let i = n - 1; i >= 0; i--) {
    if (rowHasText(i)) {
      lastRowIdx = i;
      break;
    }
  }
  let rows = lastRowIdx >= 0 ? rowNums[lastRowIdx]! : 0;
  let cols = 0;
  for (let i = 0; i <= lastRowIdx; i++) {
    const c = lastValueCol(i);
    if (c > cols) cols = c;
  }

  // 罫線か塗りだけがあるセル（値の無い書式だけのセル）まで表を広げる（末尾の空の行に罫線・塗りだけがある帳票でも出るように）。
  // 装飾だけ（塗りなし）のセル・行と列の書式では広げない（何も見えない。ADR-0030 §決定 8）
  if (hasBorders || hasFilled) {
    for (let i = n - 1; i >= 0; i--) {
      const c = lastStyledCol(i);
      if (c === 0) continue;
      if (rowNums[i]! > rows) rows = rowNums[i]!;
      if (c > cols) cols = c;
    }
  }

  // ---- 結合セル ----
  let merges: Merge[] = [];
  if (sd && source.length > dataEnd) {
    const tail = source.slice(dataEnd);
    for (const m of tail.matchAll(MERGE_RE)) {
      const [a, b] = (m[1] ?? '').split(':');
      const p = parseCellRef(a ?? '');
      const q = b === undefined ? p : parseCellRef(b);
      if (!p || !q) continue;
      const r1 = Math.min(p.r, q.r);
      const r2 = Math.max(p.r, q.r);
      const c1 = Math.min(p.c, q.c);
      const c2 = Math.max(p.c, q.c);
      if (r1 === r2 && c1 === c2) continue; // 1 セルだけの結合は結合として扱わない
      if (merges.length >= ctx.limits.maxMerges) throw new XlsxError('too-large', 'merges');
      merges.push({ r1, c1, r2, c2 });
      if (r2 > rows) rows = r2;
      if (c2 > cols) cols = c2;
    }
  }

  const rowCount = rows;
  const colCount = Math.min(cols, MAX_COLS);

  /**
   * 列の書式（`<col style>`）→ 表の列ごとの書式の番号（長さ colCount。指定の無い列・書式なしは -1）。
   * 見える書式の列が 1 つも無ければ null。重なった指定は文書で先のものを採る。
   * 範囲は、まだ決まっていない列だけを数える（重なりが多くても列数にほぼ比例）。
   */
  const colFormats: Int32Array | null = (() => {
    if (!hasFormats || colCount === 0) return null;
    const out = new Int32Array(colCount).fill(-1);
    // next[c]：c 以降で最初の未決の列（1 始まり。colCount + 1 は番兵）。経路を縮めながら探す
    const next = Int32Array.from({ length: colCount + 2 }, (_v, i) => i);
    const find = (c0: number): number => {
      let r = c0;
      while (next[r] !== r) r = next[r]!;
      for (let c = c0; next[c] !== r;) {
        const t = next[c]!;
        next[c] = r;
        c = t;
      }
      return r;
    };
    let any = false;
    for (const cl of layoutBase.cols) {
      if (cl.xf === undefined) continue;
      const f = xfFormat[cl.xf] ?? -1;
      for (let c = find(cl.min); c <= Math.min(cl.max, colCount); c = find(c)) {
        out[c - 1] = f;
        next[c] = c + 1;
        if (f >= 0) any = true;
      }
    }
    return any ? out : null;
  })();

  /** rowNums の中で num 以上の最初の位置（無ければ n） */
  const lowerBound = (num: number): number => {
    let lo = 0;
    let hi = n;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (rowNums[mid]! < num) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };

  /** 範囲の外にある結合の左上の文字を取る（その行 1 本だけを読む） */
  const textAt = (r: number, c: number): string => {
    const tmp: (CellView | null | undefined)[] = new Array<CellView | undefined>(colCount).fill(
      undefined,
    );
    for (let i = lowerBound(r); i < n && rowNums[i] === r; i++) expandRow(i, tmp);
    return tmp[c - 1]?.text ?? '';
  };

  return {
    rowCount,
    colCount,
    layout: { ...layoutBase, rows: rowLayouts },
    borders: ctx.borderTable.borders,
    formats: formatList,
    rowFormat(r) {
      if (!hasFormats || source === null || !Number.isInteger(r) || r < 1) return undefined;
      const i = lowerBound(r);
      if (i >= n || rowNums[i] !== r) return undefined;
      return rowFormatOf(source.slice(starts[i]!, ends[i]!));
    },
    colFormat(c) {
      if (!hasFormats || !(c >= 1) || c > MAX_COLS) return -1;
      for (const cl of layoutBase.cols) {
        if (cl.xf !== undefined && cl.min <= c && c <= cl.max) return xfFormat[cl.xf] ?? -1;
      }
      return -1;
    },
    getRows(first, last) {
      const from = Math.max(1, Math.trunc(first));
      const to = Math.min(rowCount, Math.trunc(last));
      if (source === null || !(from <= to)) return [];

      const out: SheetRow[] = [];
      for (let r = from; r <= to; r++)
        out.push({ r, cells: new Array<CellView | undefined>(colCount).fill(undefined) });
      for (let i = lowerBound(from); i < n && rowNums[i]! <= to; i++) {
        expandRow(i, out[rowNums[i]! - from]!.cells);
      }

      // 罫線・書式の番号：値の無いセル・結合で隠れるセルも含む（結合の処理より前に、セルの `s` だけで決める）。1 回の走査で両方を引く
      if (hasBorders || hasFormats) {
        // 行の書式（customFormat の行の s）と、セルの要素のある位置の目印（行・列の書式を当てない位置）。out の位置で持つ
        const rowFormat: (number | undefined)[] = [];
        const present: (Uint8Array | undefined)[] = [];
        for (let i = lowerBound(from); i < n && rowNums[i]! <= to; i++) {
          const oi = rowNums[i]! - from;
          const row = out[oi]!;
          const seg = source.slice(starts[i]!, ends[i]!);
          let pres: Uint8Array | undefined = present[oi];
          if (hasFormats) {
            const rf = rowFormat[oi] ?? rowFormatOf(seg);
            if (rf !== undefined) rowFormat[oi] = rf;
            // 行の書式（見える書式のとき）か列の書式があるときだけ、セルの要素のある位置を覚える
            if (pres === undefined && (rf !== undefined ? rf >= 0 : colFormats !== null))
              pres = present[oi] = new Uint8Array(colCount);
          }
          eachStyledCell(seg, (col, b, f) => {
            if (col > colCount) return;
            if (pres !== undefined) pres[col - 1] = 1;
            if (b >= 0)
              (row.borders ??= new Array<number | undefined>(colCount).fill(undefined))[col - 1] =
                b;
            if (f >= 0)
              (row.formats ??= new Array<number | undefined>(colCount).fill(undefined))[col - 1] =
                f;
          });
        }
        // セルの要素の無い位置に、行の書式、無ければ列の書式を当てる（表の列数の中だけ。表の範囲は広げない）
        if (hasFormats) {
          for (let oi = 0; oi < out.length; oi++) {
            const rf = rowFormat[oi];
            if (rf !== undefined && rf < 0) continue; // 書式なしの行の書式は、列の書式も打ち消す
            if (rf === undefined && colFormats === null) continue;
            const row = out[oi]!;
            const pres = present[oi];
            for (let c = 0; c < colCount; c++) {
              if (pres !== undefined && pres[c] === 1) continue;
              const f = rf ?? colFormats![c]!;
              if (f >= 0)
                (row.formats ??= new Array<number | undefined>(colCount).fill(undefined))[c] = f;
            }
          }
        }
      }

      // 結合セル：左上だけ描き、ほかは隠す（左上以外に値が残っていても出さない）
      for (const mg of merges) {
        if (mg.r1 > to || mg.r2 < from) continue;
        const top = Math.max(mg.r1, from);
        const bottom = Math.min(mg.r2, to);
        const right = Math.min(mg.c2, colCount);
        const text =
          mg.r1 >= from ? (out[mg.r1 - from]!.cells[mg.c1 - 1]?.text ?? '') : textAt(mg.r1, mg.c1);
        for (let r = top; r <= bottom; r++) {
          const cells = out[r - from]!.cells;
          for (let c = mg.c1; c <= right; c++) cells[c - 1] = null;
        }
        out[top - from]!.cells[mg.c1 - 1] = {
          text,
          colSpan: right - mg.c1 + 1,
          rowSpan: bottom - top + 1,
        };
      }
      return out;
    },
    release() {
      source = null;
      rowNums = new Int32Array(0);
      starts = new Int32Array(0);
      ends = new Int32Array(0);
      merges = [];
    },
  };
}

/** 数として読める属性値（空・数でない・無限は undefined） */
function numAttr(v: string | undefined): number | undefined {
  if (v === undefined || v.trim() === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** 真偽の属性（`1`・`true` が真） */
const boolAttr = (v: string | undefined): boolean => v === '1' || v === 'true';

/**
 * シート XML の `<sheetData>` より前の部分から、`<sheetFormatPr>` と `<cols>` を読む（行の寸法はここでは空）。
 * 不正な値（負・数でない）は無いものとして既定にし、極端な値は Excel の上限（列幅 255 文字・行高 409.5pt）に丸める。
 * `<col>` は `min`〜`max` を展開せず範囲のまま持つ。`<cols>` が複数あっても順にすべて読む。
 */
function readLayoutHead(head: string, mdw: number, fontRowPx: number | null): SheetLayout {
  const layout: SheetLayout = { ...emptyLayout(mdw, fontRowPx) };

  const fp = FORMAT_PR_RE.exec(head);
  if (fp) {
    const a = parseAttrs(fp[0]);
    const dcw = numAttr(a.defaultColWidth);
    if (dcw !== undefined && dcw >= 0) layout.defaultColWidth = Math.min(dcw, MAX_COL_WIDTH);
    const bcw = numAttr(a.baseColWidth);
    if (bcw !== undefined && bcw >= 0) layout.baseColWidth = Math.min(bcw, MAX_COL_WIDTH);
    const drh = numAttr(a.defaultRowHeight);
    if (drh !== undefined && drh >= 0) layout.defaultRowHeight = Math.min(drh, MAX_ROW_PT);
    layout.customHeight = boolAttr(a.customHeight);
    layout.zeroHeight = boolAttr(a.zeroHeight);
  }

  const cols: ColLayout[] = [];
  scan: for (let pos = 0; pos < head.length;) {
    const el = findElement(head, 'cols', pos); // 閉じを indexOf で探す（線形）
    if (el === null) break;
    pos = el.end;
    if (el.body === null) continue; // 自己閉じ・閉じ無し
    for (const m of el.body.matchAll(COL_RE)) {
      if (cols.length >= MAX_COLS) break scan;
      const a = parseAttrs(m[0]);
      const min = numAttr(a.min);
      const max = numAttr(a.max);
      if (min === undefined || max === undefined) continue;
      if (!Number.isInteger(min) || !Number.isInteger(max)) continue;
      if (min < 1 || min > MAX_COLS || max < min) continue;
      const w = numAttr(a.width);
      cols.push({
        min,
        max: Math.min(max, MAX_COLS),
        width: w !== undefined && w >= 0 ? Math.min(w, MAX_COL_WIDTH) : null,
        hidden: boolAttr(a.hidden),
        ...styleOf(a.style),
      });
    }
  }
  layout.cols = cols;
  return layout;
}

/** `<col style>`（cellXfs の位置）を `ColLayout` に足す形にする。無い・負・整数でないなら何も足さない */
function styleOf(v: string | undefined): { xf?: number } {
  const n = numAttr(v);
  return n !== undefined && Number.isInteger(n) && n >= 0 ? { xf: n } : {};
}

/** セルの `s` 属性（cellXfs の位置）。無い・不正なら 0 */
function xfOf(attrs: Attrs): number {
  const s = Number(attrs.s);
  return Number.isInteger(s) && s >= 0 ? s : 0;
}

/**
 * `<row>` の位置の索引を作る（セルは見ない）。
 * 行番号は `r` 属性、無ければ直前の行の次。行番号が昇順でない文書は、番号で並べ替える。
 * 範囲外の行番号（Excel の上限 1,048,576 より大きい）の行は索引に入れない（ただし前の行の終わりの目印には使う）。
 * 同じ走査で、行の `ht`（pt）・`hidden`・`customHeight` も拾う（XML を 2 度読まない）。
 * 寸法の指定のある行だけ `rowLayouts` に入れる（既定と同じ行は持たない。`zeroHeight` のときは、`<row>` のあるすべての行。
 * 無い行を既定で非表示にするため）。
 * `rowLayouts` は行番号の昇順・重複なしで、連続する同じ指定の行は 1 件にまとめる（全行に同じ `ht` があっても 1 件）。
 * 同じ行番号が複数あるときは、文書で後の行を採る（後の行に指定が無ければ、前の行の指定も取り除く）。
 */
function indexRows(
  xml: string,
  prefix: string,
  dataStart: number,
  dataEnd: number,
  base: SheetLayout,
): { rowNums: Int32Array; starts: Int32Array; ends: Int32Array; rowLayouts: RowLayout[] } {
  const keepAll = base.zeroHeight;
  const rowOpen = `<${prefix}row`;
  const allNums: number[] = []; // 文書の順のすべての行の番号（範囲外は 0）
  const allPos: number[] = [];
  // 寸法の指定（文書の順）。allFlag: 0＝指定なし／1＝指定あり（＋2 hidden・＋4 customHeight）。allHt: pt（無ければ NaN）
  const allFlag: number[] = [];
  const allHt: number[] = [];
  let prev = 0;
  for (
    let p = xml.indexOf(rowOpen, dataStart);
    p >= 0 && p < dataEnd;
    p = xml.indexOf(rowOpen, p + rowOpen.length)
  ) {
    if (!isTagDelimiter(xml[p + rowOpen.length])) continue;
    const gt = xml.indexOf('>', p);
    const tag = xml.slice(p, gt < 0 || gt > dataEnd ? dataEnd : gt);
    const m = ROW_NUM_RE.exec(tag);
    const num = m ? parseInt(m[1]!, 10) : prev + 1;
    prev = num;
    if (allPos.length >= MAX_ROWS * 2) throw new XlsxError('too-large', 'rows');
    const inRange = num >= 1 && num <= MAX_ROWS;
    allNums.push(inRange ? num : 0);
    allPos.push(p);
    let flag = 0;
    let ht = NaN;
    if (inRange) {
      const raw = numAttr(HT_RE.exec(tag)?.[1]);
      const hasHt = raw !== undefined && raw >= 0;
      if (hasHt) ht = Math.min(raw, MAX_ROW_PT);
      const hidden = boolAttr(HIDDEN_RE.exec(tag)?.[1]);
      const custom = hasHt && boolAttr(CUSTOM_RE.exec(tag)?.[1]);
      // customHeight の無い ht がファイルの既定の行高と同じなら、既定と同じ行（持たない）
      const plain =
        !hidden && (!hasHt || (!custom && Math.abs(ht - base.defaultRowHeight) <= 0.01));
      if (keepAll || !plain) flag = 1 | (hidden ? 2 : 0) | (custom ? 4 : 0);
    }
    allFlag.push(flag);
    allHt.push(ht);
  }
  // 行の終わり＝次の行の始まり（文書の順）。最後の行は sheetData の終わり
  const order: number[] = [];
  let sorted = true;
  let last = 0;
  for (let i = 0; i < allNums.length; i++) {
    const num = allNums[i]!;
    if (num === 0) continue;
    if (num < last) sorted = false;
    last = num;
    order.push(i);
  }
  if (order.length > MAX_ROWS) throw new XlsxError('too-large', 'rows');
  if (!sorted) order.sort((a, b) => allNums[a]! - allNums[b]! || a - b);
  const rowNums = new Int32Array(order.length);
  const starts = new Int32Array(order.length);
  const ends = new Int32Array(order.length);
  const rowLayouts: RowLayout[] = [];
  let lastR = 0;
  order.forEach((src, i) => {
    const r = allNums[src]!;
    rowNums[i] = r;
    starts[i] = allPos[src]!;
    ends[i] = src + 1 < allPos.length ? allPos[src + 1]! : dataEnd;

    const tail = rowLayouts[rowLayouts.length - 1];
    if (r === lastR && tail !== undefined && tail.max === r) {
      // 同じ行番号は後の行を採る：直前の行の指定を取り除く
      if (tail.min === r) rowLayouts.pop();
      else tail.max = r - 1;
    }
    lastR = r;

    const flag = allFlag[src]!;
    if (flag === 0) return;
    const ht = Number.isNaN(allHt[src]!) ? null : allHt[src]!;
    const hidden = (flag & 2) !== 0;
    const custom = (flag & 4) !== 0;
    const t = rowLayouts[rowLayouts.length - 1];
    if (
      t !== undefined &&
      t.max === r - 1 &&
      t.ht === ht &&
      t.hidden === hidden &&
      t.custom === custom
    )
      t.max = r; // 連続する同じ指定の行は 1 件にまとめる
    else rowLayouts.push({ min: r, max: r, ht, hidden, custom });
  });
  return { rowNums, starts, ends, rowLayouts };
}
