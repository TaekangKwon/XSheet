// 正: docs/adr/0007-moobius-viewer-api.md §3・§4
/**
 * 今のヒットへの移動と、画面の位置の判定（M7-6・Moobius の依頼 R3・R4）。
 * - スクロール先は寸法（`Geometry`）からスクロールする前にすべて計算する（描いた後に測って直さない。ADR-0007 §3）。
 *   縦は「セルの下端から 1 行の高さを引いた位置」を表示枠の上から約 1/3 に置く（文字は下揃えなので、高い行・縦の結合でも文字が見える）。
 *   横は、対象が見えていないときだけ、見える最小の移動（右にあれば右端に寄せる・左にあれば左端に寄せる）にする。
 * - 位置の判定は寸法の `rowAt`・`colAt` で出し、DOM を見ない。結合は左上、はみ出した文字の上ははみ出し元のセルを、
 *   描いた行（`GridRow`）から引く。
 * 見出し（列見出しの高さ・行番号の列の幅）は引数でもらう（拡大縮小・見出しを隠すで変わるため。ADR-0007 §5・§7）。
 */
import type { Geometry, TextMatch, XlsxSheet } from '../core/index';
import { formatClassName } from './format-css';
import { spillColumns, type GridRow } from './grid';
import { hitIndexOf, markCurrent, type GridHits } from './highlight';

/** スクロール枠の今の状態と見出しの寸法（px） */
export interface ViewBox {
  scrollTop: number;
  scrollLeft: number;
  /** スクロール枠の表示の幅・高さ（スクロールバーを除く。`clientWidth`・`clientHeight`） */
  clientWidth: number;
  clientHeight: number;
  /** 行番号の列の幅（A1 の左端の x） */
  originX: number;
  /** 列見出しの高さ（A1 の上端の y） */
  originY: number;
}

/** 横の範囲（A1 の左端からの px。`right` は含まない） */
export interface XRange {
  left: number;
  right: number;
  /** 範囲を含むセルの左端（左へ動かすとき、範囲が収まる限りセルの左端から見せる。無ければ left） */
  cellLeft?: number;
}

/**
 * 縦のスクロール先。行 r から rowSpan 行（結合）の下端から、最後の見える行の高さを引いた位置を、表示枠のセルの部分の上から 1/3 に置く。
 * 表の下端より先へは動かさない。
 */
export function verticalTarget(
  geo: Pick<Geometry, 'rowTop' | 'rowHeight' | 'totalHeight'>,
  view: ViewBox,
  r: number,
  rowSpan = 1,
): number {
  const end = r + Math.max(1, rowSpan); // 結合の次の行
  let last = end - 1;
  while (last > r && geo.rowHeight(last) === 0) last--; // 最後の見える行（非表示の行は高さ 0）
  const y = geo.rowTop(end) - geo.rowHeight(last);
  const inner = Math.max(0, view.clientHeight - view.originY);
  const max = Math.max(0, view.originY + geo.totalHeight - view.clientHeight);
  return Math.min(max, Math.max(0, Math.round(y - inner / 3)));
}

/**
 * 横のスクロール先。範囲が表示枠のセルの部分に収まっていれば今のまま。
 * 右にはみ出せば右端を枠の右端に、左にはみ出せば（または枠より広ければ）左端を枠の左端に合わせる。
 * 左へ動かすときは、範囲が枠に収まる限りセルの左端（`cellLeft`）から見せる（セルの文字の頭が行番号の陰に入らない）。
 */
export function horizontalTarget(
  geo: Pick<Geometry, 'totalWidth'>,
  view: ViewBox,
  range: XRange,
): number {
  const inner = Math.max(0, view.clientWidth - view.originX);
  const from = view.scrollLeft;
  let to = from;
  if (range.left >= from && range.right <= from + inner) return from;
  if (range.right > from + inner) to = range.right - inner;
  if (range.left < to) {
    const cellLeft = Math.min(range.left, range.cellLeft ?? range.left);
    to = range.right - cellLeft <= inner ? cellLeft : range.left;
  }
  const max = Math.max(0, view.originX + geo.totalWidth - view.clientWidth);
  return Math.min(max, Math.max(0, Math.round(to)));
}

/** セル（列 c から colSpan 列）の横の範囲 */
export function cellXRange(geo: Pick<Geometry, 'colLeft'>, c: number, colSpan = 1): XRange {
  return { left: geo.colLeft(c), right: geo.colLeft(c + Math.max(1, colSpan)) };
}

/**
 * 画面の位置（スクロール枠の表示の左上からの px）にある行・列。見出しの上・表示枠の外・表の外は null。
 * 結合とはみ出しは見ない（{@link resolveDrawnCell} で直す）。
 */
export function pointToCell(
  geo: Pick<Geometry, 'rowAt' | 'colAt' | 'totalWidth' | 'totalHeight'>,
  view: ViewBox,
  x: number,
  y: number,
): { r: number; c: number } | null {
  if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
  if (x < view.originX || y < view.originY || x >= view.clientWidth || y >= view.clientHeight) {
    return null; // 見出し・スクロールバー・枠の外
  }
  const cx = view.scrollLeft + x - view.originX;
  const cy = view.scrollTop + y - view.originY;
  if (cx >= geo.totalWidth || cy >= geo.totalHeight) return null; // 表の外（右・下の余白）
  const r = geo.rowAt(cy);
  const c = geo.colAt(cx);
  return r >= 1 && c >= 1 ? { r, c } : null;
}

/** 空のセル（値も結合も無い）。`grid.ts` の判定と同じ（null＝結合で隠れるセルは空ではない） */
function isEmpty(cell: GridRow['cells'][number]): boolean {
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

/**
 * 描いた行から、行 r・列 c を覆う結合の左上か、c へはみ出す文字のセルを引く。どちらでもなければ (r, c)。
 * - 結合：同じ行の左、無ければ上の行を、結合で隠れるセル（null）の続く間だけ辿る。
 * - はみ出し：左の空のセルを辿り、最初の値のあるセルの文字が c まで届くか（`spillColumns`）を見る。
 *   描く列の範囲（firstCol）より左のセルは、描くときと同じくはみ出しの元にしない。
 * 描いた範囲より上で始まる結合は、描いた範囲の最初の行を左上として返す（描いた行には元の行が無いため）。
 *
 * @param rows 描いた行（見える行だけ。上から順。`cells[i]` は列 i + 1）
 * @param cols 表の列数（はみ出しを数える右端）
 * @param firstCol 描いた最初の列
 */
export function resolveDrawnCell(
  rows: readonly GridRow[],
  r: number,
  c: number,
  cols: number,
  firstCol = 1,
): { r: number; c: number } {
  const i0 = rows.findIndex((row) => row.r === r);
  if (i0 < 0) return { r, c };
  // 結合：上の行へ辿る
  for (let i = i0; i >= 0; i--) {
    const cells = rows[i]!.cells;
    let k = c - 1;
    while (k >= 0 && cells[k] === null) k--;
    const cell = k >= 0 ? cells[k] : undefined;
    if (i === i0 && k === c - 1) break; // 隠れていないセル
    if (cell && k + Math.max(1, cell.colSpan) > c - 1 && i + Math.max(1, cell.rowSpan) > i0) {
      return { r: rows[i]!.r, c: k + 1 };
    }
    if (cells[c - 1] !== null) break; // この行の列 c は隠れていない（これより上の結合は届かない）
  }
  // はみ出し
  const cells = rows[i0]!.cells;
  if (!isEmpty(cells[c - 1])) return { r, c };
  let k = c - 2;
  while (k >= firstCol - 1 && isEmpty(cells[k])) k--;
  if (k >= firstCol - 1 && k >= 0 && spillColumns(cells, k, cols) > c - 1 - k) {
    return { r, c: k + 1 };
  }
  return { r, c };
}

/**
 * 文字の幅を測る入れ物（画面の外・見えない）。セルと同じ書式のクラスを付けて測る（書式の規則は表示の根の `<style>`）。
 * 表示の根の中に置き、使い回す。
 */
export function createTextMeasure(root: HTMLElement): {
  width(text: string, formatClass: string): number;
  dispose(): void;
} {
  let table: HTMLTableElement | null = null;
  let td: HTMLTableCellElement | null = null;
  let span: HTMLSpanElement | null = null;
  return {
    width(text, formatClass) {
      if (!table || !td || !span) {
        table = document.createElement('table');
        table.className = 'xs-grid xs-measure';
        table.setAttribute('aria-hidden', 'true');
        td = table.createTBody().insertRow().insertCell();
        span = document.createElement('span');
        td.append(span);
        root.append(table);
      }
      td.className = formatClass;
      span.textContent = text;
      return span.getBoundingClientRect().width;
    },
    dispose() {
      table?.remove();
      table = td = span = null;
    },
  };
}

/** 移動・判定が見る表示の状態（`sheet-view.ts` が渡す。どれも呼んだ時点の値） */
export interface NavigationHost {
  /** スクロール枠（`div.xs-body`） */
  readonly body: HTMLElement;
  /** 表示の根（`div.xs-sheet`。測る入れ物を置く・`highlightAt` の範囲） */
  readonly root: HTMLElement;
  /** 表示中のシート（無ければ null） */
  sheet(): XlsxSheet | null;
  /** 表の寸法（まだ描いていなければ null） */
  geometry(): Geometry | null;
  /** 行番号の列の幅・列見出しの高さ */
  origin(): { x: number; y: number };
  /** 表の列数 */
  cols(): number;
  /** 描いた行と、描いた最初の列（まだ描いていなければ rows は空） */
  drawn(): { rows: readonly GridRow[]; firstCol: number };
  /** ハイライト（無ければ null） */
  hits(): GridHits | null;
  /** 描いた表（無ければ null） */
  table(): HTMLTableElement | null;
  /** 今のスクロール位置で描く（描いた範囲に収まっていれば何もしない） */
  draw(): void;
}

/** {@link createNavigation} の戻り値（`SheetView` の §3・§4 の操作） */
export interface Navigation {
  focusHighlight(index: number, options?: { scroll?: boolean; smooth?: boolean }): void;
  scrollToCell(r: number, c: number, options?: { smooth?: boolean }): void;
  cellAt(clientX: number, clientY: number): { r: number; c: number } | null;
  highlightAt(clientX: number, clientY: number): number;
  dispose(): void;
}

/** 結合の高さを調べるために、対象の行から読む行数（これより高い結合は、ここまでの高さとして扱う） */
const MERGE_READ_ROWS = 100;

/** セルの左右の余白（`style.css` の `.xs-cell` の padding-left） */
const CELL_PAD_LEFT = 3;

/**
 * 移動・判定（ADR-0007 §3・§4）を作る。状態は持たず、呼ぶたびに host から今の表示を引く（測る入れ物だけを持つ）。
 */
export function createNavigation(host: NavigationHost): Navigation {
  const measure = createTextMeasure(host.root);

  const viewBox = (): ViewBox => {
    const o = host.origin();
    return {
      scrollTop: host.body.scrollTop,
      scrollLeft: host.body.scrollLeft,
      clientWidth: host.body.clientWidth,
      clientHeight: host.body.clientHeight,
      originX: o.x,
      originY: o.y,
    };
  };

  /** 行 r・列 c のセル（結合の大きさ・型・書式の番号）。表の範囲の外は空のセル */
  const cellInfo = (r: number, c: number) => {
    const sh = host.sheet();
    if (!sh || r > sh.rowCount) return null;
    const row = sh.getRows(r, Math.min(sh.rowCount, r + MERGE_READ_ROWS - 1))[0];
    const cell = row?.cells[c - 1];
    return cell ? { cell, format: row!.formats?.[c - 1] } : null;
  };

  /** 計算したスクロール先へ 1 回だけ動かす（smooth でなければ、すぐに描く） */
  const scrollTo = (top: number, left: number, smooth: boolean): void => {
    const body = host.body;
    if (top === body.scrollTop && left === body.scrollLeft) return;
    if (smooth && typeof body.scrollTo === 'function') {
      body.scrollTo({ top, left, behavior: 'smooth' }); // 途中の描き直しはスクロールのイベントで行う
      return;
    }
    body.scrollTop = top;
    body.scrollLeft = left;
    host.draw();
  };

  /** 行 r・列 c（横は range）が見える所へのスクロール */
  const reveal = (r: number, range: XRange | null, c: number, smooth: boolean): void => {
    const geo = host.geometry();
    if (!geo) return;
    const view = viewBox();
    const info = cellInfo(r, c);
    const top = verticalTarget(geo, view, r, info?.cell.rowSpan ?? 1);
    const left = horizontalTarget(geo, view, range ?? cellXRange(geo, c, info?.cell.colSpan ?? 1));
    scrollTo(top, left, smooth);
  };

  /**
   * 一致の横の範囲。左揃えの文字（文字列）は、一致の前までと一致の終わりまでの文字の幅を、同じ書式で測って出す
   * （はみ出す文字の中の一致にも届く）。測れない（幅 0）・文字列でないときはセルの範囲。
   */
  const matchRange = (m: TextMatch): XRange | null => {
    const geo = host.geometry();
    const info = cellInfo(m.r, m.c);
    if (!geo || !info || info.cell.kind !== 's') return null;
    const text = info.cell.text;
    const cls = info.format === undefined ? '' : formatClassName(info.format);
    const a = measure.width(text.slice(0, Math.min(m.start, text.length)), cls);
    const b = measure.width(text.slice(0, Math.min(m.end, text.length)), cls);
    if (!(b > 0)) return null;
    const x = geo.colLeft(m.c) + CELL_PAD_LEFT;
    return { left: x + a, right: Math.ceil(x + b), cellLeft: geo.colLeft(m.c) };
  };

  return {
    focusHighlight(index, options = {}) {
      const hits = host.hits();
      const count = hits?.count ?? 0;
      if (!Number.isInteger(index) || index < -1 || index >= count) {
        throw new RangeError('ヒットの番号が範囲外');
      }
      if (!hits) return;
      hits.current = index;
      const table = host.table();
      if (table) markCurrent(table, index);
      const m = index >= 0 ? hits.match(index) : undefined;
      if (!m || options.scroll === false) return;
      reveal(m.r, matchRange(m), m.c, options.smooth === true);
    },
    scrollToCell(r, c, options = {}) {
      if (!Number.isInteger(r) || !Number.isInteger(c) || r < 1 || c < 1) {
        throw new RangeError('セルの番号が範囲外');
      }
      reveal(r, null, c, options.smooth === true);
    },
    cellAt(clientX, clientY) {
      if (!host.sheet()) return null;
      host.draw(); // 描き直しの前（スクロールの直後）でも、今の位置の描いた行で結合・はみ出しを引く
      const geo = host.geometry();
      if (!geo) return null;
      const rect = host.body.getBoundingClientRect();
      const at = pointToCell(
        geo,
        viewBox(),
        clientX - rect.left - host.body.clientLeft,
        clientY - rect.top - host.body.clientTop,
      );
      if (!at) return null;
      const d = host.drawn();
      return resolveDrawnCell(d.rows, at.r, at.c, host.cols(), d.firstCol);
    },
    highlightAt(clientX, clientY) {
      if (!host.hits() || this.cellAt(clientX, clientY) === null) return -1;
      const doc = host.root.ownerDocument;
      if (typeof doc.elementsFromPoint !== 'function') return -1;
      // いちばん上の、表示の中の要素（図・罫線の層は pointer-events: none なので拾われない）
      const top = doc.elementsFromPoint(clientX, clientY).find((el) => host.root.contains(el));
      return top ? hitIndexOf(top) : -1;
    },
    dispose() {
      measure.dispose();
    },
  };
}
