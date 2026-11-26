// 正: docs/adr/0011-selection.md §1・§2
/**
 * セル・行・列の選択（XSheet の追加。M7-24・Issue #55）。
 * - 選択は 1 つの範囲（行の全体・列の全体・セル）と、今のセル（キーボードで動かす起点）。
 * - 帯はスクロール枠の中の `div.xs-sel`（重ねる層）1 つで描く。表は見えている近くだけを描き直すので、`td` に印を付けると
 *   描き直しのたびに付け直しが要り、見えていない部分は塗れないため。位置は寸法（`Geometry`）から出す。
 *   重なりは罫線の層の後・図の層の前（セル・はみ出しの文字・ハイライト・罫線の上、図・sticky の見出しの下）。
 * - 選んだ行の行番号・列の列記号は、表を描くたびに `is-selected` を付け直す（描いた行・列だけ）。
 * - クリック：セル → 指定（`'row'` は行の全体、`'cell'` はセル）、行番号 → 行の全体、列記号 → 列の全体。
 *   結合したセルは結合の範囲。動いた距離が小さいクリックだけを選択にし、文字をドラッグで選んでいるときは何もしない。
 * DOM の要素は作るだけで、ユーザーのファイル内容は入れない。
 */
import type { CellRange, colName, Geometry, XlsxSheet } from '../core/index';

// このファイルは遅れて読む部品なので、表示の部品を import しない（単独版の xsheet-selection.mjs に二重に入れないため。ADR-0011 §3）
/** Excel の最大の行数・列数（grid.ts の `MAX_GRID_ROWS`・`MAX_GRID_COLS` と同じ） */
const MAX_GRID_ROWS = 1_048_576;
const MAX_GRID_COLS = 16_384;

/** セルをクリックしたときに選ぶ物（`RenderOptions.selection`） */
export type SelectionMode = 'row' | 'cell';

/** 選択の種類 */
export type SelectionType = 'row' | 'column' | 'cell';

/** 選択の変わった元（`'user'` はクリック・キー、`'api'` は `setSelection`・シートの切り替え） */
export type SelectionOrigin = 'user' | 'api';

/** 今の選択（XSheet: ADR-0011 §1） */
export interface Selection {
  readonly type: SelectionType;
  /**
   * 選んだ範囲（1 始まり・両端を含む）。行の全体は列 1〜16,384、列の全体は行 1〜1,048,576。
   * 結合したセルを選ぶと結合の範囲（行の全体なら結合の行の範囲）
   */
  readonly range: CellRange;
  /** 今のセル（結合なら左上）。行・列の全体を選んでいても、キーボードで動かす起点として列・行を覚えている */
  readonly active: { readonly r: number; readonly c: number };
}

/** `setSelection` に渡す選択（行・列の範囲は省くと 1 つ） */
export type SelectionInput =
  | { type: 'row'; r: number; r2?: number }
  | { type: 'column'; c: number; c2?: number }
  | { type: 'cell'; r: number; c: number };

/** 選択が見る表示の状態（`sheet-view.ts` が渡す。どれも呼んだ時点の値） */
export interface SelectionHost {
  /** スクロール枠（`div.xs-body`。クリックを受ける） */
  readonly body: HTMLElement;
  /** 帯を置く入れ物（`div.xs-wrap`）と、帯をその前に置く要素（図の層） */
  readonly wrapper: HTMLElement;
  readonly before: Element;
  sheet(): XlsxSheet | null;
  geometry(): Geometry | null;
  /** 行番号の列の幅・列見出しの高さ */
  origin(): { x: number; y: number };
  /** 表の行数・列数（帯はこの範囲で切る） */
  extent(): { rows: number; cols: number };
  table(): HTMLTableElement | null;
  /** 選択が変わった（同じ選択では呼ばない） */
  changed(selection: Selection | null, origin: SelectionOrigin): void;
  /** 列の記号（列記号の見出しを探す。selection-gate.ts が渡す） */
  readonly colName: typeof colName;
}

/** {@link createSelection} の戻り値 */
export interface SelectionController {
  get(): Selection | null;
  /**
   * 選択を替える（null で外す）。範囲は表の外でもよい（帯は表の範囲で切る）。
   * @throws RangeError 種類が正しくない・1 以上の整数でない・Excel の範囲を超える番号
   */
  set(input: SelectionInput | null, origin?: SelectionOrigin): void;
  /** 表を描き直した・寸法が変わった後に、帯の位置と見出しの印を合わせ直す */
  layout(): void;
  /** 今の選択を、知らせずに外す前に知らせる（シートの切り替え）。選択が無ければ何もしない */
  clear(): void;
  dispose(): void;
}

/** 結合を探すときに上下へ読む行数（これより上から始まる結合はここからの、これより高い結合はここまでの範囲として扱う） */
const MERGE_LOOKBACK = 100;

/** クリックとみなす、押してから離すまでに動いた距離（px） */
const CLICK_SLOP = 4;

/** 1 以上 max 以下の整数か */
const inRange = (v: unknown, max: number): v is number =>
  typeof v === 'number' && Number.isInteger(v) && v >= 1 && v <= max;

/**
 * 行 r・列 c を含む結合の範囲（無ければそのセル 1 つ）。シートの行を上へ辿って、結合の左上を探す。
 * `getRows` は範囲より上で始まる結合の左上を範囲の最初の行に置くので、`MERGE_LOOKBACK` 行より上で始まる結合は、そこからの範囲になる。
 */
export function mergeRangeAt(
  sheet: Pick<XlsxSheet, 'rowCount' | 'getRows'>,
  r: number,
  c: number,
): CellRange {
  const one = { r1: r, c1: c, r2: r, c2: c };
  if (r > sheet.rowCount) return one;
  // getRows は範囲の下端でも結合を縮めるので、下にも読む（結合の高さは下へ MERGE_LOOKBACK 行まで）
  const top = Math.max(1, r - MERGE_LOOKBACK + 1);
  const rows = sheet.getRows(top, Math.min(sheet.rowCount, r + MERGE_LOOKBACK - 1));
  for (let i = r - top; i >= 0; i--) {
    const row = rows[i]!;
    let k = c - 1;
    while (k >= 0 && row.cells[k] === null) k--;
    const cell = k >= 0 ? row.cells[k] : undefined;
    if (
      cell &&
      (cell.colSpan > 1 || cell.rowSpan > 1) &&
      k + cell.colSpan > c - 1 &&
      row.r + cell.rowSpan > r
    ) {
      return { r1: row.r, c1: k + 1, r2: row.r + cell.rowSpan - 1, c2: k + cell.colSpan };
    }
    if (row.cells[c - 1] !== null) break; // この行の列 c は隠れていない（これより上の結合は届かない）
  }
  return one;
}

/** 2 つの選択が同じか */
function sameSelection(a: Selection | null, b: Selection | null): boolean {
  if (a === null || b === null) return a === b;
  return (
    a.type === b.type &&
    a.range.r1 === b.range.r1 &&
    a.range.c1 === b.range.c1 &&
    a.range.r2 === b.range.r2 &&
    a.range.c2 === b.range.c2 &&
    a.active.r === b.active.r &&
    a.active.c === b.active.c
  );
}

/**
 * 選択を作る。`mode` が無ければクリックでは選ばない（`set` だけが効く）。
 * @param mode セルをクリックしたときに選ぶ物
 */
export function createSelection(
  host: SelectionHost,
  mode: SelectionMode | undefined,
): SelectionController {
  let value: Selection | null = null;
  const band = document.createElement('div');
  band.className = 'xs-sel';
  band.setAttribute('aria-hidden', 'true');
  band.hidden = true;
  /** 印を付けた見出しのセル（付け直す前に外す） */
  let marked: Element[] = [];

  /** 行 r1〜r2 の全体（結合の行の範囲に広げない） */
  const rowSel = (r1: number, r2: number, c: number): Selection => ({
    type: 'row',
    range: { r1: Math.min(r1, r2), c1: 1, r2: Math.max(r1, r2), c2: MAX_GRID_COLS },
    active: { r: Math.min(r1, r2), c },
  });
  /** 列 c1〜c2 の全体 */
  const colSel = (c1: number, c2: number, r: number): Selection => ({
    type: 'column',
    range: { r1: 1, c1: Math.min(c1, c2), r2: MAX_GRID_ROWS, c2: Math.max(c1, c2) },
    active: { r, c: Math.min(c1, c2) },
  });
  /** セル（結合なら結合の範囲）。asRow なら、その行の範囲の全体 */
  const cellSel = (r: number, c: number, asRow: boolean): Selection => {
    const sh = host.sheet();
    const m = sh ? mergeRangeAt(sh, r, c) : { r1: r, c1: c, r2: r, c2: c };
    const active = { r: m.r1, c: m.c1 };
    return asRow
      ? { type: 'row', range: { r1: m.r1, c1: 1, r2: m.r2, c2: MAX_GRID_COLS }, active }
      : { type: 'cell', range: m, active };
  };

  const update = (next: Selection | null, origin: SelectionOrigin): void => {
    if (sameSelection(value, next)) return;
    value = next;
    controller.layout();
    host.changed(value, origin);
  };

  /** 見出しの印を付け直す（描いた行番号・列記号のうち、選択に入るもの） */
  const markHeaders = (sel: Selection | null): void => {
    marked.forEach((el) => el.classList.remove('is-selected'));
    marked = [];
    const table = host.table();
    if (!sel || !table) return;
    const { r1, r2, c1, c2 } = sel.range;
    if (sel.type !== 'column') {
      for (const tr of table.querySelectorAll('tbody tr[aria-rowindex]')) {
        const r = Number(tr.getAttribute('aria-rowindex')) - 1;
        const head = tr.firstElementChild;
        if (r >= r1 && r <= r2 && head?.classList.contains('xs-rownum')) marked.push(head);
      }
    }
    if (sel.type !== 'row') {
      const names = new Set<string>();
      for (let c = c1; c <= Math.min(c2, host.extent().cols); c++) names.add(host.colName(c));
      for (const th of table.querySelectorAll('thead th:not(.xs-rownum)')) {
        if (names.has(th.textContent ?? '')) marked.push(th);
      }
    }
    marked.forEach((el) => el.classList.add('is-selected'));
  };

  /** 画面の位置（スクロール枠の表示の左上からの px）にある物を選ぶ */
  const pick = (x: number, y: number): Selection | null => {
    const g = host.geometry();
    const sh = host.sheet();
    if (!g || !sh) return null;
    const body = host.body;
    if (x < 0 || y < 0 || x >= body.clientWidth || y >= body.clientHeight) return null;
    const o = host.origin();
    const cx = body.scrollLeft + x - o.x;
    const cy = body.scrollTop + y - o.y;
    const inRows = cy >= 0 && cy < g.totalHeight;
    const inCols = cx >= 0 && cx < g.totalWidth;
    const prev = value?.active ?? { r: 1, c: 1 };
    if (y < o.y && x >= o.x && inCols) return colSel(g.colAt(cx), g.colAt(cx), prev.r);
    if (x < o.x && y >= o.y && inRows) return rowSel(g.rowAt(cy), g.rowAt(cy), prev.c);
    if (x >= o.x && y >= o.y && inRows && inCols && mode) {
      return cellSel(g.rowAt(cy), g.colAt(cx), mode === 'row');
    }
    return null;
  };

  let down: { x: number; y: number } | null = null;
  const onDown = (e: MouseEvent): void => {
    down = e.button === 0 ? { x: e.clientX, y: e.clientY } : null;
  };
  const onClick = (e: MouseEvent): void => {
    const start = down;
    down = null;
    if (!start || e.button !== 0) return;
    if (Math.hypot(e.clientX - start.x, e.clientY - start.y) > CLICK_SLOP) return;
    // 文字をドラッグ・ダブルクリックで選んでいる（コピーの邪魔をしない）
    const text = host.body.ownerDocument.getSelection?.();
    if (text && !text.isCollapsed && host.body.contains(text.anchorNode)) return;
    const rect = host.body.getBoundingClientRect();
    const sel = pick(
      e.clientX - rect.left - host.body.clientLeft,
      e.clientY - rect.top - host.body.clientTop,
    );
    if (sel) update(sel, 'user');
  };
  if (mode) {
    host.body.addEventListener('mousedown', onDown);
    host.body.addEventListener('click', onClick);
  }

  const controller: SelectionController = {
    get: () => value,
    set(input, origin = 'api') {
      if (input === null) {
        update(null, origin);
        return;
      }
      const prev = value?.active ?? { r: 1, c: 1 };
      let next: Selection;
      if (input.type === 'row') {
        const r2 = input.r2 ?? input.r;
        if (!inRange(input.r, MAX_GRID_ROWS) || !inRange(r2, MAX_GRID_ROWS)) {
          throw new RangeError('選択の行の番号が範囲外');
        }
        next = rowSel(input.r, r2, prev.c);
      } else if (input.type === 'column') {
        const c2 = input.c2 ?? input.c;
        if (!inRange(input.c, MAX_GRID_COLS) || !inRange(c2, MAX_GRID_COLS)) {
          throw new RangeError('選択の列の番号が範囲外');
        }
        next = colSel(input.c, c2, prev.r);
      } else if (input.type === 'cell') {
        if (!inRange(input.r, MAX_GRID_ROWS) || !inRange(input.c, MAX_GRID_COLS)) {
          throw new RangeError('選択のセルの番号が範囲外');
        }
        next = cellSel(input.r, input.c, false);
      } else {
        throw new RangeError('選択の種類が正しくない');
      }
      update(next, origin);
    },
    layout() {
      const sel = value;
      const g = host.geometry();
      markHeaders(sel);
      if (!sel || !g) {
        band.hidden = true;
        return;
      }
      const ext = host.extent();
      const { r1, c1 } = sel.range;
      const r2 = Math.min(sel.range.r2, ext.rows);
      const c2 = Math.min(sel.range.c2, ext.cols);
      const o = host.origin();
      const top = g.rowTop(r1);
      const left = g.colLeft(c1);
      const h = r1 > ext.rows ? 0 : g.rowTop(r2 + 1) - top;
      const w = c1 > ext.cols ? 0 : g.colLeft(c2 + 1) - left;
      if (h <= 0 || w <= 0) {
        band.hidden = true;
        return;
      }
      if (band.parentNode !== host.wrapper) host.wrapper.insertBefore(band, host.before);
      band.className = `xs-sel is-${sel.type}`;
      band.style.left = `${o.x + left}px`;
      band.style.top = `${o.y + top}px`;
      band.style.width = `${w}px`;
      band.style.height = `${h}px`;
      band.hidden = false;
    },
    clear() {
      update(null, 'api');
    },
    dispose() {
      host.body.removeEventListener('mousedown', onDown);
      host.body.removeEventListener('click', onClick);
      marked = [];
      value = null;
      band.remove();
    },
  };
  return controller;
}
