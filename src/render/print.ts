// 正: docs/adr/0010-scroll-element-and-print.md（§2）
/**
 * 印刷の描画（`renderForPrint`・XSheet の追加。M7-20）。
 * 表示中のシートのデータの範囲を、用紙の中身の大きさのページに切り分けて、仮想化せずに描く。今の表示（倍率・位置）には触らない。
 * - ページは「下へ、それから右へ」の順（Excel の既定）。各ページは、その行・列の範囲を窓にした固定の表示：
 *   表（`buildGridTable`）・罫線の層（`svg.xs-borders` 1 枚）・図の層（範囲で切る入れ物の中）を重ね、ページの端で切る。
 *   ブラウザの改ページには任せない（幅のはみ出しは次のページへ送られず切れるため）。
 * - 範囲は、データ・図の右下・はみ出しの届く列（全部の行から）。枠を埋める空のセルは入れない。何も描く物の無いページは出さない。
 *   ファイルに印刷範囲（`_xlnm.Print_Area`）があれば、その範囲を定義の順に、範囲ごとにページを分けて描く（§3・M7-23）。
 * - 倍率はいつも 100%。`fit: 'width'` のときだけ、全部の列が 1 ページの幅に収まるまで縮める。
 * - 描くセルの数が上限を超えたら、超える手前の行で打ち切る（列は打ち切らない）。
 * - 画像・グラフを描き終えてから解決する。中断（`signal`・表示の destroy・シートの切り替え）では、足した物を片付けてから AbortError。
 * ユーザーのファイル内容は `textContent` だけで入れる（表・図の層の部品と同じ）。例外のメッセージにセル値・シート名を含めない。
 */
import type {
  CellRange,
  colName,
  createGeometry,
  Geometry,
  SheetDrawing,
  SheetRow,
  XlsxBook,
  XlsxSheet,
} from '../core/index';
import type { createBorderLayer } from './border-layer';
import type { buildBorderPaths } from './borders';
import type { drawingExtent, DrawingBox, PxRect, rectsIntersect } from './drawings';
import type { buildFormatCss } from './format-css';
import type { applyMarginFormats, collectVisibleRows, GridRow, rowNumWidth } from './grid';
import type { buildGridTable } from './grid-table';
import type { createImageLayer, ImageLayer } from './image-layer';
import type { XSheetMessages } from './messages';
import type { spillReachCols } from './spill-extent';
import type { zoomGeometry } from './zoom';

/**
 * 印刷の描画が使う、表示と共有の部品（print-gate.ts の `PRINT_KIT`）。
 * このファイルは遅れて読む部品なので、表示の部品は import せずに受け取る（単独版の xsheet-print.mjs に表示の部品を二重に入れないため。ADR-0010 §2）。
 */
export interface PrintKit {
  colName: typeof colName;
  createGeometry: typeof createGeometry;
  createBorderLayer: typeof createBorderLayer;
  buildBorderPaths: typeof buildBorderPaths;
  drawingExtent: typeof drawingExtent;
  rectsIntersect: typeof rectsIntersect;
  buildFormatCss: typeof buildFormatCss;
  applyMarginFormats: typeof applyMarginFormats;
  collectVisibleRows: typeof collectVisibleRows;
  rowNumWidth: typeof rowNumWidth;
  buildGridTable: typeof buildGridTable;
  createImageLayer: typeof createImageLayer;
  spillReachCols: typeof spillReachCols;
  zoomGeometry: typeof zoomGeometry;
  /** 見出し行の高さ（grid.ts の `GRID_HEADER_HEIGHT`） */
  GRID_HEADER_HEIGHT: number;
  /** Excel の最大の行数・列数（grid.ts の `MAX_GRID_ROWS`・`MAX_GRID_COLS`） */
  MAX_GRID_ROWS: number;
  MAX_GRID_COLS: number;
}

/** 印刷の指定（ADR-0010 §2-1） */
export interface PrintOptions {
  /**
   * 1 ページの中身の大きさ（CSS px。96px = 1in）。既定は A4 縦（210×297mm）から Excel の「標準」の余白
   * （左右 0.7in・上下 0.75in）を引いた 659×978。使う側の `@page` と同じ大きさにすること。
   */
  page?: { width: number; height: number };
  /** `'width'` は全部の列が 1 ページの幅に収まるように縮める（縦だけページを分ける）。既定は縮めない（100%） */
  fit?: 'width';
  /** 列記号・行番号を各ページに出す（既定 false） */
  headers?: boolean;
  /** 目盛り線（セルの境の薄い線）を出す（既定 false）。罫線はいつも出す */
  gridlines?: boolean;
  /** 描くセルの数の上限（行 × 列。非表示の行・列は数えない）。既定 200000 */
  maxCells?: number;
  /**
   * ファイルの印刷範囲（`_xlnm.Print_Area`）を使わず、データの範囲を印刷する（Excel の「印刷範囲を無視」）。
   * 既定 false（印刷範囲があれば、その範囲だけを範囲ごとにページを分けて描く）
   */
  ignorePrintArea?: boolean;
  /** 中断する（描いている間・画像を待つ間に見る） */
  signal?: AbortSignal;
}

/** 印刷の描画の結果（ADR-0010 §2-1） */
export interface PrintResult {
  /** 描いたページの数（表の無い表示・シートを開けていない・描く物の無いシートは 0） */
  readonly pages: number;
  /** 上限（`maxCells`）で行を打ち切った */
  readonly truncated: boolean;
  /** `target` に足した要素を外し、blob URL を放す。何度呼んでもよい */
  destroy(): void;
}

/** 既定の 1 ページの中身の大きさ（A4 縦・Excel の「標準」の余白。CSS px） */
export const PRINT_PAGE_DEFAULT: Readonly<{ width: number; height: number }> = {
  width: 659,
  height: 978,
};

/** 描くセルの数の上限の既定（仮の値。ADR-0010 §2-5） */
export const PRINT_MAX_CELLS_DEFAULT = 200_000;

/** 解いた印刷の指定 */
export interface ResolvedPrintOptions {
  page: { width: number; height: number };
  fit: boolean;
  headers: boolean;
  gridlines: boolean;
  maxCells: number;
  ignorePrintArea: boolean;
}

/**
 * 印刷の指定を解く。
 * @throws RangeError 用紙の大きさが正の有限の数でない・`fit` が `'width'` でない・`maxCells` が 1 以上の数でない
 */
export function resolvePrintOptions(options: PrintOptions = {}): ResolvedPrintOptions {
  const page = options.page ?? PRINT_PAGE_DEFAULT;
  const positive = (v: unknown): v is number =>
    typeof v === 'number' && Number.isFinite(v) && v > 0;
  if (!positive(page.width) || !positive(page.height)) {
    throw new RangeError('印刷の用紙の大きさが正しくない');
  }
  if (options.fit !== undefined && options.fit !== 'width') {
    throw new RangeError('印刷の fit は width だけ');
  }
  const maxCells = options.maxCells ?? PRINT_MAX_CELLS_DEFAULT;
  if (typeof maxCells !== 'number' || Number.isNaN(maxCells) || maxCells < 1) {
    throw new RangeError('印刷の maxCells は 1 以上の数');
  }
  return {
    page: { width: page.width, height: page.height },
    fit: options.fit === 'width',
    headers: options.headers === true,
    gridlines: options.gridlines === true,
    maxCells: Math.floor(maxCells),
    ignorePrintArea: options.ignorePrintArea === true,
  };
}

/** 1 ページの範囲（行・列は 1 始まり・両端を含む） */
export interface PrintPage {
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

/** 1 つの範囲のページの切れ目 */
export interface PrintAreaBands {
  /** 行のページの切れ目（上から） */
  rowBands: readonly [number, number][];
  /** 列のページの切れ目（左から） */
  colBands: readonly [number, number][];
}

/** 印刷の割り付け（DOM を使わない計算の結果） */
export interface PrintPlan {
  /** 描く最後の行（全部の範囲で一番下） */
  rows: number;
  /** 描く最後の列（全部の範囲で一番右） */
  cols: number;
  /** 倍率（`fit: 'width'` のときだけ 1 未満） */
  zoom: number;
  truncated: boolean;
  /** 範囲ごとの切れ目（印刷範囲の順。印刷範囲の無いシートはデータの範囲 1 つ）。範囲ごとにページを分ける */
  areas: readonly PrintAreaBands[];
}

/** 割り付けに使うシートの部分 */
export type PrintSheet = Pick<
  XlsxSheet,
  'layout' | 'rowCount' | 'colCount' | 'getRows' | 'formats'
>;

/** 範囲を読む行数の単位（はみ出しの届く列を数える読み取り） */
const READ_CHUNK = 512;

/**
 * 番号 start〜end を、長さ limit に収まる区切りに分ける（ページの切れ目）。切れ目は行・列の境。
 * 1 つだけで limit を超える番号は、その番号だけの区切りにする。長さ 0 の番号（非表示）だけの区切りは返さない。
 *
 * @param edge 番号 i の始まりの位置（`rowTop`・`colLeft`）
 * @param at 位置を含む番号（`rowAt`・`colAt`）
 */
export function splitBands(
  start: number,
  end: number,
  edge: (i: number) => number,
  at: (v: number) => number,
  limit: number,
): [number, number][] {
  const out: [number, number][] = [];
  let a = start;
  while (a <= end) {
    const top = edge(a);
    let b = Math.max(a, Math.min(end, at(top + limit)));
    while (b > a && edge(b + 1) - top > limit) {
      b--;
    }
    if (edge(b + 1) > top) {
      out.push([a, b]);
    }
    a = b + 1;
  }
  return out;
}

/** 印刷する範囲 1 つ（`spill` は、はみ出しの届く列まで右へ広げるか） */
interface AreaSpec extends PrintPage {
  spill: boolean;
}

/**
 * 印刷の範囲・倍率・ページの切れ目を決める（ADR-0010 §2-2・§2-5・§3）。
 * - 印刷範囲が無ければ、行は データ・図の最下行まで、列は データ・図の最右列・はみ出しの届く列（見える全部の行から）まで。
 * - 印刷範囲があれば、その範囲を定義の順に使う（範囲の外へのはみ出しは広げない）。列だけ（`$A:$C`）の範囲は
 *   データ・図の最下行まで、行だけ（`$1:$5`）の範囲はデータ・図の最右列・はみ出しの届く列までに縮める。
 * - 見える行を上から数え、全部の範囲を通して見える行 × 見える列が `maxCells` を超える手前で打ち切る
 *   （最初の 1 行は必ず描く。打ち切った後の範囲は描かない）。
 */
export function planPrint(
  kit: PrintKit,
  sheet: PrintSheet,
  boxes: readonly DrawingBox[],
  opts: ResolvedPrintOptions,
  printArea?: readonly CellRange[],
): PrintPlan {
  const { createGeometry, drawingExtent, spillReachCols, rowNumWidth, zoomGeometry } = kit;
  const { GRID_HEADER_HEIGHT, MAX_GRID_ROWS, MAX_GRID_COLS } = kit;
  const big = createGeometry(sheet.layout, MAX_GRID_ROWS, MAX_GRID_COLS);
  const fig = drawingExtent(big, boxes);
  const lastRow = Math.min(MAX_GRID_ROWS, Math.max(1, sheet.rowCount, fig.rows));
  const lastCol = Math.min(MAX_GRID_COLS, Math.max(1, sheet.colCount, fig.cols));
  const specs: AreaSpec[] =
    printArea && printArea.length > 0
      ? printArea.map((a) => {
          const r1 = Math.min(a.r1, MAX_GRID_ROWS);
          const c1 = Math.min(a.c1, MAX_GRID_COLS);
          const wholeRows = a.c2 >= MAX_GRID_COLS;
          return {
            r1,
            c1,
            r2: a.r2 >= MAX_GRID_ROWS ? Math.max(r1, lastRow) : a.r2,
            c2: wholeRows ? Math.max(c1, lastCol) : a.c2,
            spill: wholeRows,
          };
        })
      : [{ r1: 1, c1: 1, r2: lastRow, c2: lastCol, spill: true }];

  // 見える列の数（列 1〜n の分。先頭から数えた結果を覚えておく）
  const prefix = [0];
  const shownUpTo = (n: number): number => {
    for (let c = prefix.length; c <= n; c++) {
      prefix.push(prefix[c - 1]! + (big.colWidth(c) > 0 ? 1 : 0));
    }
    return prefix[n]!;
  };
  const shownCols = (c1: number, c2: number): number => shownUpTo(c2) - shownUpTo(c1 - 1);

  let used = 0;
  let truncated = false;
  const kept: PrintPage[] = [];
  for (const s of specs) {
    let cols = s.c2;
    let seen = 0;
    let end = s.r1 - 1;
    for (let a = s.r1; a <= s.r2 && !truncated; a += READ_CHUNK) {
      const b = Math.min(s.r2, a + READ_CHUNK - 1);
      const data: SheetRow[] =
        s.spill && a <= sheet.rowCount ? sheet.getRows(a, Math.min(b, sheet.rowCount)) : [];
      for (let r = a; r <= b; r++) {
        if (big.rowHeight(r) === 0) {
          end = r;
          continue;
        }
        const row = data[r - a];
        const reach = row ? Math.min(MAX_GRID_COLS, spillReachCols(big, [row], sheet.formats)) : 0;
        const need = Math.max(cols, reach);
        if ((used > 0 || seen > 0) && used + (seen + 1) * shownCols(s.c1, need) > opts.maxCells) {
          truncated = true;
          break;
        }
        seen++;
        cols = need;
        end = r;
      }
    }
    if (end >= s.r1) {
      kept.push({ r1: s.r1, r2: end, c1: s.c1, c2: cols });
    }
    used += seen * shownCols(s.c1, cols);
    if (truncated) break;
  }

  const rows = Math.max(1, ...kept.map((p) => p.r2));
  const cols = Math.max(1, ...kept.map((p) => p.c2));
  const headX = opts.headers ? rowNumWidth(rows) : 0;
  const headY = opts.headers ? GRID_HEADER_HEIGHT : 0;
  let zoom = 1;
  if (opts.fit) {
    const content = Math.max(0, ...kept.map((p) => big.colLeft(p.c2 + 1) - big.colLeft(p.c1)));
    const avail = Math.floor(opts.page.width - headX);
    if (avail > 0 && content > avail) {
      // 境目は zoomGeometry が丸めるので、右端が avail を超えないよう avail（整数）÷ 中身の幅にする
      zoom = avail / content;
    }
  }
  const g = zoomGeometry(createGeometry(sheet.layout, rows, cols), zoom);
  return {
    rows,
    cols,
    zoom,
    truncated,
    areas: kept.map((p) => ({
      rowBands: splitBands(
        p.r1,
        p.r2,
        (r) => g.rowTop(r),
        (y) => g.rowAt(y),
        opts.page.height - headY,
      ),
      colBands: splitBands(
        p.c1,
        p.c2,
        (c) => g.colLeft(c),
        (x) => g.colAt(x),
        opts.page.width - headX,
      ),
    })),
  };
}

/**
 * 寸法を 1 ページの窓にする：範囲の外の行・列は高さ・幅 0、位置は範囲の左上を 0 とする。
 * 表・罫線の部品にそのまま渡すと、範囲の外の余白を作らずに範囲だけを描く。
 */
export function windowGeometry(g: Geometry, p: PrintPage): Geometry {
  const y0 = g.rowTop(p.r1);
  const x0 = g.colLeft(p.c1);
  const clamp = (v: number, lo: number, hi: number): number => Math.min(hi, Math.max(lo, v));
  return {
    rowTop: (r) => (r <= p.r1 ? 0 : g.rowTop(Math.min(r, p.r2 + 1)) - y0),
    rowHeight: (r) => (r < p.r1 || r > p.r2 ? 0 : g.rowHeight(r)),
    colLeft: (c) => (c <= p.c1 ? 0 : g.colLeft(Math.min(c, p.c2 + 1)) - x0),
    colWidth: (c) => (c < p.c1 || c > p.c2 ? 0 : g.colWidth(c)),
    rowAt: (y) => clamp(g.rowAt(y + y0), p.r1, p.r2),
    colAt: (x) => clamp(g.colAt(x + x0), p.c1, p.c2),
    totalWidth: g.colLeft(p.c2 + 1) - x0,
    totalHeight: g.rowTop(p.r2 + 1) - y0,
  };
}

/** 範囲の列に、値・書式・罫線のあるセルがあるか（何も描く物の無いページを出さないため） */
function hasCellContent(rows: readonly GridRow[], c1: number, c2: number): boolean {
  for (const row of rows) {
    for (let c = c1; c <= c2; c++) {
      const cell = row.cells[c - 1];
      if (
        (cell && (cell.text !== '' || cell.colSpan > 1 || cell.rowSpan > 1)) ||
        row.formats?.[c - 1] !== undefined ||
        row.borders?.[c - 1] !== undefined
      ) {
        return true;
      }
    }
  }
  return false;
}

/** 印刷の描画に使う、表示中のシートの状態（sheet-view.ts が渡す） */
export interface PrintSource {
  /** `sheets` は印刷範囲（`SheetInfo.printArea`）を引くのに使う */
  book: Pick<XlsxBook, 'readImage' | 'readChart' | 'sheets'>;
  /** シートの番号（`SheetInfo.index`） */
  index: number;
  sheet: XlsxSheet;
  /** 図（図を出さない指定なら空） */
  drawings: readonly SheetDrawing[];
  /** 配置済みの図（A1 の左上基準・100% の px） */
  boxes: readonly DrawingBox[];
  messages: Readonly<XSheetMessages>;
  nonce?: string;
}

/** 書式の CSS の連番（印刷の分）。表示の連番（正の数）と重ならないよう負の数にする */
let printSeq = 0;

/** 中断の例外（メッセージにセル値・シート名を含めない） */
function abortError(): DOMException {
  return new DOMException('印刷の描画を中断しました', 'AbortError');
}

/** 次のタスクまで譲る（長い描画の間にも中断を受けられるように） */
const yieldTask = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

/** 今の時刻（ms） */
const now = (): number => (typeof performance !== 'undefined' ? performance.now() : Date.now());

/** 描く間に何 ms ごとに譲るか */
const YIELD_MS = 16;

/**
 * 表示中のシートを、印刷のためにページに分けて `target` の中に描く（ADR-0010 §2）。
 * @throws RangeError 正しくない指定
 * @throws DOMException AbortError（中断。足した物は片付けてから拒否する）
 */
export async function renderPrint(
  kit: PrintKit,
  target: HTMLElement,
  src: PrintSource | null,
  options: PrintOptions = {},
): Promise<PrintResult> {
  const { colName, createGeometry, createBorderLayer, buildBorderPaths, rectsIntersect } = kit;
  const { buildFormatCss, applyMarginFormats, collectVisibleRows, rowNumWidth } = kit;
  const { buildGridTable, createImageLayer, zoomGeometry } = kit;
  const { GRID_HEADER_HEIGHT, MAX_GRID_ROWS, MAX_GRID_COLS } = kit;
  const opts = resolvePrintOptions(options);
  const signal = options.signal;
  if (signal?.aborted) {
    throw abortError();
  }
  const empty: PrintResult = { pages: 0, truncated: false, destroy: () => undefined };
  if (!src) {
    return empty;
  }
  const sh = src.sheet;
  const area = opts.ignorePrintArea ? undefined : src.book.sheets[src.index - 1]?.printArea;
  const plan = planPrint(kit, sh, src.boxes, opts, area);
  const z = plan.zoom;
  const g = zoomGeometry(createGeometry(sh.layout, plan.rows, plan.cols), z);
  const big = createGeometry(sh.layout, MAX_GRID_ROWS, MAX_GRID_COLS);
  const headX = opts.headers ? rowNumWidth(plan.rows) : 0;
  const headY = opts.headers ? GRID_HEADER_HEIGHT : 0;

  const root = document.createElement('div');
  root.className = 'xs-sheet xs-print';
  if (opts.gridlines) {
    root.classList.add('is-gridlines');
  }
  root.style.setProperty('--xs-zoom', String(z));
  const seq = --printSeq;
  root.setAttribute('data-xs-fmt', String(seq));
  const fmtStyle = document.createElement('style');
  if (src.nonce !== undefined) {
    fmtStyle.setAttribute('nonce', src.nonce);
  }
  fmtStyle.textContent = buildFormatCss(sh.formats, seq);
  root.append(fmtStyle);

  const layers: ImageLayer[] = [];
  let destroyed = false;
  const destroy = (): void => {
    if (destroyed) return;
    destroyed = true;
    layers.forEach((l) => l.dispose());
    layers.length = 0;
    root.replaceChildren();
    root.remove();
  };
  const check = (): void => {
    if (signal?.aborted) {
      throw abortError();
    }
  };

  /** 行 a〜b の読み出し（データより下は空の行） */
  const readRows = (a: number, b: number): SheetRow[] => {
    const rows = a <= sh.rowCount ? sh.getRows(a, Math.min(b, sh.rowCount)) : [];
    for (let r = Math.max(a, sh.rowCount + 1); r <= b; r++) {
      rows.push({ r, cells: [] });
    }
    return rows;
  };

  /**
   * 1 ページの要素（表・罫線の層・図の層）
   * @param wg ページの窓の寸法（{@link windowGeometry}）
   * @param win 図の範囲（100% の px。A1 の左上基準）
   * @param figures ページに図があるか
   */
  const buildPage = (
    rows: GridRow[],
    wg: Geometry,
    p: PrintPage,
    win: PxRect,
    figures: boolean,
  ): HTMLDivElement => {
    const page = document.createElement('div');
    page.className = 'xs-print-page';
    page.style.width = `${opts.page.width}px`;
    page.style.height = `${opts.page.height}px`;
    const wrap = document.createElement('div');
    wrap.className = 'xs-wrap';
    wrap.append(
      buildGridTable(
        colName,
        p.c2,
        rows,
        wg,
        plan.rows,
        sh.layout.mdw * z,
        { first: p.c1, last: p.c2 },
        sh.formats,
        null,
        opts.headers,
      ),
    );
    if (sh.borders.length > 0) {
      const borders = createBorderLayer();
      borders.render(
        buildBorderPaths(rows, sh.borders, wg, { first: p.c1, last: p.c2 }),
        headX,
        headY,
      );
      wrap.append(borders.element);
    }
    if (figures) {
      // 図の層は 100% で描いて倍率を掛け、セルの範囲（見出しを除く）で切る
      const clip = document.createElement('div');
      clip.className = 'xs-print-clip';
      clip.style.left = `${headX}px`;
      clip.style.top = `${headY}px`;
      clip.style.width = `${wg.totalWidth}px`;
      clip.style.height = `${wg.totalHeight}px`;
      const layer = createImageLayer({ messages: src.messages });
      layers.push(layer);
      layer.setSheet({
        drawings: src.drawings,
        boxes: src.boxes,
        readImage: (path) => src.book.readImage(src.index, path),
        readChart: (path) => src.book.readChart(src.index, path),
      });
      layer.place(-win.x * z, -win.y * z, win.x + win.w, win.y + win.h);
      layer.element.style.transform = z === 1 ? '' : `scale(${z})`;
      layer.update(win);
      clip.append(layer.element);
      wrap.append(clip);
    }
    page.append(wrap);
    return page;
  };

  try {
    // ページは文書の外で組み、描き終えてから 1 回だけ取り付ける（取り付けたまま譲ると、譲るたびに
    // 増えていく木の全体のスタイル・レイアウトを計算し直し、200,000 セルで約 9 倍遅くなる）
    let pages = 0;
    let last = now();
    // 範囲の順に、範囲の中は下へ、それから右へ（Excel の既定）
    const order: PrintPage[] = plan.areas.flatMap((a) =>
      a.colBands.flatMap(([c1, c2]) => a.rowBands.map(([r1, r2]) => ({ r1, r2, c1, c2 }))),
    );
    for (const p of order) {
      const { r1, r2, c1, c2 } = p;
      const wg = windowGeometry(g, p);
      const read = collectVisibleRows(wg, r1, r2, r2, readRows);
      const rows =
        sh.formats.length > 0
          ? applyMarginFormats(
              read,
              c1,
              c2,
              sh.rowCount,
              sh.colCount,
              (r) => sh.rowFormat(r),
              (c) => sh.colFormat(c),
            )
          : read;
      const win: PxRect = {
        x: big.colLeft(c1),
        y: big.rowTop(r1),
        w: big.colLeft(c2 + 1) - big.colLeft(c1),
        h: big.rowTop(r2 + 1) - big.rowTop(r1),
      };
      const figures = src.boxes.some((b) => rectsIntersect(b.rect, win));
      if (!figures && !hasCellContent(rows, c1, c2)) {
        continue; // 何も描く物の無いページは出さない（Excel と同じ）
      }
      root.append(buildPage(rows, wg, p, win, figures));
      pages++;
      if (now() - last > YIELD_MS) {
        await yieldTask();
        check();
        last = now();
      }
    }
    check();
    if (pages === 0) {
      destroy(); // 描く物が無い（何も足さない）
      return { ...empty, truncated: plan.truncated };
    }
    target.append(root);
    await waitAbortable(settleAll(layers, root), signal);
    check();
    return { pages, truncated: plan.truncated, destroy };
  } catch (e) {
    destroy();
    throw e;
  }
}

/** 図の層の画像・グラフの読み込みを待ち、画像のデコードを待つ（壊れた画像は破線枠に替わる） */
async function settleAll(layers: readonly ImageLayer[], root: HTMLElement): Promise<void> {
  await Promise.all(layers.map((l) => l.settled()));
  const images = [...root.querySelectorAll('img')];
  await Promise.all(
    images.map((img) =>
      typeof img.decode === 'function' ? img.decode().catch(() => undefined) : undefined,
    ),
  );
}

/** p を待つ。先に中断されたら AbortError で拒否する */
function waitAbortable<T>(p: Promise<T>, signal: AbortSignal | undefined): Promise<T> {
  if (!signal) return p;
  return new Promise<T>((resolve, reject) => {
    const onAbort = (): void => reject(abortError());
    if (signal.aborted) {
      onAbort();
      return;
    }
    signal.addEventListener('abort', onAbort, { once: true });
    p.then(
      (v) => {
        signal.removeEventListener('abort', onAbort);
        resolve(v);
      },
      (e: unknown) => {
        signal.removeEventListener('abort', onAbort);
        reject(e instanceof Error ? e : new Error(String(e)));
      },
    );
  });
}
