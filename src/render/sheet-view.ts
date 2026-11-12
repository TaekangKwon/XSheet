// 正: InterHub docs/plan/adr/0025-xlsx-own-reader.md（§決定 4・6）・0026-xlsx-image-overlay.md（§決定 8）
/**
 * `.xlsx` の 1 シートの表示（InterHub `spreadsheet.ts` の `renderXlsxView`・`showSheet`・`draw`・`dispose` の .xlsx の部分。M2-3）。
 * - 見えている行・列（と前後の余分）だけを DOM にする行・列の仮想化グリッド。表は Excel の列幅・行高（geometry）で固定する。
 *   罫線は表の上の SVG 層、図（画像・図形・線）は表の上の図の層に、Excel と同じセルの位置で重ねる。
 * - 状態はすべてこの表示（インスタンス）が持つ。InterHub のモジュール変数（`renderGeneration`・`disposeXlsxView`）は無く、
 *   同じページに複数の表示を置ける。書式の CSS の連番（`data-xs-fmt`）だけはモジュールで 1 つの連番。
 * - 表示の根は自前の `div.xs-sheet`（使う側の `container` には属性を付けない）。子は通知の行（`div.xs-notice`）・
 *   スクロール枠（`div.xs-body`。仮想化の窓）・書式の `<style>`。枠の大きさは使う側が決める（`container` の高さ）。
 * - 書式の CSS はブックで 1 回だけ作り、シートを切り替えても作り直さない（`renderWorkbook` が同じ表示でシートを替える）。
 * - タブは持たない（`renderWorkbook` の仕事）。ブックも閉じない（開いた側が閉じる）。
 * ユーザーのファイル内容（セル値）は信頼できない入力として扱い、`innerHTML` は使わず `textContent` だけで DOM を組む。
 */
import {
  colName,
  createGeometry,
  XlsxError,
  type DrawingNotice,
  type Geometry,
  type SheetDrawings,
  type SheetRow,
  type XlsxBook,
  type XlsxSheet,
} from '../core/index';
import { createBorderLayer } from './border-layer';
import { buildBorderPaths } from './borders';
import { drawingExtent, layoutDrawings, viewWindow, type DrawingBox } from './drawings';
import { buildFormatCss } from './format-css';
import {
  applyMarginFormats,
  collectVisibleRows,
  GRID_HEADER_HEIGHT,
  gridExtent,
  MAX_GRID_COLS,
  MAX_GRID_ROWS,
  rowNumWidth,
  visibleColRange,
  visibleRowRange,
} from './grid';
import { buildGridTable } from './grid-table';
// XSheet: 検索の一致のハイライト（ADR-0007 §2）
import { createGridHits, type GridHits } from './highlight';
import type { TextMatch } from '../core/index';
// XSheet: 今のヒットへの移動・位置の判定（ADR-0007 §3・§4）
import { createNavigation, type Navigation } from './navigate';
import type { GridRow } from './grid';
// XSheet: 拡大縮小（ADR-0007 §5）
import {
  anchorScroll,
  contentWidth,
  fitZoom,
  resolveZoom,
  scrollAnchor,
  scrollbarWidth,
  unzoomRect,
  zoomGeometry,
  type Zoom,
  type ZoomSpec,
} from './zoom';
import { createImageLayer } from './image-layer';
import { formatMessage, type XSheetMessages } from './messages';
import { loadGeom } from './shapes';

/** 見えている範囲の上下に余分に描く行数（xrag-Viewer の値を踏襲）。 */
const OVERSCAN = 40;

/** 通知の行にする理由（図の読み取りの通知＋画像の合計の上限超過＋形の表の読み込み失敗） */
type ViewNotice = DrawingNotice | 'over-budget' | 'geom-failed';

/** 通知の理由 → 文言のキー */
const NOTICE_KEYS: Record<ViewNotice, keyof XSheetMessages> = {
  'part-too-large': 'noticeTooLarge',
  broken: 'noticeBroken',
  'too-many': 'noticeTooMany',
  'text-truncated': 'noticeTextTruncated',
  'shape-simplified': 'noticeShapeSimplified',
  'over-budget': 'noticeOverBudget',
  'geom-failed': 'libLoadFailed',
};

/** 書式 CSS の表示ごとの連番（`data-xs-fmt`）。表示を作るたびに 1 増やす（ファイル名・パスは使わない。ADR-0030 §決定 1） */
let formatSeq = 0;

/** {@link createSheetView} の指定（`RenderOptions` を解決したもの） */
export interface SheetViewOptions {
  /** 文言の表（`resolveMessages` の結果） */
  messages: Readonly<XSheetMessages>;
  /** 図の層を出すか（false なら `openDrawings` を呼ばず、図も通知も出さない） */
  drawings: boolean;
  /** 書式の `<style>` に付ける nonce（CSP 用） */
  nonce?: string;
  /** XSheet: 最初の倍率（既定 1。ADR-0007 §5） */
  zoom?: Zoom;
  /** XSheet: 行番号の列と列見出しを出すか（既定 true。ADR-0007 §7） */
  headers?: boolean;
}

/** 1 シートの表示（`renderSheet`・`renderWorkbook` の中身） */
export interface SheetViewInstance {
  /** 表示の根（`div.xs-sheet`）。`container` の子 */
  readonly root: HTMLDivElement;
  /** 表示中のシートの番号（`SheetInfo.index`）。開けなかった・まだ無いときは null */
  readonly activeIndex: number | null;
  /**
   * シートを開いて表示する（前のシートの図・blob URL は放す）。開けなければ枠の中にメッセージを出して終わる（例外にしない）。
   * 表示中に別の `showSheet`・`destroy` が来たら、この結果は捨てる。
   * @throws RangeError ブックに無い番号
   */
  showSheet(index: number): Promise<void>;
  /**
   * 検索の一致を光らせる（null・空の配列で消す）。描き直しても残り、シートを替えると消える。
   * XSheet: ADR-0007 §2
   */
  setHighlights(matches: readonly TextMatch[] | null): void;
  /** XSheet: ADR-0007 §3・§4（中身は navigate.ts） */
  focusHighlight: Navigation['focusHighlight'];
  scrollToCell: Navigation['scrollToCell'];
  cellAt: Navigation['cellAt'];
  highlightAt: Navigation['highlightAt'];
  /**
   * XSheet: 倍率を替える（見ていた左上のセルを左上に残す。ADR-0007 §5）
   * @throws RangeError 数でない倍率・正しくない指定
   */
  setZoom(zoom: Zoom): void;
  /** XSheet: 今の倍率（fit のときは計算した値） */
  readonly zoom: number;
  /** リスナ・ResizeObserver・rAF・blob URL・シートを放し、根を `container` から外す。何度呼んでもよい。ブックは閉じない */
  destroy(): void;
}

/**
 * 読み取りの例外を種類ごとの文言にする（ADR-0025 決定 6）。
 * not-zip・broken やそれ以外の例外は「解析に失敗しました」に詳細を添える。
 * 詳細（XlsxError のメッセージ）にはセル値・シート名を含まない。
 */
export function xlsxErrorMessage(e: unknown, messages: Readonly<XSheetMessages>): string {
  if (e instanceof XlsxError) {
    switch (e.kind) {
      case 'encrypted':
        return messages.encrypted;
      case 'too-large':
        return messages.tooLarge;
      case 'unsupported':
        return messages.unsupported;
      default:
        break;
    }
  }
  const detail = e instanceof Error ? e.message : String(e);
  return formatMessage(messages.parseFailed, { detail });
}

/** データ無しの表示（空のシート）。 */
function emptyView(message: string): HTMLElement {
  const div = document.createElement('div');
  div.className = 'xs-empty';
  div.textContent = message;
  return div;
}

/** 警告（シートを開けない）の要素を作る。 */
function bannerElement(message: string): HTMLElement {
  const banner = document.createElement('div');
  banner.className = 'xs-banner';
  banner.textContent = message;
  return banner;
}

/** 形を評価する図（形の表が要る図）があるか */
function needsGeom(found: SheetDrawings): boolean {
  return found.drawings.some((d) => d.geom !== undefined);
}

/**
 * 1 シートの表示を作り、根を `container` に取り付ける（シートは {@link SheetViewInstance.showSheet} で開く）。
 * 保持するのは表示中の 1 シートだけ。シートを替えると前のシートを放す（ADR-0025 決定 4）。
 *
 * @param container 表示を置く要素（高さが決まっていること）
 * @param book 開いたブック（閉じるのは呼び出し側）
 */
export function createSheetView(
  container: HTMLElement,
  book: XlsxBook,
  options: SheetViewOptions,
): SheetViewInstance {
  const { messages } = options;
  const root = document.createElement('div');
  root.className = 'xs-sheet';

  // 書式（塗り・文字の装飾）の CSS（ISSUE-100・ADR-0030 §決定 1）：表示の根に連番の属性と <style> 1 枚を置く。
  // 書式はブック単位なので、最初に開いたシートで 1 回だけ作り、シートの切替では作り直さない（中身は textContent だけで入れる）。
  // 表示の根の子のため、表示を消せば規則も消える（destroy でも取り除く）。
  const fmtSeq = ++formatSeq;
  root.setAttribute('data-xs-fmt', String(fmtSeq));
  const fmtStyle = document.createElement('style');
  if (options.nonce !== undefined) {
    fmtStyle.setAttribute('nonce', options.nonce);
  }
  let fmtReady = false;

  /** スクロール枠（仮想化の窓）。大きさは CSS（`container` の高さ）で決まり、ここでは測るだけ */
  const body = document.createElement('div');
  body.className = 'xs-body';

  // 表示中のシートとグリッドの状態。シートは 1 枚だけ強く持つ（読み取り側が捨てたシートの XML を放せるように）。
  let sheet: XlsxSheet | null = null;
  let activeIndex: number | null = null;
  /** 表示中のシートの寸法（Excel の列幅・行高を px にしたもの）。シートを開くたび・表の範囲が変わるたびに作り直す */
  let geo: Geometry | null = null;
  /** 表の行数・列数（データより広げて、スクロール枠を埋める）。`geo` はこの範囲で作る */
  let extent = { rows: 0, cols: 0 };
  /** 描いた行・列の範囲（描き直しの判定に使う） */
  let drawn: { first: number; last: number; firstCol: number; lastCol: number } | null = null;
  /** 新しい showSheet・destroy が来たら、前のシート読み込みの結果を捨てるための番号。 */
  let openSeq = 0;
  let frame = 0;
  let initFrame = 0;
  let closed = false;
  // XSheet: 検索の一致のハイライト（描くたびに付け直す。シートを替えると消す。ADR-0007 §2）
  let hits: GridHits | null = null;
  // XSheet: 描いた行（位置の判定で結合・はみ出しを引く。ADR-0007 §4）
  let drawnRows: GridRow[] = [];
  // XSheet: 倍率（ADR-0007 §5）。zoomSpec は指定、zoom は今の値（fit は draw が合わせる）。contentW は fit の中身の幅（100%）
  let zoomSpec: ZoomSpec = resolveZoom(options.zoom ?? 1);
  let zoom = 'fixed' in zoomSpec ? zoomSpec.fixed : 1;
  let contentW = 0;
  /** 倍率を替えている間（中の draw で幅に合わせ直さない） */
  let zooming = false;
  // XSheet: 見出しを隠すときは、行番号の列の幅・列見出しの高さを 0 として計算する（ADR-0007 §7）
  const showHeaders = options.headers !== false;
  const headY = showHeaders ? GRID_HEADER_HEIGHT : 0;
  const headX = (rows: number): number => (showHeaders ? rowNumWidth(rows) : 0);
  root.style.setProperty('--xs-zoom', String(zoom));

  // --- 図の層（ISSUE-86・ADR-0026 §決定 8）---
  // スクロール枠の中の常駐の wrapper（position: relative）に、表（描き直しのたびに差し替える）と図の層（position: absolute）を並べる。
  const wrapper = document.createElement('div');
  wrapper.className = 'xs-wrap';
  const layer = createImageLayer({ messages });
  // 罫線の層（ISSUE-98・ADR-0029）：表の直後・図の層の直前（DOM の順で、セルの上・図の下・sticky の見出しの下を通る）。
  // 表（wrapper の先頭）は描き直しのたびに prepend／replaceWith されるため、この 2 つは常駐のまま順序が保たれる。
  const borderLayer = createBorderLayer();
  wrapper.append(borderLayer.element, layer.element);
  /** wrapper の中の今の表（無ければ null） */
  let tableEl: HTMLTableElement | null = null;
  /** 表示中のシートの図の配置（A1 の左上基準 px。大きさ 0 の図は含まない） */
  let boxes: DrawingBox[] = [];
  /** 図を収めるのに必要な表の行数・列数（表を図のある範囲まで広げる） */
  let drawingMin = { rows: 0, cols: 0 };
  /** 図の読み取り・上限超過をまとめて知らせる 1 行（無ければ hidden） */
  const noticeBar = document.createElement('div');
  noticeBar.className = 'xs-notice';
  noticeBar.hidden = true;
  /** 表示中のシートで知らせる理由 */
  const notices = new Set<ViewNotice>();

  /** 通知の行の文言を notices に合わせる（理由の順に 1 行で並べる） */
  const renderNotices = (): void => {
    const text = [...notices].map((n) => messages[NOTICE_KEYS[n]]).join(' ');
    noticeBar.textContent = text;
    noticeBar.hidden = text === '';
  };

  /** 表も図も出さず、メッセージ（空・エラー）だけを本文枠に出す。前のシートの blob URL はここで revoke する */
  const showMessage = (el: HTMLElement): void => {
    layer.setSheet(null);
    boxes = [];
    drawingMin = { rows: 0, cols: 0 };
    tableEl = null;
    borderLayer.clear();
    wrapper.replaceChildren(borderLayer.element, layer.element);
    body.replaceChildren(el);
  };

  /**
   * 表の範囲（行数・列数）を、スクロール枠の大きさに合わせて決め直す。データより広げて空セルを足し、画面を埋める
   * （Excel のシートのように、右端の列の文字もはみ出せる）。範囲が変わったら geometry を作り直して true を返す。
   * 図のある範囲まで広げる（ISSUE-86）ため、図の最下行・最右列を gridExtent の minRows・minCols に渡す。
   * 枠の大きさは使う側が決める（InterHub はウィンドウ高に合わせた max-height も見ていたが、ここでは clientWidth・clientHeight だけ）。
   */
  const refreshExtent = (sh: XlsxSheet): boolean => {
    const next = gridExtent(
      sh.layout,
      sh.rowCount,
      sh.colCount,
      body.clientWidth / zoom, // XSheet: 倍率を掛けた寸法で枠を埋める（ADR-0007 §5）
      body.clientHeight / zoom,
      drawingMin.rows,
      drawingMin.cols,
    );
    if (geo && next.rows === extent.rows && next.cols === extent.cols) {
      return false;
    }
    extent = next;
    geo = zoomGeometry(createGeometry(sh.layout, next.rows, next.cols), zoom); // XSheet: ADR-0007 §5
    return true;
  };

  /** 行 a〜b の読み出し。データより下（足した空の行）は空の行にする。 */
  const readRows = (sh: XlsxSheet, a: number, b: number): SheetRow[] => {
    const rows = a <= sh.rowCount ? sh.getRows(a, Math.min(b, sh.rowCount)) : [];
    for (let r = Math.max(a, sh.rowCount + 1); r <= b; r++) {
      rows.push({ r, cells: [] });
    }
    return rows;
  };

  /**
   * 見えている行（と上下の余分）を描く。force でなければ、描いた範囲に見えている行が収まる間は描き直さない。
   * 行の位置・上下の余白は geometry（Excel の寸法）から出し、描いた行の高さは実測しない。
   * 非表示の行は DOM に出さず、前後の余分も見える行の数で数える（フィルターで大半を畳んだシートでも描く行は一定）。
   */
  const draw = (force = false): void => {
    const sh = sheet;
    if (!sh || closed) {
      return;
    }
    // XSheet: 幅に合わせる倍率を、枠の大きさが変わるたびに合わせ直す（ADR-0007 §5）
    const fitted = fitTarget();
    if (fitted !== null) {
      applyZoom(fitted);
      return;
    }
    const extended = refreshExtent(sh);
    const g = geo;
    if (!g) {
      return;
    }
    const range = visibleRowRange(g, body.scrollTop, body.clientHeight, extent.rows, OVERSCAN);
    const d = drawn;
    const originX = headX(extent.rows); // XSheet: ADR-0007 §7
    // 列も仮想化する：描く列は可視範囲の左右 1 画面分。見える列がその外へ出たら描き直す
    const cols = visibleColRange(g, body.scrollLeft, body.clientWidth, originX, extent.cols);
    if (
      force ||
      extended ||
      !d ||
      range.top < d.first ||
      range.bottom > d.last ||
      cols.left < d.firstCol ||
      cols.right > d.lastCol
    ) {
      const read = collectVisibleRows(g, range.first, range.last, extent.rows, (a, b) =>
        readRows(sh, a, b),
      );
      // 表の範囲の外の余白（足した空の行・列）にも、行の書式 → 列の書式の順で塗りの書式の番号を当てる（見える書式のあるブックだけ）
      const rows =
        sh.formats.length > 0
          ? applyMarginFormats(
              read,
              cols.first,
              cols.last,
              sh.rowCount,
              sh.colCount,
              (r) => sh.rowFormat(r),
              (c) => sh.colFormat(c),
            )
          : read;
      const table = buildGridTable(
        colName,
        extent.cols,
        rows,
        g,
        extent.rows,
        sh.layout.mdw * zoom, // XSheet: 収まる数値の判定も倍率で（ADR-0007 §5）
        cols,
        sh.formats,
        hits, // XSheet: ADR-0007 §2
        showHeaders, // XSheet: ADR-0007 §7
      );
      if (wrapper.parentNode !== body) {
        body.replaceChildren(wrapper);
      }
      if (tableEl) {
        tableEl.replaceWith(table);
      } else {
        wrapper.prepend(table);
      }
      tableEl = table;
      drawnRows = rows; // XSheet: ADR-0007 §4
      drawn = { first: range.first, last: range.last, firstCol: cols.first, lastCol: cols.last };
      // 図の層の原点は A1 の左上（行番号の列の幅・見出し行の高さだけ表の左上からずらす）
      // XSheet: 図の層は中を 100% で描き、層に倍率を掛ける（ADR-0007 §5）
      layer.place(originX, headY, g.totalWidth / zoom, g.totalHeight / zoom);
      layer.element.style.transform = zoom === 1 ? '' : `scale(${zoom})`;
      // 罫線の層も、表を作り直す時だけ同じ行・列の範囲から作り直す（スクロール・テーマ切替では触らない）
      if (sh.borders.length > 0) {
        borderLayer.render(
          buildBorderPaths(rows, sh.borders, g, { first: cols.first, last: cols.last }),
          originX,
          headY, // XSheet: ADR-0007 §7
        );
      } else {
        borderLayer.clear();
      }
    }
    // 図の出し入れは、表を描き直さない横スクロール・縦スクロールでも毎回判定する
    layer.update(
      unzoomRect(
        viewWindow(
          body.scrollLeft,
          body.scrollTop,
          body.clientWidth,
          body.clientHeight,
          originX,
          headY, // XSheet: ADR-0007 §7
        ),
        zoom, // XSheet: 図の層の中の座標（100%）で判定する（ADR-0007 §5）
      ),
    );
  };

  // XSheet: 拡大縮小（ADR-0007 §5）
  /** 幅に合わせる倍率（合わせ直さなくてよいときは null）。縦のスクロールバーが無いときは、出る分を先に引く */
  const fitTarget = (): number | null => {
    if (!('fit' in zoomSpec) || !sheet || zooming) return null;
    const bar = body.offsetWidth - body.clientWidth > 0 ? 0 : scrollbarWidth();
    const available = body.clientWidth - bar - headX(extent.rows || sheet.rowCount);
    return fitZoom(available, contentW, zoomSpec.fit, zoom);
  };
  /** 倍率を z にして描き直し、左上に見ていたセルを左上に残す */
  const applyZoom = (z: number): void => {
    const anchor = geo ? scrollAnchor(geo, body.scrollTop, body.scrollLeft) : null;
    zoom = z;
    root.style.setProperty('--xs-zoom', String(z));
    geo = null;
    zooming = true;
    try {
      // 新しい寸法でのスクロール位置が今の表の大きさに収まる（縮小など）なら、先にスクロールして 1 回で描く
      if (anchor && sheet) {
        refreshExtent(sheet);
        const g = geo as Geometry | null;
        const pos = g ? anchorScroll(g, anchor) : null;
        if (
          pos &&
          pos.top <= body.scrollHeight - body.clientHeight &&
          pos.left <= body.scrollWidth - body.clientWidth
        ) {
          body.scrollTop = pos.top;
          body.scrollLeft = pos.left;
          draw(true);
          return;
        }
      }
      draw(true);
      if (anchor && geo) {
        const pos = anchorScroll(geo, anchor);
        body.scrollTop = pos.top;
        body.scrollLeft = pos.left;
        draw();
      }
    } finally {
      zooming = false;
    }
  };

  /** スクロール・枠の大きさの変化を次のフレームの描画 1 回にまとめる */
  const schedule = (): void => {
    if (frame || closed) {
      return;
    }
    frame = requestAnimationFrame(() => {
      frame = 0;
      draw();
    });
  };

  // 枠の大きさが変わると見えている行・列の数が変わるため、描き足す（InterHub は window の resize。ここは枠そのものを見る）。
  // jsdom 等には ResizeObserver が無いためガードする。
  const observer = typeof ResizeObserver === 'function' ? new ResizeObserver(schedule) : null;

  const showSheet = async (index: number): Promise<void> => {
    if (closed) {
      return;
    }
    if (!book.sheets.some((s) => s.index === index)) {
      throw new RangeError('シートの番号が範囲外');
    }
    const seq = ++openSeq;
    const isStale = (): boolean => seq !== openSeq || closed;
    let opened: XlsxSheet;
    try {
      opened = await book.openSheet(index);
    } catch (e) {
      if (isStale()) {
        return;
      }
      sheet = null;
      geo = null;
      activeIndex = null;
      drawn = null;
      notices.clear();
      renderNotices();
      showMessage(bannerElement(xlsxErrorMessage(e, messages)));
      return;
    }
    // 図の読み取り。失敗しても表は出す（理由は通知にする）。図を出さない指定なら読まない
    let found: SheetDrawings = { drawings: [], notices: [] };
    let geomFailed = false;
    if (options.drawings) {
      try {
        found = await book.openDrawings(index);
      } catch {
        found = { drawings: [], notices: ['broken'] };
      }
      // 形のある図形があるときだけ、形の表（geom）を読む。読めなければ代替（種類名の破線枠）で描き、通知する
      if (needsGeom(found)) {
        try {
          await loadGeom();
        } catch {
          geomFailed = true;
        }
      }
    }
    if (isStale()) {
      return;
    }
    if (!fmtReady) {
      fmtReady = true;
      fmtStyle.textContent = buildFormatCss(opened.formats, fmtSeq);
    }
    activeIndex = index;
    drawn = null;
    hits = null; // XSheet: 一致の一覧はシートごと（ADR-0007 §7）
    drawnRows = []; // XSheet: ADR-0007 §4
    body.scrollTop = 0;
    body.scrollLeft = 0;
    notices.clear();
    // text-truncated・shape-simplified は読み取り側が 1 シートに 1 回だけ出す（Set なので重ならない）
    found.notices.forEach((n) => notices.add(n));
    if (geomFailed) {
      notices.add('geom-failed');
    }
    renderNotices();
    // 図の位置は geometry（行数・列数を Excel の最大にしたもの。範囲外の番号も既定の寸法で延びる）から出す。
    const bigGeo = createGeometry(opened.layout, MAX_GRID_ROWS, MAX_GRID_COLS);
    const drawings = found.drawings;
    const placed = layoutDrawings(bigGeo, drawings);
    if (opened.rowCount === 0 && placed.length === 0) {
      // 空シート（グラフシートを含む）。保持せず、データ無しの表示にする。
      // 大きさ 0 の図（置けない図）しか無いシートも、値が無ければ空シート扱いにする。
      sheet = null;
      geo = null;
      showMessage(emptyView(messages.noData));
      return;
    }
    sheet = opened;
    geo = null; // 範囲は draw が（スクロール枠の大きさに合わせて）決める
    // 値が無く図だけのシートも空シートにせず、図のある範囲までグリッドを出す。
    boxes = placed;
    drawingMin = drawingExtent(bigGeo, boxes);
    // XSheet: 幅に合わせる中身の幅（データのある列と図の右端。ADR-0007 §5）
    contentW = contentWidth(bigGeo.colLeft(opened.colCount + 1), boxes);
    // 前のシートの blob URL はここ（新しいシートへの切替）ですべて revoke される
    layer.setSheet(
      options.drawings
        ? {
            drawings,
            boxes,
            readImage: (path) => book.readImage(index, path),
            // グラフ（InterHub ISSUE-103）：図が見える範囲に入った時に読む。描けなければ破線枠のまま
            readChart: (path) => book.readChart(index, path),
            onOverBudget: () => {
              // 画像の合計の上限超過（通知）。別のシートへ切り替わった後の結果は捨てる
              if (!isStale()) {
                notices.add('over-budget');
                renderNotices();
              }
            },
          }
        : null,
    );
    draw(true);
    // 取り付け直後は枠の大きさが確定していないことがあるため、次のフレームで実寸で描き直す
    cancelAnimationFrame(initFrame);
    initFrame = requestAnimationFrame(() => {
      initFrame = 0;
      draw(true);
    });
  };

  const destroy = (): void => {
    if (closed) {
      return;
    }
    closed = true;
    openSeq++;
    // blob URL を先にすべて revoke する（読み込み中の画像の結果は捨てられる）
    layer.dispose();
    cancelAnimationFrame(frame);
    cancelAnimationFrame(initFrame);
    frame = 0;
    initFrame = 0;
    body.removeEventListener('scroll', schedule);
    observer?.disconnect();
    sheet = null;
    geo = null;
    drawn = null;
    activeIndex = null;
    borderLayer.clear();
    boxes = [];
    tableEl = null;
    notices.clear();
    hits = null; // XSheet: ADR-0007 §2
    drawnRows = []; // XSheet: ADR-0007 §4
    nav.dispose();
    wrapper.replaceChildren();
    body.replaceChildren();
    root.replaceChildren();
    root.removeAttribute('data-xs-fmt');
    root.remove();
  };

  // XSheet: 今のヒットへの移動・位置の判定（ADR-0007 §3・§4）。表示の今の状態を引いて計算する
  const nav = createNavigation({
    body,
    root,
    sheet: () => sheet,
    geometry: () => geo,
    origin: () => ({ x: headX(extent.rows), y: headY }),
    cols: () => extent.cols,
    drawn: () => ({ rows: drawnRows, firstCol: drawn?.firstCol ?? 1 }),
    hits: () => hits,
    table: () => tableEl,
    draw: () => draw(),
  });

  root.append(noticeBar, body, fmtStyle);
  container.append(root);
  body.addEventListener('scroll', schedule, { passive: true });
  observer?.observe(body);

  return {
    root,
    get activeIndex() {
      return activeIndex;
    },
    showSheet,
    // XSheet: ADR-0007 §2
    setHighlights(matches) {
      hits = createGridHits(matches);
      if (!closed && sheet) {
        draw(true);
      }
    },
    // XSheet: ADR-0007 §3・§4
    focusHighlight: (index, opts) => {
      if (!closed) nav.focusHighlight(index, opts);
    },
    scrollToCell: (r, c, opts) => {
      if (!closed) nav.scrollToCell(r, c, opts);
    },
    cellAt: (x, y) => (closed ? null : nav.cellAt(x, y)),
    highlightAt: (x, y) => (closed ? -1 : nav.highlightAt(x, y)),
    // XSheet: ADR-0007 §5
    setZoom(next) {
      const spec = resolveZoom(next);
      if (closed) return;
      zoomSpec = spec;
      if ('fixed' in spec) {
        if (spec.fixed !== zoom) applyZoom(spec.fixed);
      } else {
        draw(); // 幅に合わせる（draw が倍率を出す）
      }
    },
    get zoom() {
      return zoom;
    },
    destroy,
  };
}
