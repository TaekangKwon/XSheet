// 正: docs/adr/0007-moobius-viewer-api.md §6・§7（未決の 2 つを試す確認用）
/**
 * 確認用の入口 `view2`（公開しない。`src/facade/index.ts` からは出さず、配布物にも入らない）。
 * ADR-0007 の未決の 2 つを、今の `view` を変えずに試すために置く。決めたら `view` に入れるか、このファイルを消す。
 * 1. `Viewer` に 0.5.0 の操作を足す（§7）：ハイライト・移動・位置の判定・倍率・今のシート・シートの切り替えの通知に加え、
 *    表示中のシートを探してハイライトまで付ける `find`（ブックを外で開き直さずに済む）。
 * 2. 組み込みの日付の書式の既定を日本語版 Excel の表にする（§6）：`builtinDates` の既定を `'japanese-era'` にする。
 *    `'default'` を渡せば今の `view` と同じ表になる（比べられるように）。
 * 作りは `view` と同じ（表示の根 `div.xs-viewer` の中に `renderWorkbook` を置く）。
 */
import {
  findText,
  openXlsx,
  type BuiltinDates,
  type FindTextOptions,
  type FindTextResult,
  type SheetInfo,
  type TextMatch,
  type XlsxBook,
  type XlsxSheet,
} from '../core/index';
import {
  renderWorkbook,
  type CellPosition,
  type FocusHighlightOptions,
  type ScrollToCellOptions,
  type Zoom,
} from '../render/index';
import { toSourceBytes, type XSheetSource } from './source';
import { applyHeight, applyTheme, ensureStyle } from './style';
import type { ViewOptions } from './view';
import { pickMessages, resolveTarget } from './view-shared';

/** `view2` の指定（`view` の指定に 0.5.0 の表示の指定を足す） */
export interface View2Options extends ViewOptions {
  /** 組み込みの日付の書式 27〜36・50〜58 の見せ方（既定 `'japanese-era'`＝日本語版 Excel の表。`view` の既定は `'default'`） */
  builtinDates?: BuiltinDates;
  /** 最初の倍率（`renderWorkbook` の `zoom` と同じ。既定 1） */
  zoom?: Zoom;
  /** 列記号・行番号の見出しを出すか（既定 true） */
  headers?: boolean;
  /** シートが替わり、表示し終えた後に呼ぶ（`renderWorkbook` の `onSheetChange` と同じ。番号は `SheetInfo.index`） */
  onSheetChange?: (index: number) => void;
}

/** `Viewer2.find` の指定 */
export interface View2FindOptions extends FindTextOptions {
  /** 一致があれば最初のヒットを今のヒットにしてスクロールする（既定 true） */
  focus?: boolean;
}

/** `view2` の戻り値（`Viewer` に 0.5.0 の操作を足したもの） */
export interface Viewer2 {
  /** シートの一覧（タブの並び。非表示のシートも入る。開けなかったファイルでは空） */
  readonly sheets: readonly SheetInfo[];
  /** 表示中のシートの番号（`SheetInfo.index`・1 始まり）。開けなかったファイルでは null */
  readonly activeSheet: number | null;
  /** 今の倍率（fit のときは計算した値） */
  readonly zoom: number;
  /**
   * シートを替える（表示を作り直さない。倍率は保ち、ハイライトは消す）。
   * @param sheet シートの番号（`SheetInfo.index`・1 始まり）か名前
   * @throws RangeError ブックに無い番号・名前
   */
  showSheet(sheet: number | string): Promise<void>;
  /**
   * 表示中のシートを探し、一致をハイライトする（前の `find` は中断する）。空の文字ならハイライトを消す。
   * ハイライトはシートを替えると消える（`onSheetChange` で探し直す）。
   * @throws AbortError 次の `find`・シートの切り替え・`destroy`・`signal` で中断した
   */
  find(query: string, options?: View2FindOptions): Promise<FindTextResult>;
  /** 一致を光らせる（`renderWorkbook` の `setHighlights` と同じ） */
  setHighlights(matches: readonly TextMatch[] | null): void;
  /** 今のヒットを替えてスクロールする（`focusHighlight` と同じ） */
  focusHighlight(index: number, options?: FocusHighlightOptions): void;
  /** セルが見える所までスクロールする（行・列は 1 始まり） */
  scrollToCell(r: number, c: number, options?: ScrollToCellOptions): void;
  /** 画面の座標にあるセル（無ければ null） */
  cellAt(clientX: number, clientY: number): CellPosition | null;
  /** 画面の座標にあるハイライトの番号（無ければ -1） */
  highlightAt(clientX: number, clientY: number): number;
  /** 倍率を替える（0.5〜2.0、または `{ fit: 'width' }`） */
  setZoom(zoom: Zoom): void;
  /** 表示を外し、ブック・リスナ・blob URL を放す。何度呼んでもよい */
  destroy(): void;
}

/** 中断の例外（`findText` と同じ形） */
function abortError(): Error {
  if (typeof DOMException !== 'undefined') return new DOMException('検索を中断した', 'AbortError');
  const e = new Error('検索を中断した');
  e.name = 'AbortError';
  return e;
}

/** 同じ container の今の表示（`view` とは別に持つ） */
const current = new WeakMap<HTMLElement, { token: object; viewer: Viewer2 | null }>();

/** 捨てた読み込みの戻り値（表示は無い） */
function discarded(): Viewer2 {
  const replaced = (): Promise<never> =>
    Promise.reject(new RangeError('XSheet: the viewer was replaced or destroyed'));
  return {
    sheets: [],
    activeSheet: null,
    zoom: 1,
    showSheet: replaced,
    find: replaced,
    setHighlights: () => undefined,
    focusHighlight: () => undefined,
    scrollToCell: () => undefined,
    cellAt: () => null,
    highlightAt: () => -1,
    setZoom: () => undefined,
    destroy: () => undefined,
  };
}

/**
 * xlsx をタブ付きで表示する（確認用。`view` と同じ動きに 0.5.0 の操作と日本語版 Excel の組み込みの表を足す）。
 * @param target 表示を置く要素か CSS セレクタ
 * @param source File・Blob・ArrayBuffer・型付き配列・Response、またはその Promise
 * @throws TypeError 要素が見つからない・受けない形の source・失敗した Response
 */
export async function view2(
  target: HTMLElement | string,
  source: XSheetSource,
  options: View2Options = {},
): Promise<Viewer2> {
  const container = resolveTarget(target);
  const token = {};
  current.get(container)?.viewer?.destroy();
  current.set(container, { token, viewer: null });
  const isCurrent = (): boolean => current.get(container)?.token === token;
  let bytes;
  try {
    bytes = await toSourceBytes(source);
  } catch (e) {
    if (isCurrent()) current.delete(container);
    throw e;
  }
  if (!isCurrent()) return discarded();
  if (options.css !== 'none') ensureStyle(container, options.nonce);
  const root = container.ownerDocument.createElement('div');
  root.className = 'xs-viewer';
  applyTheme(root, options.theme);
  const disposeHeight = applyHeight(container, root, options.height);
  container.append(root);

  // §6 を試す：既定を日本語版 Excel の表にする（検索用に開くブックも同じ表で読む）
  const builtinDates: BuiltinDates = options.builtinDates ?? 'japanese-era';
  /** 今の検索（次の検索・シートの切り替え・destroy で中断する） */
  let finding: AbortController | null = null;
  const stopFinding = (): void => {
    finding?.abort();
    finding = null;
  };
  let workbook;
  try {
    workbook = await renderWorkbook(root, bytes, {
      messages: pickMessages(container.ownerDocument, options),
      drawings: options.drawings,
      tabPosition: options.tabPosition,
      nonce: options.nonce,
      limits: options.limits,
      builtinDates,
      zoom: options.zoom,
      headers: options.headers,
      onSheetChange: (index) => {
        stopFinding();
        options.onSheetChange?.(index);
      },
    });
  } catch (e) {
    disposeHeight();
    root.remove();
    if (isCurrent()) current.delete(container);
    throw e;
  }
  const wb = workbook;
  if (!isCurrent()) {
    disposeHeight();
    wb.destroy();
    root.remove();
    return discarded();
  }

  // 検索用のブック（表示の中のブックは外から見えないので、初めて探すときに同じ bytes・同じ指定で開く）
  const data = bytes;
  let searchBook: Promise<XlsxBook> | null = null;
  const sheets = new Map<number, Promise<XlsxSheet>>();
  const openForSearch = (index: number): Promise<XlsxSheet> => {
    searchBook ??= openXlsx(data, { limits: options.limits, builtinDates });
    let sheet = sheets.get(index);
    if (sheet === undefined) {
      const opened = searchBook.then((book) => book.openSheet(index));
      opened.catch(() => sheets.delete(index)); // 開けなかったら次は開き直す
      sheets.set(index, opened);
      sheet = opened;
    }
    return sheet;
  };

  let destroyed = false;
  const destroy = (): void => {
    if (destroyed) return;
    destroyed = true;
    stopFinding();
    void searchBook?.then((book) => book.close()).catch(() => undefined);
    disposeHeight();
    wb.destroy();
    root.remove();
    if (current.get(container)?.viewer === viewer) current.delete(container);
  };

  const viewer: Viewer2 = {
    sheets: wb.sheets,
    get activeSheet() {
      return wb.activeSheet;
    },
    get zoom() {
      return wb.zoom;
    },
    showSheet(sheet: number | string): Promise<void> {
      const info = wb.sheets.find((s) => (typeof sheet === 'number' ? s.index : s.name) === sheet);
      if (info === undefined) return Promise.reject(new RangeError('XSheet: no such sheet'));
      if (info.index !== wb.activeSheet) stopFinding();
      return wb.showSheet(info.index);
    },
    async find(query: string, findOptions: View2FindOptions = {}): Promise<FindTextResult> {
      stopFinding();
      const index = wb.activeSheet;
      if (destroyed || index === null || query === '') {
        wb.setHighlights(null);
        return { matches: [], truncated: false };
      }
      const ac = new AbortController();
      finding = ac;
      const outer = findOptions.signal;
      const signal = {
        get aborted() {
          return ac.signal.aborted || outer?.aborted === true;
        },
      };
      if (signal.aborted) throw abortError();
      const sheet = await openForSearch(index);
      if (signal.aborted) throw abortError();
      const result = await findText(sheet, query, { ...findOptions, signal });
      // 探している間にシートが替わった・次の検索が始まった（findText が気づく前に終わった場合）
      if (signal.aborted || finding !== ac || wb.activeSheet !== index) throw abortError();
      finding = null;
      wb.setHighlights(result.matches);
      if (result.matches.length > 0 && findOptions.focus !== false) wb.focusHighlight(0);
      return result;
    },
    setHighlights: (matches) => wb.setHighlights(matches),
    focusHighlight: (index, focusOptions) => wb.focusHighlight(index, focusOptions),
    scrollToCell: (r, c, scrollOptions) => wb.scrollToCell(r, c, scrollOptions),
    cellAt: (x, y) => wb.cellAt(x, y),
    highlightAt: (x, y) => wb.highlightAt(x, y),
    setZoom: (zoom) => wb.setZoom(zoom),
    destroy,
  };
  current.set(container, { token, viewer });
  return viewer;
}
