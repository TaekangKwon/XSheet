// 正: docs/02_architecture.md §3
/**
 * 公開の入口 `renderWorkbook`（M2-4）：xlsx のバイト列を開き、シートタブ付きで `container` の中に表示する。
 * InterHub `spreadsheet.ts` の `renderXlsxView` のうち、ブックを開く・タブ・最初のシートの選択・エラー表示の部分。
 * - 表示は 1 冊につき 1 つのシートの表示（`createSheetView`）を使い回し、タブでシートを替える
 *   （書式の `<style>` はブックで 1 回だけ作り、タブを替えても作り直さない）。
 * - 開けない（暗号化・壊れている・大きすぎる・対応しない形式）ときは例外にせず、種類ごとの文言を `div.xs-banner` に出す。
 * - `destroy()` でタブ列・シートの表示（blob URL・リスナ）を放し、ブックを閉じ、根を外す。
 * シート名はユーザーのファイル内容（信頼できない入力）なので、`textContent` と `title` にだけ入れる。
 */
import { openXlsx, type SheetInfo, type XlsxBook, type XlsxLimits } from '../core/index';
import { resolveMessages, type XSheetMessages } from './messages';
import type { RenderOptions, SheetView } from './render-sheet';
import { buildSheetTabStrip, makeSheetTab } from './sheet-tabs';
import { createSheetView, xlsxErrorMessage } from './sheet-view';
// XSheet: ADR-0007 §5
import { resolveZoom } from './zoom';

/** `renderWorkbook` の指定 */
export interface WorkbookOptions extends RenderOptions {
  /** 読み取りの上限を替える（`openXlsx` の `limits`。既定は `DEFAULT_LIMITS`） */
  limits?: Partial<XlsxLimits>;
  /**
   * シートタブの列を置く場所（既定 `'top'`＝表の上。`'bottom'` は Excel と同じく表の下）。
   * `'bottom'` ではタブ列を DOM でもシートの表示の後ろに置き、根に `xs-tabs-bottom` を付ける。
   * 知らない値は `'top'` と同じ。
   * XSheet: タブの位置を選べるように足した（Issue #17）
   */
  tabPosition?: 'top' | 'bottom';
}

/** タブ付きのブックの表示 */
export interface WorkbookView extends SheetView {
  /**
   * ブックのシートの一覧（タブの並び。開けなかったブックでは空）。
   * XSheet: `view` の `Viewer.sheets`・名前でのシートの切り替えのために足した（M6-2）
   */
  readonly sheets: readonly SheetInfo[];
  /**
   * シートを替える（タブを押したのと同じ）。表示中のシートなら何もしない。
   * @param index シートの番号（`SheetInfo.index`。1 始まり・非表示のシートも数える）
   * @throws RangeError ブックに無い番号（開けなかったブック・シートの無いブックでは常に）
   */
  showSheet(index: number): Promise<void>;
}

/** メッセージだけの要素（`div.xs-banner`・`div.xs-empty`） */
function messageElement(className: string, text: string): HTMLElement {
  const el = document.createElement('div');
  el.className = className;
  el.textContent = text;
  return el;
}

/** シートを持たない表示（開けない・シートが 0 枚）。destroy は根を外すだけ */
function messageView(root: HTMLElement): WorkbookView {
  return {
    sheets: [],
    showSheet: () => Promise.reject(new RangeError('シートの番号が範囲外')),
    setHighlights: () => undefined, // XSheet: 表が無いので何もしない（ADR-0007 §2）
    // XSheet: 一致もセルも無い（番号の確かめだけは表のある表示と同じ。ADR-0007 §3・§4）
    focusHighlight: (index) => {
      if (index !== -1) throw new RangeError('ヒットの番号が範囲外');
    },
    scrollToCell: (r, c) => {
      if (!Number.isInteger(r) || !Number.isInteger(c) || r < 1 || c < 1) {
        throw new RangeError('セルの番号が範囲外');
      }
    },
    cellAt: () => null,
    highlightAt: () => -1,
    // XSheet: 表が無いので倍率は 1 のまま（指定の確かめだけ。ADR-0007 §5）
    setZoom: (zoom) => {
      resolveZoom(zoom);
    },
    zoom: 1,
    destroy: () => {
      root.replaceChildren();
      root.remove();
    },
  };
}

/**
 * xlsx を開き、シートタブ付きで `container` の中に表示する。
 * 最初に出すシートは、最初の表示（非表示でない）シート。全部非表示なら先頭。
 * タブ列は既定で表の上、`tabPosition: 'bottom'` なら表の下。
 * `container` は高さが決まっていること（タブ列の下のスクロール枠 `.xs-body` が仮想化の窓になる）。
 *
 * @param container 表示を置く要素（子として `div.xs-workbook` を足す。`container` 自身には何も付けない）
 * @param data xlsx のバイト列
 */
export async function renderWorkbook(
  container: HTMLElement,
  data: ArrayBuffer | Uint8Array,
  options: WorkbookOptions = {},
): Promise<WorkbookView> {
  const messages: Readonly<XSheetMessages> = resolveMessages(options.messages);
  if (options.zoom !== undefined) {
    resolveZoom(options.zoom); // XSheet: 正しくない倍率は、DOM を足す前・ブックを開く前に RangeError（ADR-0007 §5）
  }
  const root = document.createElement('div');
  root.className = 'xs-workbook';
  container.append(root);

  let book: XlsxBook;
  try {
    book = await openXlsx(data, { limits: options.limits });
  } catch (e) {
    root.append(messageElement('xs-banner', xlsxErrorMessage(e, messages)));
    return messageView(root);
  }
  // 開いている間に destroy は呼べない（戻り値の約束がまだ解決していない）ため、ここでの離脱の判定は要らない

  const sheets = book.sheets;
  if (sheets.length === 0) {
    book.close();
    root.append(messageElement('xs-empty', messages.noSheets));
    return messageView(root);
  }

  let closed = false;
  const buttons = sheets.map((info) => makeSheetTab(info.name, info.hidden, messages));
  const strip = buildSheetTabStrip(buttons, messages);
  // XSheet: タブ列を下に置くときは、DOM の順も見た目と同じ（シートの表示 → タブ列）にする（Issue #17）
  const tabsBottom = options.tabPosition === 'bottom';
  root.append(strip.element);
  // シートの表示は 1 冊に 1 つ（タブ列の後ろに `div.xs-sheet` を足す）
  const view = createSheetView(root, book, {
    messages,
    drawings: options.drawings !== false,
    nonce: options.nonce,
    zoom: options.zoom,
  });
  if (tabsBottom) {
    root.classList.add('xs-tabs-bottom');
    root.append(strip.element); // 足し直すと末尾へ動く
  }

  /** 指定のシートを開いて表示する（タブ操作・初期表示）。 */
  const showSheet = async (index: number): Promise<void> => {
    if (closed) {
      return;
    }
    const at = sheets.findIndex((s) => s.index === index);
    if (at < 0) {
      throw new RangeError('シートの番号が範囲外');
    }
    if (view.activeIndex === index) {
      return; // 表示中（開けた）のシート。開けなかったシートは activeIndex が null なので開き直す
    }
    buttons.forEach((b, bi) => b.classList.toggle('is-active', bi === at));
    // 選択中タブが見える位置までタブ列をスクロールする
    strip.reveal(buttons[at]);
    await view.showSheet(index);
  };

  const onClicks = buttons.map((btn, bi) => {
    const onClick = (): void => {
      void showSheet(sheets[bi]!.index);
    };
    btn.addEventListener('click', onClick);
    return onClick;
  });

  const destroy = (): void => {
    if (closed) {
      return;
    }
    closed = true;
    buttons.forEach((btn, bi) => btn.removeEventListener('click', onClicks[bi]!));
    strip.dispose();
    // シートの表示（blob URL の revoke・リスナ）を放してから、ブックを閉じる
    view.destroy();
    book.close();
    root.replaceChildren();
    root.remove();
  };

  try {
    await showSheet((sheets.find((s) => !s.hidden) ?? sheets[0]!).index);
  } catch (e) {
    // 想定外の例外でも、リスナ・blob URL・ブックを残さない
    destroy();
    throw e;
  }
  // XSheet: setHighlights は中のシートの表示に渡す（ADR-0007 §2）
  return {
    sheets,
    showSheet,
    setHighlights: (matches) => view.setHighlights(matches),
    focusHighlight: (index, opts) => view.focusHighlight(index, opts),
    scrollToCell: (r, c, opts) => view.scrollToCell(r, c, opts),
    cellAt: (x, y) => view.cellAt(x, y),
    highlightAt: (x, y) => view.highlightAt(x, y),
    setZoom: (zoom) => view.setZoom(zoom),
    get zoom() {
      return view.zoom;
    },
    destroy,
  };
}
