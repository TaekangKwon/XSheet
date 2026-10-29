// 正: docs/02_architecture.md §3
/**
 * 公開の入口 `renderSheet`（M2-3）：開いたブックの 1 シートを `container` の中に表示する。
 * 表示は呼び出しごとのインスタンスで、`destroy()` でリスナ・blob URL・シートを放す（ブックは閉じない）。
 * 中身は `sheet-view.ts`（InterHub `spreadsheet.ts` の .xlsx の部分）。タブ付きのブック全体は `renderWorkbook`（M2-4）。
 */
import type { XlsxBook } from '../core/index';
import { resolveMessages, type XSheetMessages } from './messages';
import { createSheetView } from './sheet-view';

/** 表示の指定 */
export interface RenderOptions {
  /** 文言（一部のキーだけでよい。既定は日本語の `MESSAGES_JA`） */
  messages?: Partial<XSheetMessages>;
  /** 図の層（画像・図形・線）を出すか（既定 true）。false なら図を読まず、図の通知も出さない */
  drawings?: boolean;
  /** 書式の `<style>` に付ける nonce（CSP の `style-src 'nonce-…'` 用） */
  nonce?: string;
}

/** 表示 1 つ */
export interface SheetView {
  /** リスナ・ResizeObserver・blob URL・シートを放し、表示を `container` から外す。何度呼んでもよい。ブックは閉じない */
  destroy(): void;
}

/**
 * ブックの 1 シートを `container` の中に表示する。
 * `container` は高さが決まっていること（中のスクロール枠 `.xs-body` が仮想化の窓になる）。
 * シートを開けない（壊れている等）ときは、例外にせず枠の中にメッセージを出す。
 *
 * @param container 表示を置く要素（子として `div.xs-sheet` を足す。`container` 自身には何も付けない）
 * @param book `openXlsx` で開いたブック（閉じるのは呼び出し側）
 * @param sheetIndex シートの番号（`SheetInfo.index`。1 始まり・非表示のシートも数える）
 * @throws RangeError ブックに無い番号（DOM には何も足さない）
 */
export async function renderSheet(
  container: HTMLElement,
  book: XlsxBook,
  sheetIndex: number,
  options: RenderOptions = {},
): Promise<SheetView> {
  if (!book.sheets.some((s) => s.index === sheetIndex)) {
    throw new RangeError('シートの番号が範囲外');
  }
  const view = createSheetView(container, book, {
    messages: resolveMessages(options.messages),
    drawings: options.drawings !== false,
    nonce: options.nonce,
  });
  try {
    await view.showSheet(sheetIndex);
  } catch (e) {
    view.destroy();
    throw e;
  }
  return { destroy: () => view.destroy() };
}
