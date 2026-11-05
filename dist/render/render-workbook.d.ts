/**
 * 公開の入口 `renderWorkbook`（M2-4）：xlsx のバイト列を開き、シートタブ付きで `container` の中に表示する。
 * InterHub `spreadsheet.ts` の `renderXlsxView` のうち、ブックを開く・タブ・最初のシートの選択・エラー表示の部分。
 * - 表示は 1 冊につき 1 つのシートの表示（`createSheetView`）を使い回し、タブでシートを替える
 *   （書式の `<style>` はブックで 1 回だけ作り、タブを替えても作り直さない）。
 * - 開けない（暗号化・壊れている・大きすぎる・対応しない形式）ときは例外にせず、種類ごとの文言を `div.xs-banner` に出す。
 * - `destroy()` でタブ列・シートの表示（blob URL・リスナ）を放し、ブックを閉じ、根を外す。
 * シート名はユーザーのファイル内容（信頼できない入力）なので、`textContent` と `title` にだけ入れる。
 */
import { type SheetInfo, type XlsxLimits } from '../core/index';
import type { RenderOptions, SheetView } from './render-sheet';
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
/**
 * xlsx を開き、シートタブ付きで `container` の中に表示する。
 * 最初に出すシートは、最初の表示（非表示でない）シート。全部非表示なら先頭。
 * タブ列は既定で表の上、`tabPosition: 'bottom'` なら表の下。
 * `container` は高さが決まっていること（タブ列の下のスクロール枠 `.xs-body` が仮想化の窓になる）。
 *
 * @param container 表示を置く要素（子として `div.xs-workbook` を足す。`container` 自身には何も付けない）
 * @param data xlsx のバイト列
 */
export declare function renderWorkbook(container: HTMLElement, data: ArrayBuffer | Uint8Array, options?: WorkbookOptions): Promise<WorkbookView>;
