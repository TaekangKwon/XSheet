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
import { type XlsxBook } from '../core/index';
import type { TextMatch } from '../core/index';
import { type Navigation } from './navigate';
import { type Zoom } from './zoom';
import { type XSheetMessages } from './messages';
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
export declare function xlsxErrorMessage(e: unknown, messages: Readonly<XSheetMessages>): string;
/**
 * 1 シートの表示を作り、根を `container` に取り付ける（シートは {@link SheetViewInstance.showSheet} で開く）。
 * 保持するのは表示中の 1 シートだけ。シートを替えると前のシートを放す（ADR-0025 決定 4）。
 *
 * @param container 表示を置く要素（高さが決まっていること）
 * @param book 開いたブック（閉じるのは呼び出し側）
 */
export declare function createSheetView(container: HTMLElement, book: XlsxBook, options: SheetViewOptions): SheetViewInstance;
