/**
 * 確認用の入口 `view2`（公開しない。`src/facade/index.ts` からは出さず、配布物にも入らない）。
 * ADR-0007 の未決の 2 つを、今の `view` を変えずに試すために置く。決めたら `view` に入れるか、このファイルを消す。
 * 1. `Viewer` に 0.5.0 の操作を足す（§7）：ハイライト・移動・位置の判定・倍率・今のシート・シートの切り替えの通知に加え、
 *    表示中のシートを探してハイライトまで付ける `find`（ブックを外で開き直さずに済む）。
 * 2. 組み込みの日付の書式の既定を日本語版 Excel の表にする（§6）：`builtinDates` の既定を `'japanese-era'` にする。
 *    `'default'` を渡せば今の `view` と同じ表になる（比べられるように）。
 * 作りは `view` と同じ（表示の根 `div.xs-viewer` の中に `renderWorkbook` を置く）。
 */
import { type BuiltinDates, type FindTextOptions, type FindTextResult, type SheetInfo, type TextMatch } from '../core/index';
import { type CellPosition, type FocusHighlightOptions, type PrintOptions, type PrintResult, type ScrollToCellOptions, type Selection, type SelectionInput, type SelectionMode, type SelectionOrigin, type Zoom } from '../render/index';
import { type XSheetSource } from './source';
import type { ViewOptions } from './view';
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
    /** クリックで選ぶ物（`renderWorkbook` の `selection` と同じ。既定は無効。ADR-0011） */
    selection?: SelectionMode;
    /** 選択が変わったときに呼ぶ（`renderWorkbook` の `onSelectionChange` と同じ。ADR-0011） */
    onSelectionChange?: (selection: Selection | null, origin: SelectionOrigin) => void;
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
    /** スクロール枠（`renderWorkbook` の `scrollElement` と同じ。開けなかったファイル・destroy の後は null。ADR-0010 §1） */
    readonly scrollElement: HTMLElement | null;
    /** 印刷のために表示中のシートをページに分けて描く（`renderWorkbook` の `renderForPrint` と同じ。ADR-0010 §2） */
    renderForPrint(target: HTMLElement, options?: PrintOptions): Promise<PrintResult>;
    /** 今の選択（`renderWorkbook` の `getSelection` と同じ。ADR-0011） */
    getSelection(): Selection | null;
    /** 選択を替える（`renderWorkbook` の `setSelection` と同じ。ADR-0011） */
    setSelection(selection: SelectionInput | null): void;
    /** 選択をコピーする文字（`renderWorkbook` の `selectionText` と同じ。ADR-0011） */
    selectionText(): string | null;
    /** 表示を外し、ブック・リスナ・blob URL を放す。何度呼んでもよい */
    destroy(): void;
}
/**
 * xlsx をタブ付きで表示する（確認用。`view` と同じ動きに 0.5.0 の操作と日本語版 Excel の組み込みの表を足す）。
 * @param target 表示を置く要素か CSS セレクタ
 * @param source File・Blob・ArrayBuffer・型付き配列・Response、またはその Promise
 * @throws TypeError 要素が見つからない・受けない形の source・失敗した Response
 */
export declare function view2(target: HTMLElement | string, source: XSheetSource, options?: View2Options): Promise<Viewer2>;
