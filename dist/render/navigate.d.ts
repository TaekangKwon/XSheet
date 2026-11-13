/**
 * 今のヒットへの移動と、画面の位置の判定（M7-6・Moobius の依頼 R3・R4）。
 * - スクロール先は寸法（`Geometry`）からスクロールする前にすべて計算する（描いた後に測って直さない。ADR-0007 §3）。
 *   縦は「セルの下端から 1 行の高さを引いた位置」を表示枠の上から約 1/3 に置く（文字は下揃えなので、高い行・縦の結合でも文字が見える）。
 *   横は、対象が見えていないときだけ、見える最小の移動（右にあれば右端に寄せる・左にあれば左端に寄せる）にする。
 * - 位置の判定は寸法の `rowAt`・`colAt` で出し、DOM を見ない。結合は左上、はみ出した文字の上ははみ出し元のセルを、
 *   描いた行（`GridRow`）から引く。
 * 見出し（列見出しの高さ・行番号の列の幅）は引数でもらう（拡大縮小・見出しを隠すで変わるため。ADR-0007 §5・§7）。
 */
import type { Geometry, XlsxSheet } from '../core/index';
import { type GridRow } from './grid';
import { type GridHits } from './highlight';
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
export declare function verticalTarget(geo: Pick<Geometry, 'rowTop' | 'rowHeight' | 'totalHeight'>, view: ViewBox, r: number, rowSpan?: number): number;
/**
 * 横のスクロール先。範囲が表示枠のセルの部分に収まっていれば今のまま。
 * 右にはみ出せば右端を枠の右端に、左にはみ出せば（または枠より広ければ）左端を枠の左端に合わせる。
 * 左へ動かすときは、範囲が枠に収まる限りセルの左端（`cellLeft`）から見せる（セルの文字の頭が行番号の陰に入らない）。
 */
export declare function horizontalTarget(geo: Pick<Geometry, 'totalWidth'>, view: ViewBox, range: XRange): number;
/** セル（列 c から colSpan 列）の横の範囲 */
export declare function cellXRange(geo: Pick<Geometry, 'colLeft'>, c: number, colSpan?: number): XRange;
/**
 * 画面の位置（スクロール枠の表示の左上からの px）にある行・列。見出しの上・表示枠の外・表の外は null。
 * 結合とはみ出しは見ない（{@link resolveDrawnCell} で直す）。
 */
export declare function pointToCell(geo: Pick<Geometry, 'rowAt' | 'colAt' | 'totalWidth' | 'totalHeight'>, view: ViewBox, x: number, y: number): {
    r: number;
    c: number;
} | null;
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
export declare function resolveDrawnCell(rows: readonly GridRow[], r: number, c: number, cols: number, firstCol?: number): {
    r: number;
    c: number;
};
/**
 * 文字の幅を測る入れ物（画面の外・見えない）。セルと同じ書式のクラスを付けて測る（書式の規則は表示の根の `<style>`）。
 * 表示の根の中に置き、使い回す。
 */
export declare function createTextMeasure(root: HTMLElement): {
    width(text: string, formatClass: string): number;
    dispose(): void;
};
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
    origin(): {
        x: number;
        y: number;
    };
    /** 表の列数 */
    cols(): number;
    /** 描いた行と、描いた最初の列（まだ描いていなければ rows は空） */
    drawn(): {
        rows: readonly GridRow[];
        firstCol: number;
    };
    /** ハイライト（無ければ null） */
    hits(): GridHits | null;
    /** 描いた表（無ければ null） */
    table(): HTMLTableElement | null;
    /** 今のスクロール位置で描く（描いた範囲に収まっていれば何もしない） */
    draw(): void;
}
/** {@link createNavigation} の戻り値（`SheetView` の §3・§4 の操作） */
export interface Navigation {
    focusHighlight(index: number, options?: {
        scroll?: boolean;
        smooth?: boolean;
    }): void;
    scrollToCell(r: number, c: number, options?: {
        smooth?: boolean;
    }): void;
    cellAt(clientX: number, clientY: number): {
        r: number;
        c: number;
    } | null;
    highlightAt(clientX: number, clientY: number): number;
    dispose(): void;
}
/**
 * 移動・判定（ADR-0007 §3・§4）を作る。状態は持たず、呼ぶたびに host から今の表示を引く（測る入れ物だけを持つ）。
 */
export declare function createNavigation(host: NavigationHost): Navigation;
