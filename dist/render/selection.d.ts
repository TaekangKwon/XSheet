/**
 * セル・行・列の選択（XSheet の追加。M7-24・Issue #55）。
 * - 選択は 1 つの範囲（行の全体・列の全体・セル）と、今のセル（キーボードで動かす起点）。
 * - 帯はスクロール枠の中の `div.xs-sel`（重ねる層）1 つで描く。表は見えている近くだけを描き直すので、`td` に印を付けると
 *   描き直しのたびに付け直しが要り、見えていない部分は塗れないため。位置は寸法（`Geometry`）から出す。
 *   重なりは罫線の層の後・図の層の前（セル・はみ出しの文字・ハイライト・罫線の上、図・sticky の見出しの下）。
 * - 選んだ行の行番号・列の列記号は、表を描くたびに `is-selected` を付け直す（描いた行・列だけ）。
 * - クリック：セル → 指定（`'row'` は行の全体、`'cell'` はセル）、行番号 → 行の全体、列記号 → 列の全体。
 *   結合したセルは結合の範囲。動いた距離が小さいクリックだけを選択にし、文字をドラッグで選んでいるときは何もしない。
 * - キーボード（M7-25）：スクロール枠に `tabindex="0"` を付け、クリックで選んだらフォーカスを移す。行き先は selection-keys.ts。
 *   選択を動かしたキーだけ既定の動きを止め、見えない所へ動いたら最小の量だけスクロールする。
 *   選択の番地（`C5`・`5:5`・`C:C`）を見えない `aria-live` の要素に入れて、支援技術に伝える。
 * - コピー（M7-26）：フォーカスのあるスクロール枠の `copy` イベントで、選択をタブ区切りの文字にして書く（selection-copy.ts）。
 *   文字を選んでいるときはブラウザの既定のコピーに任せる。上限を超えたらコピーせず、通知の帯に出す。
 * DOM の要素は作るだけで、ユーザーのファイル内容は入れない。
 */
import type { CellRange, colName, createGeometry, Geometry, XlsxSheet } from '../core/index';
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
    readonly active: {
        readonly r: number;
        readonly c: number;
    };
}
/** `setSelection` に渡す選択（行・列の範囲は省くと 1 つ） */
export type SelectionInput = {
    type: 'row';
    r: number;
    r2?: number;
} | {
    type: 'column';
    c: number;
    c2?: number;
} | {
    type: 'cell';
    r: number;
    c: number;
};
/** 選択が見る表示の状態（`sheet-view.ts` が渡す。どれも呼んだ時点の値） */
export interface SelectionHost {
    /** 表示の根（`div.xs-sheet`。読み上げの要素を置く） */
    readonly root: HTMLElement;
    /** スクロール枠（`div.xs-body`。クリック・キーを受ける） */
    readonly body: HTMLElement;
    /** 帯を置く入れ物（`div.xs-wrap`）と、帯をその前に置く要素（図の層） */
    readonly wrapper: HTMLElement;
    readonly before: Element;
    sheet(): XlsxSheet | null;
    geometry(): Geometry | null;
    /** 行番号の列の幅・列見出しの高さ */
    origin(): {
        x: number;
        y: number;
    };
    /** 表の行数・列数（帯はこの範囲で切る） */
    extent(): {
        rows: number;
        cols: number;
    };
    table(): HTMLTableElement | null;
    /** 今のスクロール位置で描く（描いた範囲に収まっていれば何もしない） */
    draw(): void;
    /** 選択が変わった（同じ選択では呼ばない） */
    changed(selection: Selection | null, origin: SelectionOrigin): void;
    /** コピーするセルの数が上限を超えた（通知の帯に出す） */
    copyTooLarge(): void;
    /** 列の記号（列記号の見出しを探す。selection-gate.ts が渡す） */
    readonly colName: typeof colName;
    /** シートの寸法（コピーで非表示の行・列を見る。selection-gate.ts が渡す） */
    readonly createGeometry: typeof createGeometry;
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
    /** 選択をコピーする文字（タブ区切り・`
  `）。選択が無い・上限を超えると null */
    text(): string | null;
    /** 今の選択を、知らせずに外す前に知らせる（シートの切り替え）。選択が無ければ何もしない */
    clear(): void;
    dispose(): void;
}
/**
 * 行 r・列 c を含む結合の範囲（無ければそのセル 1 つ）。シートの行を上へ辿って、結合の左上を探す。
 * `getRows` は範囲より上で始まる結合の左上を範囲の最初の行に置くので、`MERGE_LOOKBACK` 行より上で始まる結合は、そこからの範囲になる。
 */
export declare function mergeRangeAt(sheet: Pick<XlsxSheet, 'rowCount' | 'getRows'>, r: number, c: number): CellRange;
/**
 * 選択を作る。`mode` が無ければクリックでは選ばない（`set` だけが効く）。
 * @param mode セルをクリックしたときに選ぶ物
 */
export declare function createSelection(host: SelectionHost, mode: SelectionMode | undefined): SelectionController;
