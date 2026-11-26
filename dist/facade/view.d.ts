/**
 * 1 行で使える入口 `view`・`read`（M6-2）。core（`openXlsx`）と render（`renderWorkbook`）の上の薄い層で、
 * 写したファイル（`src/core/`・`src/render/`）の作りには立ち入らない。
 * - `view`：xlsx をタブ付きで出す。表示の根 `div.xs-viewer` を container の子に足し、その中に `renderWorkbook` の表示を置く。
 *   言語・テーマ・CSS・高さを語で選べる。開けないファイルは今までどおり例外にせず、帯（`div.xs-banner`）に文言を出す。
 * - `read`：DOM なしで開く（`openXlsx` と同じ。受け取れる形を `view` と揃える）。
 */
import { type BuiltinDates, type SheetInfo, type XlsxBook, type XlsxLimits } from '../core/index';
import { type XSheetMessages } from '../render/index';
import { type XSheetSource } from './source';
import { type XSheetTheme } from './style';
/** 文言の言語。`auto` は `<html lang>`、無ければブラウザの言語。合うものが無ければ英語 */
export type XSheetLang = 'ja' | 'en' | 'ko' | 'auto';
/** `read` の指定 */
export interface ReadOptions {
    /** 読み取りの上限を替える（既定は `DEFAULT_LIMITS`） */
    limits?: Partial<XlsxLimits>;
    /** 組み込みの日付の書式 27〜36・50〜58 の見せ方（`OpenXlsxOptions.builtinDates` と同じ。ADR-0007 §6） */
    builtinDates?: BuiltinDates;
}
/** `view` の指定（どれも省略できる） */
export interface ViewOptions extends ReadOptions {
    /** 文言の言語（既定 `auto`） */
    lang?: XSheetLang;
    /** 文言を一部だけ替える（`lang` で選んだ表の上に重ねる） */
    messages?: Partial<XSheetMessages>;
    /** テーマ（既定 `light`）。`auto` は `prefers-color-scheme` に従う */
    theme?: XSheetTheme;
    /** 表示の CSS を自動で入れるか（既定 `auto`＝読み込み済みでなければ入れる。`none` は入れない） */
    css?: 'auto' | 'none';
    /** 表示の高さ（既定 `auto`＝container の高さが 0 のときだけ 480px）。数値は px、文字列は CSS の長さ */
    height?: number | string;
    /** 図の層（画像・図形・線）を出すか（既定 true） */
    drawings?: boolean;
    /** シートタブの列を置く場所（既定 `'top'`＝表の上。`'bottom'` は表の下） */
    tabPosition?: 'top' | 'bottom';
    /** 自動で入れる `<style>`・書式の `<style>` に付ける nonce（CSP の `style-src 'nonce-…'` 用） */
    nonce?: string;
}
/** `view` の戻り値 */
export interface Viewer {
    /** シートの一覧（タブの並び。非表示のシートも入る。開けなかったファイルでは空） */
    readonly sheets: readonly SheetInfo[];
    /**
     * シートを替える。
     * @param sheet シートの番号（`SheetInfo.index`・1 始まり）か名前
     * @throws RangeError ブックに無い番号・名前
     */
    showSheet(sheet: number | string): Promise<void>;
    /** 表示を外し、ブック・リスナ・blob URL を放す。何度呼んでもよい */
    destroy(): void;
}
/**
 * xlsx をタブ付きで表示する。
 * 同じ container に続けて呼ぶと、前の表示を外してから出す（読み込み中の前の呼び出しは、表示の無い Viewer を返す）。
 * @param target 表示を置く要素か CSS セレクタ（子として `div.xs-viewer` を足す）
 * @param source File・Blob・ArrayBuffer・型付き配列・Response、またはその Promise（`fetch(url)` をそのまま渡せる）
 * @throws TypeError 要素が見つからない・受けない形の source・失敗した Response
 */
export declare function view(target: HTMLElement | string, source: XSheetSource, options?: ViewOptions): Promise<Viewer>;
/**
 * xlsx を DOM なしで開く（`openXlsx` と同じ。閉じるのは呼び出し側の `close()`）。
 * @throws TypeError 受けない形の source・失敗した Response
 * @throws XlsxError 開けないファイル（`kind` で種類が分かる）
 */
export declare function read(source: XSheetSource, options?: ReadOptions): Promise<XlsxBook>;
