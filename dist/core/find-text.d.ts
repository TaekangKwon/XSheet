import type { XlsxSheet } from './types';
/** `findText` の指定 */
export interface FindTextOptions {
    /** `'partial'`＝文字の一部に一致・`'cell'`＝セルの文字全体が一致（既定 `'partial'`） */
    match?: 'partial' | 'cell';
    /** 英字などの大文字と小文字を区別しない（既定 true）。全角と半角はいつも区別する */
    ignoreCase?: boolean;
    /** セルの中の改行（`\r\n`・`\n`・`\r`）を照合のときだけ除く（既定 true）。位置は元の文字で返す */
    ignoreNewlines?: boolean;
    /** 非表示の行・列（高さ・幅 0 の行・列も）を調べない（既定 true） */
    skipHidden?: boolean;
    /** 中断（`AbortController` の `signal`）。中断されたら `AbortError`（`DOMException`）で reject する */
    signal?: {
        readonly aborted: boolean;
    };
    /** 返す一致の数の上限（既定 10,000。100,000 より大きい値は 100,000 にする） */
    maxMatches?: number;
}
/** 1 つの一致 */
export interface TextMatch {
    /** 行（1 始まり） */
    r: number;
    /** 列（1 始まり） */
    c: number;
    /** 一致の始まり（`CellView.text` 上の UTF-16 の位置） */
    start: number;
    /** 一致の終わり（含まない） */
    end: number;
}
/** `findText` の結果 */
export interface FindTextResult {
    /** 一致（行 → 列 → セルの中の位置の順） */
    matches: TextMatch[];
    /** `maxMatches` で打ち切ったか（上限を超える一致があった） */
    truncated: boolean;
}
/**
 * シートの中の文字を探す。表示形式を当てた後の文字（`CellView.text`）で照らす。
 * - 1 セルの中の重ならない一致をすべて返す。結合したセルは左上のセルだけ。空の検索語は空の結果。
 * - 図形・グラフの中の文字は探さない。
 * - 約 10ms ごとに休むので、大きなシートでも画面は固まらない。
 * - ブック全体を探すときは、呼ぶ側がシートを 1 つずつ `openSheet` で開いて呼ぶ。
 * @param sheet `XlsxBook.openSheet` で開いたシート
 * @param query 検索語
 * @throws DOMException `signal` で中断された（`name` は `'AbortError'`）
 */
export declare function findText(sheet: XlsxSheet, query: string, options?: FindTextOptions): Promise<FindTextResult>;
