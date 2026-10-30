/** 共有文字列の表 */
export interface SharedStrings {
    readonly count: number;
    /** i 番目（0 始まり）の文字列。範囲外は空文字 */
    get(i: number): string;
}
/**
 * `<si>`・`<is>` の中身（XML の断片）を平文にする。ふりがなを除き、`<t>` を順に連結する。
 * @param inner 開きタグと閉じタグの間の文字列
 */
export declare function richTextToPlain(inner: string): string;
/** sharedStrings.xml を読む。`<si>` の位置の索引だけを作り、文字への展開は get で行う */
export declare function parseSharedStrings(xml: string | null): SharedStrings;
