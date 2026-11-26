/** 読んだ styles.xml */
export interface Styles {
    /** セルの `s` 属性（cellXfs の位置）→ numFmtId */
    xfNumFmtIds: number[];
    /** 独自の表示形式（numFmtId → 書式コード）。Map なので `__proto__` などを入れても汚染されない */
    customFormats: Map<number, string>;
    /** 最大桁幅 MDW（px）。先頭の font（`<fonts>` の最初の `<font>`）から決める。読めなければ 7 */
    mdw: number;
    /** 先頭の font の既定の行高（px・96dpi）。実測表に無ければ null */
    fontRowPx: number | null;
}
/**
 * フォントの名前・サイズ（pt）→ 最大桁幅 MDW（px）。
 * 実測表（font-metrics.ts）から、サイズの間は線形補間・範囲外は比例で求め、表に無いフォントは Calibri の行を使う。
 * Calibri 11＝7px、游ゴシック 11＝8px、맑은 고딕 11＝8px。サイズが不正なら 7px。
 */
export declare function maxDigitWidth(name: string, size: number): number;
/** styles.xml を読む（無ければ空）。書式コードは実体参照を戻した文字列 */
export declare function parseStyles(xml: string | null): Styles;
