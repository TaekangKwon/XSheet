/**
 * フォントごとの寸法の実測表：最大桁幅 MDW（px）と、既定の行高（px）。
 * 出所: 2026-10-27・Windows GDI（TextRenderer）96dpi で数字 0〜9 の最大幅を実測。
 * Calibri 11＝7・游ゴシック 11＝8・맑은 고딕 11＝8・MS Pゴシック 11＝8 は、既知の Excel の値と一致する。
 * 既定の行高は Excel 16 の COM で確認した値（96dpi 換算）。
 */
/**
 * フォント名を表の名前にそろえる。NFKC 正規化（全角英字を半角に）・小文字・空白除去のあと、現地語名の別名を引く。
 * 表に無い名前は正規化しただけの名前を返す。
 */
export declare function canonicalFontName(name: string): string;
/** サイズの上限（pt）。Excel のフォントサイズの上限 */
export declare const MAX_FONT_SIZE = 409;
/**
 * フォントの名前・サイズ（pt）→ 最大桁幅 MDW（px・96dpi）。
 * 実測表のサイズの間は、近い 2 点から線形補間して四捨五入する。表の範囲外は、端の点に比例させる。
 * 表に無いフォントは Calibri の行を使う。サイズが不正（数でない・0 以下）なら null。
 */
export declare function lookupMdw(name: string, size: number): number | null;
/**
 * フォントの既定の行高（px・96dpi）。実測した（名前・サイズ）の組だけを持ち、無ければ null。
 * Calibri 11＝20、游ゴシック 11＝25、맑은 고딕 11＝22、MS Pゴシック 11＝18、Arial 10＝17。
 */
export declare function fontDefaultRowPx(name: string, size: number): number | null;
