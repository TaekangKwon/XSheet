/**
 * リッチテキストの区切り（`<si>`・`<is>` の `<r>`）と、その書式（`<rPr>`）の読み取り（M7-16）。
 * - `<rPr>` の中身は `<font>` と同じ読み方（`cell-format.ts` が渡す読み手）。
 * - 区切りの文字を連結したものが平文と違えば、区切りを持たない（平文で描く）。
 * - 走査は線形（`forEachElement`・`findElement`。閉じは indexOf）。DOM 非依存。
 */
import type { RunFont, TextRun } from './types';
/** `<rPr>` の中身 → 区切りの書式 */
export type RunFontReader = (rPrBody: string | null) => RunFont;
/** 区切りの書式を持ちうる断片か（`<rPr` が無ければ読まない） */
export declare function mayHaveRuns(inner: string): boolean;
/**
 * `<si>`・`<is>` の中身から区切りを読む。
 * @param inner 開きタグと閉じタグの間の文字列
 * @param text セルの文字（平文。`MAX_CELL_TEXT` で切った後）。区切りも同じ長さで切る
 * @param readFont `<rPr>` の読み手
 * @returns `<rPr>` のある区切りが無い・連結が平文と合わないときは undefined
 */
export declare function readRuns(inner: string, text: string, readFont: RunFontReader): TextRun[] | undefined;
