/**
 * 図形・線・文字の色の読み取り（画面に依存しない）：色の出どころの順と、塗りの代表色。
 * - 塗り：`spPr` 直下の塗り → `a:fillRef` の子の色 → テーマの `lt1`。
 * - 枠線・線：`a:ln` の塗り → `a:lnRef` の子の色 → テーマの `tx1`。
 * - 文字：ランの `a:rPr` の塗り → `a:fontRef` の子の色 → テーマの `tx1`。
 * - 解けない色（不正な値・`phClr`・色の要素が無い）はその出どころを飛ばして次へ進む。`a:grpFill` は未指定として扱う。
 * - グラデーション・模様・画像の塗りは代表色の 1 色にする（`a:gradFill`＝50% の位置・`a:pattFill`＝前景と背景の混色・`a:blipFill`＝#D9D9D9）。
 * 塗り・枠線の「有無」は `drawing-shape.ts`（ADR-0027）が決める。ここは「ある」と決まった塗り・線の色を返す。
 * 色の要素そのものの解釈（変換・テーマ）は `color.ts`。結果は RGB の整数と不透明度だけで、ファイルの文字列を持たない。
 * 閉じの無い要素・大量の stop・変換の列でも例外にせず、走査は線形時間（上限は `limits.ts`）。
 */
import { type Rgba } from './color';
import { type FillEl, type ShapeProps } from './drawing-shape';
import { type Theme } from './theme';
/**
 * 塗りの要素 1 つの色（ADR-0028 §決定 6.4）。`a:solidFill`＝その色・`a:gradFill`・`a:pattFill`＝代表色・`a:blipFill`＝#D9D9D9。
 * `a:noFill`・`a:grpFill`・色の要素が無い・解けないときは null（次の出どころへ進む）。
 */
export declare function fillElementColor(el: FillEl, theme: Theme): Rgba | null;
/**
 * 図形の塗りの色：`spPr` 直下の塗り → `a:fillRef` の子の色 → テーマの `lt1`。
 * 塗りがあると決まった図形（`props.hasFill`）にだけ呼ぶ（有無はここで判定しない）。
 * `xdr:style` も `spPr` の塗りも無い図形は、有無の規則（ADR-0027）で「塗り・線あり」なので、色は既定の `lt1`（枠線・線は `tx1`）になる。ISSUE-95 の「既定」どおり。
 */
export declare function readFillColor(props: ShapeProps, theme: Theme): Rgba;
/**
 * 図形の枠線・線の色：`a:ln` の `a:solidFill`・`a:gradFill`・`a:pattFill` → `a:lnRef` の子の色 → テーマの `tx1`。
 * 枠線が見えると決まったもの（`props.hasOutline`）にだけ呼ぶ。矢じりも同じ色。
 */
export declare function readOutlineColor(props: ShapeProps, theme: Theme): Rgba;
/** 図形の文字の既定の色：`xdr:style` の `a:fontRef` の子の色 → テーマの `tx1`（`spBody` は `xdr:sp` の中身） */
export declare function readFontColor(spBody: string, theme: Theme): Rgba;
/**
 * 文字のラン（`a:r`・`a:fld`・`a:br`）の色：`a:rPr` の `a:solidFill`・`a:gradFill`（代表色）→ `fallback`（図形の `fontRef` の色。`readFontColor`）。
 * `a:noFill` は透明（不透明度 0）。解けない色・`a:pattFill` などは `fallback`。`a:ln`（文字の輪郭）・`a:highlight` は読まない。
 * @param rPrBody `a:rPr` の中身（`a:rPr` が無い・自己閉じのときは null）
 */
export declare function readRunColor(rPrBody: string | null, fallback: Rgba, theme: Theme): Rgba;
