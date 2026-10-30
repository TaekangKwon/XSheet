import { parseHex6, sysClrValue, type Theme } from './theme';
export { parseHex6, sysClrValue };
/** 色：RGB の整数（0〜0xFFFFFF）と不透明度（0〜1） */
export interface Rgba {
    rgb: number;
    alpha: number;
}
/** 見つけた色の要素 */
export interface ColorElement {
    /** 要素名（接頭辞を除く。例 `srgbClr`） */
    name: string;
    /** 開きタグ全体（`parseAttrs` に渡せる） */
    open: string;
    /** 中身（子の変換）。自己閉じ・閉じタグの無いときは null */
    body: string | null;
}
/**
 * sRGB の成分（0〜1）と線形 RGB を行き来する式の指定。null は sRGB の区分式（IEC 61966-2-1）。
 * 数値なら単純なべき乗（γ 2.2・2.3 など）。`tint`・`shade`・`scrgbClr` の変換が使う（U7）。
 */
export declare const LINEAR_GAMMA: number | null;
/** sRGB の成分（0〜1）→ 線形 RGB（0〜1）。範囲外は収め、NaN は 0 */
export declare function toLinear(c: number, gamma?: number | null): number;
/** 線形 RGB（0〜1）→ sRGB の成分（0〜1）。範囲外は収め、NaN は 0 */
export declare function fromLinear(l: number, gamma?: number | null): number;
/**
 * path の `fill` の濃淡：`shade`（黒へ寄せる）か `tint`（白へ寄せる）の係数（0〜1）。
 * Excel の画素から決めた値：darken＝0.6・lighten＝0.6 は成分ごとの比がちょうど 0.6、darkenLess・lightenLess は 0.802〜0.809 に揃うので 0.805。
 * 濃淡は `applyFillMode` が **sRGB の値のまま**掛ける（線形 RGB にすると Excel の面より暗く／明るくなりすぎる）。
 */
export declare const FILL_MODE_FACTORS: ReadonlyMap<string, {
    kind: 'shade' | 'tint';
    factor: number;
}>;
/** プリセット色の数（別名を含む。試験用） */
export declare const PRESET_COLOR_COUNT: number;
/** `prstClr` の名前から RGB の整数を引く。未知の名前は null */
export declare function lookupPresetColor(name: string): number | null;
/**
 * `xml`（例 `<a:solidFill>` の中身）で最初に見つかる色の要素を切り出す。無ければ null。
 * 閉じタグは `indexOf` で探す（閉じが無ければ body は null＝子の変換を読まない）。
 */
export declare function findColorElement(xml: string): ColorElement | null;
/**
 * 基本の色に、子の変換を出現順に掛ける（ADR-0028 §決定 6.3）。
 * 読む子は最初の `MAX_COLOR_TRANSFORMS` 個まで（無視する要素・不正な値も 1 個と数える）。
 * 各段の結果は 0〜1 に収める。値が整数でない変換は飛ばす。無視する変換（`comp`・`inv`・`gray`・`gamma`・`invGamma`・`red`/`green`/`blue` の系統）は色を変えない。
 */
export declare function applyTransforms(base: Rgba, body: string | null): Rgba;
/**
 * 切り出した色の要素 1 つを解く（基本の色 → 子の変換）。解けなければ null（例外にしない）。
 * 解けないときに次の出どころへ進むのは呼び出し側（ISSUE-95）。
 */
export declare function resolveColorElement(el: ColorElement, theme: Theme): Rgba | null;
/**
 * `xml`（`<a:solidFill>` の中身など）の最初の色の要素を解く。色の要素が無い・解けなければ null。
 * 色は RGB の整数と不透明度で返す（ファイルの文字列は返さない）。
 */
export declare function resolveColor(xml: string, theme: Theme): Rgba | null;
/**
 * path の `fill`（`norm`・`darken`・`darkenLess`・`lighten`・`lightenLess`）を色に掛ける。
 * 係数は `FILL_MODE_FACTORS`（U8）。未知の名前・`norm` は色のまま。不透明度は変えない。
 * 掛けるのは sRGB の値のまま（shade＝`c × k`・tint＝`1 − (1 − c) × k`）。変換の `shade`・`tint`（線形 RGB）とは別の式で、Excel の面の画素と合う方を採った。
 */
export declare function applyFillMode(color: Rgba, mode: string): Rgba;
/**
 * 2 色を **線形 RGB** で割合を混ぜる（`t`＝0 で a、1 で b。0〜1 に収め、NaN は 0）。不透明度は sRGB と同じ線形補間。
 * Excel の 2 stop のグラデーションの中央の色がこの混色と合う（ISSUE-95・U7。3 stop 以上は sRGB の混色 `mixColors` と合う）。
 */
export declare function mixColorsLinear(a: Rgba, b: Rgba, t: number): Rgba;
/**
 * 2 色を割合で混ぜる（`t`＝0 で a、1 で b。0〜1 に収め、NaN は 0）。各成分は sRGB の値の線形補間で、不透明度も混ぜる。
 * グラデーションの補間・模様の混色に使う（適用は ISSUE-95）。
 */
export declare function mixColors(a: Rgba, b: Rgba, t: number): Rgba;
