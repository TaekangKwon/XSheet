/**
 * `.xlsx` の罫線層（ISSUE-98）：`svg.xs-borders` 1 枚の中身を、{@link buildBorderPaths} の結果から作る。
 * - `createElementNS` だけで作る（`innerHTML` は使わない）。`d`・`stroke-width`・`stroke-dasharray` は数値から組み立てた文字列だけ。
 * - 墨色（自動・黒）は class `is-ink` だけ（色は CSS のテーマ変数 `--xs-border-ink`。テーマの切替で描き直さない）。
 *   ほかの明示色は整数から組み立てた `#rrggbb` を CSSOM（`style.setProperty`）で入れる（CSP のため style 属性の文字列は組み立てない）。
 * - 二重線は、セルの面の色の隙間（`is-gap`・幅 3）の上に 1px の線 2 本。隙間は全 path の線より先（下）に描く。
 * - 位置と大きさは描いた線の外接矩形だけ（A1 の左上を原点とする座標を viewBox で保つ）。線が無ければ中身を空にする。
 */
import type { BorderPaths } from './borders';
/** 罫線層 */
export interface BorderLayer {
    /** wrapper に置く svg（常に 1 枚。中身だけ入れ替わる） */
    readonly element: SVGSVGElement;
    /**
     * 結果で中身を作り直す。線が無ければ空にする。
     * @param result {@link buildBorderPaths} の結果（座標は A1 の左上が原点）
     * @param originX A1 の左上の x（行番号の列の幅）
     * @param originY A1 の左上の y（見出し行の高さ）
     */
    render(result: BorderPaths, originX: number, originY: number): void;
    /** 中身を空にする（シートの切替・後始末） */
    clear(): void;
}
/** 罫線層を作る（svg は常に 1 枚で、中身だけを `render` で入れ替える） */
export declare function createBorderLayer(): BorderLayer;
