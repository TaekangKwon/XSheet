/**
 * 拡大縮小（M7-7・Moobius の依頼 R5）。**寸法に倍率を掛けて描き直す**（CSS の `zoom`・`transform` で表を縮めない）。
 * - `Geometry` を倍率で包む。境目を整数に丸める（`top′(r) = round(top(r) × z)`・高さは境目の差）ので、
 *   罫線の層（SVG）と td の線がずれず、誤差もたまらない。線の幅（td の 1px の枠・罫線）には倍率を掛けない。
 * - 見出し（列見出しの高さ・行番号の列の幅・見出しの文字）には倍率を掛けない（`sheet-view.ts` が原点を倍率なしで渡す）。
 * - 文字の大きさは CSS 変数 `--xs-zoom`（`style.css`）。図の層は中を 100% で描き、層に `transform: scale()` を掛ける。
 * - 幅に合わせる（fit）は、データのある列と図の右端までを表示枠の幅に合わせる。
 */
import type { Geometry } from '../core/index';
import type { DrawingBox, PxRect } from './drawings';
/** 倍率。数は 0.5〜2.0、`{ fit: 'width' }` はデータのある列と図の右端までを表示枠の幅に合わせる */
export type Zoom = number | {
    fit: 'width';
    min?: number;
    max?: number;
};
/** 倍率の下限・上限 */
export declare const ZOOM_MIN = 0.5;
export declare const ZOOM_MAX = 2;
/** 解決した倍率の指定（数は範囲に丸めた値。fit は範囲に丸めた下限・上限） */
export type ZoomSpec = {
    fixed: number;
} | {
    fit: {
        min: number;
        max: number;
    };
};
/**
 * 倍率の指定を確かめて解決する。数は 0.5〜2.0 に丸める。fit の `min`・`max` は数でなければ既定（0.5・2.0）、
 * 逆転していれば入れ替える。
 * @throws RangeError 数でない値（`NaN`・`Infinity` を含む）・`{ fit: 'width' }` でないもの
 */
export declare function resolveZoom(zoom: Zoom): ZoomSpec;
/**
 * 寸法を倍率で包む。境目（行の上端・列の左端）を整数に丸め、高さ・幅は境目の差にする。
 * `rowAt`・`colAt` は倍率で割って元の寸法で引いた後、前後の行（列）を丸めた境目と比べて合わせる。
 * 倍率 1 なら元の寸法をそのまま返す（今までと同じ道を通る）。
 */
export declare function zoomGeometry(g: Geometry, z: number): Geometry;
/** 図の出し入れの判定の範囲（倍率を掛けた座標）を、図の層の中の座標（100%）に戻す */
export declare function unzoomRect(rect: PxRect, z: number): PxRect;
/**
 * 幅に合わせるときの中身の幅（100% の px）：データのある列の右端と、図の右端の大きいほう。
 * @param dataWidth データのある列（1〜`colCount`）の幅の合計
 */
export declare function contentWidth(dataWidth: number, boxes: readonly DrawingBox[]): number;
/**
 * 幅に合わせる倍率。合わせ直さなくてよい（中身・枠の幅が無い、今の倍率との差が 1% 未満）ときは null。
 * @param available セルに使える表示枠の幅（px。行番号の列とスクロールバーを除く）
 * @param content 中身の幅（100% の px）
 * @param current 今の倍率
 */
export declare function fitZoom(available: number, content: number, range: {
    min: number;
    max: number;
}, current: number): number | null;
/** 表示枠の左上にあるセルと、そのセルの中の位置の割合（倍率を替えても同じセルを左上に残すため） */
export interface ScrollAnchor {
    r: number;
    fy: number;
    c: number;
    fx: number;
}
/** 今のスクロール位置（A1 の左上からの px）の左上のセル */
export declare function scrollAnchor(g: Geometry, scrollTop: number, scrollLeft: number): ScrollAnchor;
/** 左上のセルを、新しい寸法でのスクロール位置にする */
export declare function anchorScroll(g: Geometry, a: ScrollAnchor): {
    top: number;
    left: number;
};
export declare function scrollbarWidth(): number;
