/**
 * `.xlsx` のセルの罫線を、SVG の path にするための純関数（DOM 非依存・ISSUE-98）。
 * - 描いた行（`GridRow.borders`）と罫線の表（`XlsxSheet.borders`）から、すべての横の辺・縦の辺（と斜線）について
 *   「隣り合う 2 つの指定」から 1 つの線を決め（太さの順位 → 線種 → 左・上のセル）、結合セルの内側の辺は描かず、
 *   同じ境目の上で同じ線が続く区間を 1 本にまとめて、線種×色ごとの path の `d` にする。
 * - 線種の幅・パターン・太さの順位は {@link BORDER_STYLES}（定数 1 か所）。Excel の画像との照合（E1〜E5）で直すのはここだけ。
 * - 座標は A1 の左上を原点とする px（geometry の `colLeft`・`rowTop` と同じ座標）。1px の線はグリッド線の画素に重ねる。
 */
import type { BorderEdge, BorderStyle, CellBorder, Geometry } from '../core/index';
import type { GridRow } from './grid';
/** 1 回の描画で作る区間（線 1 本・斜線 1 本）の数の上限。超えた分は描かない（ADR-0029 §決定 7） */
export declare const MAX_BORDER_SEGMENTS = 50000;
/** 線種 1 つの描き方 */
export interface BorderStyleSpec {
    /** 線の幅（px）。二重線は隙間を含む全体の幅（1px の線 2 本＋ 1px の隙間） */
    width: number;
    /** `stroke-dasharray` の数値（px）。実線は null */
    dash: readonly number[] | null;
    /** 太さの順位（大きいほど強い。辺の優先の 1 番目の基準） */
    weight: number;
    /** 線種の分類（同じ順位の中の優先。大きいほど強い） */
    kind: number;
    /** 二重線（隙間＋ 1px の線 2 本で描く） */
    double?: boolean;
}
/**
 * 13 線種の幅・パターン・優先の表（ADR-0029 §決定 3。暫定値で、E2 の照合で直す定数 1 か所）。
 * パターンの数値は px。`slantDashDot` の斜めの切り口は再現しない。
 */
export declare const BORDER_STYLES: Readonly<Record<BorderStyle, BorderStyleSpec>>;
/**
 * 同じ辺の 2 つの指定から、描く 1 つを選ぶ（ADR-0029 §決定 4）。
 * 片方だけならそれ。両方あれば、太さの順位が大きい方 → 同じなら線種の分類（double ＞ 実線 ＞ 破線系 ＞ dotted ＞ hair）
 * → それでも同じなら左（上）のセルの指定 `a`。線種が表に無い指定は無いものとして扱う。
 *
 * @param a 左（上）のセルの、この辺（右・下）の指定
 * @param b 右（下）のセルの、この辺（左・上）の指定
 */
export declare function pickBorderEdge(a: BorderEdge | undefined, b: BorderEdge | undefined): BorderEdge | undefined;
/** 1 本の path にまとめた線（線種×色×斜線かどうかごと） */
export interface BorderPath {
    /** 線種 */
    style: BorderStyle;
    /** 墨色（自動・黒）。色はテーマ変数で決める（`color` は null） */
    ink: boolean;
    /** 明示色（`#rrggbb`）。墨色なら null */
    color: string | null;
    /** 明示色の不透明度（0〜1。1 なら不透明） */
    alpha: number;
    /** `stroke-width`（px）。二重線は隙間の幅（3） */
    width: number;
    /** `stroke-dasharray`。実線・二重線は null */
    dasharray: string | null;
    /** 斜線の path か（縦横の線は画素にそろえて描くが、斜線はそろえない） */
    diagonal: boolean;
    /** path の `d`。二重線は中心線（隙間の描画に使う） */
    d: string;
    /** 二重線だけ：1px の線 2 本の `d`（中心線から 1px ずつ離した 2 本） */
    lineD?: string;
}
/** 描いた線の外接矩形（A1 の左上を 0 とする px。線の幅を含む） */
export interface BorderBounds {
    left: number;
    top: number;
    right: number;
    bottom: number;
}
/** {@link buildBorderPaths} の結果 */
export interface BorderPaths {
    paths: BorderPath[];
    /** 線が 1 本も無ければ null */
    bounds: BorderBounds | null;
    /** 作った区間の数 */
    segments: number;
    /** 区間の上限で打ち切ったか */
    truncated: boolean;
}
/**
 * 描いた範囲の罫線を、線種×色ごとの path にする。
 *
 * - 範囲の左端の縦の辺は左隣の列、右端の縦の辺は右隣の列のセル（範囲外でも `rows` は全列を持つ）の指定も見る。
 *   上端の横の辺は下のセルの上の辺だけ、下端の横の辺は上のセルの下の辺だけで決める。
 * - 結合セルは `GridRow` の左上のセルの `colSpan`・`rowSpan`（行の畳みで縮んだ見える行の数）から範囲を復元し、
 *   内側の辺は描かない。外周は面する構成セルの辺から取る。斜線は左上のセルの指定で結合範囲の対角に引く。
 * - 幅 0 の列・行の辺は描かない（`rows` は見える行だけ）。非表示の列・行をはさむ辺は、前後の見える列・行の指定で決める。
 * - 同じ境目の上で同じ線が続く区間は 1 本（`M x1 y H x2`・`M x y1 V y2`）にまとめる。区間の数が `maxSegments` を超えたら
 *   以後は描かない（`truncated`）。
 *
 * @param rows 描く行（見える行だけ・行番号の昇順。全列を持つ）
 * @param table 罫線の表（`XlsxSheet.borders`）
 * @param geo 寸法（`colLeft`・`colWidth`・`rowTop`）
 * @param range 描く列の範囲（1 始まり・両端を含む）
 * @param maxSegments 区間の数の上限（既定 {@link MAX_BORDER_SEGMENTS}）
 */
export declare function buildBorderPaths(rows: readonly GridRow[], table: readonly CellBorder[], geo: Pick<Geometry, 'colLeft' | 'colWidth' | 'rowTop'>, range: {
    first: number;
    last: number;
}, maxSegments?: number): BorderPaths;
