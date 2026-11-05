/**
 * `.xlsx` の図（画像・図形など）をセルの上に重ねるための純ロジック（DOM 非依存・ISSUE-86）。
 * - アンカー（`lib/xlsx` の `DrawingAnchor`）を、A1 の左上を原点とする px の矩形にする（geometry から出す）。
 * - 図の最下行・最右列（表をそこまで広げる）・スクロール枠の可視範囲との交差判定を出す。
 * 位置は列幅・行高と同じ geometry から計算するため、表のセルと 1px もずれない。
 * グループの中の図（ISSUE-89）は、最も外側のグループのアンカーの矩形へ、グループの子の座標系（`chOff`・`chExt`）から線形に写して矩形にする。
 */
import type { AnchorPoint, DrawingAnchor, SheetDrawing, Geometry, DrawingGroup } from '../core/index';
import { type LineRender, type ShapeDrawn } from './shapes';
/** 1px あたりの EMU（Office の英語メートル単位。px＝EMU / 9525） */
export declare const EMU_PER_PX = 9525;
/** A1 の左上を原点とする矩形（px・整数） */
export interface PxRect {
    x: number;
    y: number;
    w: number;
    h: number;
}
/**
 * 配置済みの図（`index` は `SheetDrawing[]` の位置。重ね順はこの昇順＝drawing XML の順）。
 * 線（kind が line）の `rect` は、アンカーの矩形ではなく `line.bbox`（矢じり・太さ・U 字の張り出しを含む外接矩形）。
 * 形で描く図形（`shape`）の `rect` は、経路の外接矩形とアンカーの矩形を合わせた矩形（`shape.rect`。ISSUE-92）。
 */
export interface DrawingBox {
    index: number;
    rect: PxRect;
    /** 線（kind が line）の描画結果。描画側が再計算しない。線以外は持たない */
    line?: LineRender;
    /** 形で描く図形（kind が shape・textbox で、形を評価できたもの）の描画結果。形を評価できない図形（代替）・それ以外は持たない */
    shape?: ShapeDrawn;
}
/** アンカーの矩形の計算に使う geometry の部分（テストで差し替えやすいよう必要なものだけ） */
export type AnchorGeometry = Pick<Geometry, 'colLeft' | 'colWidth' | 'rowTop' | 'rowHeight'>;
/** 図の配置に使う geometry の部分（`AnchorGeometry` ＋ 線をシートの範囲に丸めるための全幅・全高） */
export type DrawingGeometry = AnchorGeometry & Pick<Geometry, 'totalWidth' | 'totalHeight'>;
/**
 * アンカーの点の x（px）：その列の左端＋ずれ。ずれが列の幅を超える時は列の右端で止める
 * （列幅の僅差で隣のセルへ飛び出さない。ADR-0026 §決定 8）。
 * geometry の番号は 1 始まり、アンカーの `col` は 0 始まり。
 */
export declare function anchorPointX(geo: AnchorGeometry, p: AnchorPoint): number;
/** アンカーの点の y（px）：その行の上端＋ずれ（行の高さを超える時は行の下端で止める） */
export declare function anchorPointY(geo: AnchorGeometry, p: AnchorPoint): number;
/**
 * グループの中の図の矩形（px・整数）を出す（ADR-0027 §決定 5）。
 * `anchor` は最も外側のグループのアンカーの矩形。グループの子の座標系（`chOff`・`chExt`）から親の矩形へ、外側から順に線形に写す：
 * `x = 親.x + (子.off.x − chOff.x) × 親.w / chExt.cx`（y・幅・高さも同じ）。入れ子のグループは、内側のグループ自身の `off`・`ext` を外側の座標系から写して、その矩形を次の親にする。
 * グループ自身の回転・反転は無視する。図の `xfrm` は回転前の矩形そのものなので、縦横の入れ替えはしない。
 * 端は四捨五入して幅・高さは端の差で出す。矩形は A1 の左・上へはみ出しても切らない（切ると線の傾き・文字の折り返しが変わる。はみ出しは
 * 線の `clampPath` とレイヤー（`overflow: hidden`）が隠す）。位置は ±MAX_PX、大きさは MAX_PX までで止める。座標系が壊れている・有限でない値になるときは null。
 */
export declare function groupChildRect(anchor: PxRect, group: DrawingGroup): PxRect | null;
/**
 * 図の矩形（px）。トップレベルの図はアンカーの矩形、グループの中の図は最も外側のグループのアンカーから写した矩形
 * （幅・高さは 0 や負のままで、大きさの判定はしない。判定は `layoutDrawings` が行う）。
 */
export declare function placeDrawing(geo: AnchorGeometry, drawing: SheetDrawing): PxRect | null;
/**
 * アンカーを A1 の左上基準の矩形（px）にする。大きさが 0 以下の図（非表示の行・列だけにかかる・逆向きの twoCell・
 * 大きさ 0 の ext）は null（出さない）。
 * - twoCell：`from`〜`to` の差。
 * - oneCell：`from` を起点に `cx`×`cy`。
 * - absolute：シートの左上からの `x`・`y` に `cx`×`cy`。
 */
export declare function anchorRect(geo: AnchorGeometry, anchor: DrawingAnchor): PxRect | null;
/**
 * 図の一覧を配置する。`index` を元の位置のまま持つので重ね順が保たれる。
 * - グループの中の図（`group` を持つ）：矩形は最も外側のグループのアンカーから写した値（`placeDrawing`）。`xfrm` は回転前の矩形なので、線の描画では縦横を入れ替えない。
 * - 図形（kind が shape・textbox で形を持つもの。ISSUE-92）：大きさが 0 以下の図は除く。形を評価できれば `rect` は経路の外接矩形とアンカーの矩形を合わせた矩形で、
 *   `shape` に描画結果を持つ。枠も塗りも文字も無い図形（Excel でも見えない）は除く。形を評価できない図形は、次の「線以外」と同じ（代替の描き方にする）。
 * - 線以外：大きさが 0 以下の図は除く。`rect` はアンカーの矩形（グループの中の図は写した矩形）。
 * - 線（kind が line）：アンカーの幅か高さが 0 でも残す（水平・垂直の線）。幅・高さが両方 0 以下のときだけ除く
 *   （負の値＝逆向きの twoCell は 0 とみなす）。`rect` は `renderLine` の外接矩形（シートの範囲に丸め済み）で、
 *   `line` に描画結果を持つ。外接矩形が空（シートの範囲の外へ追い出された）なら除く。
 *   `SheetDrawing.line` が無い線（壊れたデータ）は既定の直線として扱う。
 */
export declare function layoutDrawings(geo: DrawingGeometry, drawings: readonly SheetDrawing[]): DrawingBox[];
/**
 * 図を収めるのに必要な表の行数・列数（最下行・最右列。1 始まり）。図が無ければ 0。
 * `gridExtent` の `minRows`・`minCols` に渡して、図のある範囲まで表を広げる。
 *
 * @param geo 行数・列数の上限を Excel の最大にした geometry（`rowAt`・`colAt` が図の端を含む行・列へ届くように）
 */
export declare function drawingExtent(geo: Pick<Geometry, 'rowAt' | 'colAt'>, boxes: readonly DrawingBox[]): {
    rows: number;
    cols: number;
};
/**
 * スクロール枠の可視範囲（A1 の左上基準・px）に、上下左右 1 画面分の余白を足した範囲。
 * 表の原点（A1 の左上）はスクロール枠の内容の (`originX`, `originY`)＝行番号の列の幅・見出し行の高さ。
 * 枠の大きさが 0（まだ DOM に付いていない）なら幅・高さ 0 の範囲を返す（何も交差しない）。
 */
export declare function viewWindow(scrollLeft: number, scrollTop: number, clientWidth: number, clientHeight: number, originX: number, originY: number): PxRect;
/** 2 つの矩形が交わるか（接するだけは交わらない。幅・高さ 0 の矩形は何とも交わらない） */
export declare function rectsIntersect(a: PxRect, b: PxRect): boolean;
/** 範囲 `win` と交わる図だけを、元の順（重ね順）で返す。図は最大 1,000 個なので毎回の線形走査でよい */
export declare function pickVisible(boxes: readonly DrawingBox[], win: PxRect): DrawingBox[];
