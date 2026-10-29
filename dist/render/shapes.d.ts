/**
 * `.xlsx` の線・矢印を SVG で描くための純ロジック（DOM 非依存・ISSUE-88）。
 * 流れ：アンカーの矩形 → 回転前の矩形 → 形の経路（prst・調整値）→ 反転 → 中心で時計回りに回転
 * → シートの範囲に丸める → 矢じり（端の向き）・線の端の縮め・外接矩形 → SVG 用の `d` 文字列。
 * 座標はすべて px（A1 の左上を原点、y は下向き）。`<marker>` は使わず、矢じりは多角形・楕円として計算する。
 * 描画（`createElementNS`）は後続の `image-layer.ts` が行い、ここは数値と数値だけの文字列を返す。
 * 図形の形（ISSUE-92・ADR-0028）も同じ写像（反転 → 回転 → 丸め）で扱う：`renderShape` が、形の評価器（geom の `evaluate.ts`）の経路を
 * 部分経路の列（閉じ＝`Z`・path ごとの塗り・線・面の濃淡）・外接矩形・文字領域にする。
 */
import type { SheetDrawing, DrawingLine, LineDash, LineEnd } from '../core/index';
import type { PxRect } from './drawings';
import type { PathSeg, Pt } from '../geom/index';
/**
 * 形の表と評価器（geom）を読む（M2-3）。表が大きいため render からは動的 import にし、形のある図形を描くシートを開く時だけ読む。
 * 何度呼んでも読むのは 1 回（約束を共有する）。読めなければ例外（呼び出し側は代替で描き、通知を出す）。
 */
export declare function loadGeom(): Promise<void>;
/** geom を読み込み済みか（未読なら `renderShape` は形のある図形も代替を返す） */
export declare function isGeomLoaded(): boolean;
/** 点（px）と経路の区間（直線 `L`・3 次ベジェ `C`）。定義は geom の `defs.ts`（評価器と同じ型） */
export type { PathSeg, Pt };
/**
 * 経路（部分経路）：始点と、そこから続く区間の列。
 * `closed` が true なら閉じる（`Z`。図形の経路）。線の経路は持たない（開いたまま）。
 */
export interface ShapePath {
    start: Pt;
    segs: PathSeg[];
    closed?: boolean;
}
/** シートの範囲（0〜全幅・0〜全高。geometry の `totalWidth`・`totalHeight`） */
export interface SheetBounds {
    totalWidth: number;
    totalHeight: number;
}
/** 回転・反転（`SheetDrawing` の `rot`［1/60000 度］・`flipH`・`flipV`。無ければ 0・なし） */
export interface Xfrm {
    rot?: number;
    flipH?: boolean;
    flipV?: boolean;
}
/** 矢じり 1 つの形（多角形か楕円。座標系は呼び出し元が決める） */
export type ArrowShape = {
    kind: 'polygon';
    /** どちらの端か（`head`＝経路の始点、`tail`＝終点） */
    end: 'head' | 'tail';
    points: Pt[];
    /** 塗る（`arrow`＝開いた V だけ false。線で描く） */
    filled: boolean;
    /** 閉じる（`Z`。開いた V だけ false） */
    closed: boolean;
} | {
    kind: 'ellipse';
    end: 'head' | 'tail';
    cx: number;
    cy: number;
    rx: number;
    ry: number;
    /** 長軸（rx 側）の向き（度。時計回り＝y 下向きの座標系） */
    angle: number;
};
/** 1 本の線の描画結果（`renderLine` の戻り値） */
export interface LineRender {
    /** 外接矩形（シートの原点基準・px・シートの範囲に丸め済み）。`<svg>` の位置・可視判定・`drawingExtent` に使う */
    bbox: PxRect;
    /** `<path d>`：外接矩形の左上を原点とした座標。線の端は矢じりの分だけ縮めてある */
    d: string;
    /** 線の太さ（px。1〜16） */
    strokeWidth: number;
    /** `stroke-dasharray` の数値（空なら実線） */
    dash: number[];
    /** 矢じり（始点・終点の順。無いものは出さない）。座標は外接矩形の左上を原点とする */
    heads: ArrowShape[];
    /** 経路（シートの原点基準・丸め済み・線の端は縮める前）。テスト・後続の判定用 */
    path: ShapePath;
}
/** 図形の枠線の miterlimit（SVG の既定値。`stroke-miterlimit` に明示し、外接矩形の余白もこれで決める。ISSUE-92） */
export declare const SHAPE_MITER_LIMIT = 4;
/**
 * アンカーの矩形から「回転前の矩形」を出す。図の回転が 45° 以上 135° 未満、または 225° 以上 315° 未満なら、
 * アンカー（回転後の外接）は縦横が入れ替わっているため、幅と高さを入れ替える。中心はアンカーと同じ（ADR-0027 §決定 4）。
 *
 * 45°・135°・225°・315° ちょうどの境界は、Excel のアンカーが縦横を入れ替えたか入れ替えていないかが回転の値だけでは決まらない（45° ちょうどの
 * 図形は入れ替えていない）。`ext`（`a:xfrm` の `a:ext`＝回転前の大きさ。px）があれば、アンカーの幅・高さを `ext` に近い向きへ合わせて決める
 * （入れ替えた方が `ext` に近ければ入れ替える。差が 1px 以内＝正方形に近いなど決められないときは、回転の値で決める。ISSUE-93）。
 *
 * @param rect アンカーの矩形（px。幅・高さは 0 でもよい）
 * @param rot 回転（1/60000 度。無ければ 0）
 * @param ext 回転前の大きさ（px。無ければ回転の値だけで決める）
 */
export declare function preRotationRect(rect: PxRect, rot: number | undefined, ext?: {
    w: number;
    h: number;
}): PxRect;
/** 図形の回転前の大きさ（`xfrmExt`。EMU）を px にする。持たない・読めないなら undefined（`preRotationRect` の `ext` に渡す） */
export declare function xfrmExtPx(drawing: Pick<SheetDrawing, 'xfrmExt'>): {
    w: number;
    h: number;
} | undefined;
/**
 * 形（`prst`）と調整値から、回転前の矩形（幅 w・高さ h・左上が原点）の中の経路を作る。式は ECMA-376 の preset 定義に合わせる。
 * - `line`・`straightConnector1`・未知の値：(0,0)→(w,h) の直線。
 * - `bentConnector2〜5`：折れ線。`curvedConnector4・5` は同じ調整値の `bentConnector4・5` の折れ線で代替する。
 * - `curvedConnector2・3`：3 次ベジェ。
 * 調整値は割合のまま使い（範囲外も張り出しとして許す）、丸めは回転後の点に対して行う（`clampPath`）。
 *
 * @param adj 調整値 adj1〜adj3（1/100000。無い要素は 50000）
 */
export declare function shapePath(geom: string, adj: readonly (number | undefined)[] | undefined, w: number, h: number): ShapePath;
/**
 * 矩形の中の経路を、反転（`flipH`：x→w-x、`flipV`：y→h-y）してから、矩形の中心で時計回りに `rot` 回転し、
 * シートの座標（px）にする。画面座標（y 下向き）では `x' = cx + dx·cos − dy·sin`、`y' = cy + dx·sin + dy·cos`。
 *
 * @param path 矩形の左上を原点とする経路（`shapePath` の結果）
 * @param box 回転前の矩形（`preRotationRect` の結果）
 */
export declare function transformPath(path: ShapePath, box: PxRect, xfrm: Xfrm): ShapePath;
/**
 * 経路の全点（制御点を含む）をシートの範囲（0〜全幅・0〜全高）に丸める。
 * 細工した極端な調整値で、座標・`<svg>` の大きさ・表の範囲が巨大にならないための安全弁（ADR-0027 §決定 4）。
 * 範囲の中の点は変えない（Excel の U 字線の張り出しは縮まない）。
 */
export declare function clampPath(path: ShapePath, sheet: SheetBounds): ShapePath;
/** 経路の全点（始点・制御点・終点）を並べる */
export declare function pathPoints(path: ShapePath): Pt[];
/** 線の太さ（px）：EMU / 9525 を 1〜16 に丸める。0・未指定・有限でない値は 1px */
export declare function strokeWidthPx(widthEmu: number | undefined): number;
/**
 * 線種を `stroke-dasharray` の数値にする。`solid` は空（実線）、`dot`・`sysDot` は点線 [w, 2w]、
 * それ以外は破線 [4w, 3w]（w は線の太さ px）。
 */
export declare function dashArray(dash: LineDash | string | undefined, strokeWidth: number): number[];
/** 矢じりの大きさ（px）：max(線の太さ, 2px) × 係数（`sm` 2・`med` 3・`lg` 5。未知は `med`）。上限 48px */
export declare function arrowDimension(strokeWidth: number, size: string | undefined): number;
/**
 * 端の接線の向き（単位ベクトル）：経路の外へ向かう向き。
 * `tail`＝終点で、一つ前の（終点と違う）点から終点への向き。`head`＝始点で、次の（始点と違う）点から始点への向き。
 * 直線・折れ線は端の線分、ベジェは端に隣る制御点→端点。長さ 0 なら更に遡る。全点が同じなら (1, 0)。
 */
export declare function endDirection(path: ShapePath, end: 'head' | 'tail'): Pt;
/**
 * 矢じり 1 つの形を、端点 `tip` と外向きの単位ベクトル `dir` から作る（座標はシートの原点基準）。
 * - `triangle`（と未知の値）：塗った三角。先端が端点。
 * - `stealth`：後ろが切れ込んだ塗った四角形。先端が端点。
 * - `arrow`：開いた V（線だけ）。先端が端点。
 * - `diamond`・`oval`：塗ったひし形・楕円。中心が端点。
 *
 * @param strokeWidth 線の太さ（px）。大きさ＝max(太さ, 2px)×係数（上限 48px）
 */
export declare function arrowShape(end: 'head' | 'tail', spec: LineEnd, tip: Pt, dir: Pt, strokeWidth: number): ArrowShape;
/**
 * 線の端を縮めた経路を返す（塗った三角・stealth の下に線が突き出ないため）。
 * 始点側は始点と最初の区間の制御点、終点側は終点と最後の区間の制御点を、端の接線に沿って動かす。
 * 縮める長さは、その端の区間の長さ（弦）を超えない。区間が 1 つで両端を縮めるときは、合計が区間の長さに収まるよう按分する。
 *
 * @param headShrink 始点側を縮める長さ（px）
 * @param tailShrink 終点側を縮める長さ（px）
 */
export declare function shrinkPath(path: ShapePath, headShrink: number, tailShrink: number): ShapePath;
/**
 * 経路を `<path d>` の文字列にする（数値だけから組み立てる。NaN・Infinity は出さない）。閉じた経路は末尾に `Z`。
 *
 * @param originX 座標の原点の x（この値を引く。外接矩形の左上を原点にするときに渡す）
 * @param originY 座標の原点の y
 */
export declare function pathData(path: ShapePath, originX?: number, originY?: number): string;
/** 部分経路の列を 1 つの `<path d>` にする（図形の path 1 つ。部分経路を空白でつなぐ。空なら空文字） */
export declare function subpathsData(paths: readonly ShapePath[], originX?: number, originY?: number): string;
/** 多角形の矢じりを `<path d>` の文字列にする（閉じる図形は `Z`）。座標はそのまま（原点の移動は `renderLine` が済ませる） */
export declare function arrowPolygonData(a: Extract<ArrowShape, {
    kind: 'polygon';
}>): string;
/**
 * 1 本の線を、SVG に描ける数値にまとめる。
 * アンカーの矩形 → 回転前の矩形 → 経路 → 反転 → 回転 → シートの範囲へ丸め → 矢じり・縮め・外接矩形。
 * 外接矩形は、丸めた経路の点（制御点を含む）と矢じりの範囲の外接矩形に、線の太さの半分と矢じりの大きさの半分
 * （大きい方）を四方に足し、シートの範囲に丸めたもの。水平・垂直の線（幅か高さが 0）でも太さ分の大きさを持つ。
 *
 * @param rect アンカーの矩形（px。回転後の外接。幅か高さが 0 でもよい）。`preRotated` が true なら回転前の矩形そのもの
 * @param line 線の形と見た目
 * @param xfrm 回転・反転
 * @param sheet シートの範囲
 * @param preRotated `rect` が回転前の矩形か（グループの中の線。`xfrm` の矩形そのもので、縦横を入れ替えない。ISSUE-89）
 */
export declare function renderLine(rect: PxRect, line: DrawingLine, xfrm: Xfrm, sheet: SheetBounds, preRotated?: boolean): LineRender;
/** 面の濃淡の種類（path の `fill` が `darken`・`darkenLess`・`lighten`・`lightenLess`。`norm`・`none` は膜なし） */
export type ShapeShade = 'darken' | 'darkenLess' | 'lighten' | 'lightenLess';
/** 図形の path 1 つ分の描画結果（`<path>` 1 つ） */
export interface ShapeFace {
    /** `<path d>`：外接矩形（`ShapeDrawn.rect`）の左上を原点とした座標。部分経路は空白でつなぎ、閉じた部分経路は `Z` */
    d: string;
    /** 塗る（Excel で塗りがあり、path の `fill` が `none` でない） */
    fill: boolean;
    /** 線を引く（Excel で枠線が見え、path の `stroke` が false でない） */
    stroke: boolean;
    /** 面の濃淡（塗るときだけ。無ければ null） */
    shade: ShapeShade | null;
    /** 部分経路（シートの原点基準・丸め済み）。テスト・後続の判定用 */
    subpaths: ShapePath[];
}
/** ダーク表示の下地の、縁取りの片側の幅（px。ADR-0028 §決定 6.7）。下地の svg はこの分だけ広げ、縁取りの太さは線の太さ＋この 2 倍 */
export declare const BACKDROP_PAD = 3;
/** 下地（ダーク表示の白）の path 1 つ：`fill`＝形の塗りとして白く塗る、`stroke`＝線の縁取り（線の太さ＋片側 3px の白い線） */
export interface ShapeBackdropPath {
    /** `<path d>`：面と同じ座標（外接矩形の左上が原点） */
    d: string;
    fill: boolean;
    stroke: boolean;
}
/** 形で描く図形の結果 */
export interface ShapeDrawn {
    kind: 'drawn';
    /** 外接矩形（シートの原点基準・px・整数）：経路の外接矩形（線の太さの半分を足し、丸めた点から出す）とアンカーの矩形を合わせたもの。可視判定・`drawingExtent`・DOM の位置に使う */
    rect: PxRect;
    /** 回転前の矩形（シートの原点基準。文字の枠を置く。回転は DOM が中心で掛ける） */
    box: PxRect;
    /** 文字領域（回転前の矩形の左上が原点。反転済み。文字そのものは反転しない） */
    textRect: PxRect;
    /** 描く path（何も塗らず線も引かない path は含めない。文字だけの図形は空） */
    faces: ShapeFace[];
    /**
     * ダーク表示の下地の path（ISSUE-96。ADR-0028 §決定 6.7）。塗りのある図形・塗りが無く文字のある図形は形の塗りの path（`fill` が `none` でない path）、
     * 枠線は線の縁取り。塗りも文字も無い図形は枠線の縁取りだけ。塗りの path が無く文字がある形は、回転前の矩形を回した形
     */
    backdrop: ShapeBackdropPath[];
    /** 枠線の太さ（px。1〜16） */
    strokeWidth: number;
    /** 枠線の線種（`stroke-dasharray` の数値。空なら実線） */
    dash: number[];
}
/**
 * 図形 1 つの描き方。
 * - `drawn`：形の経路で描く（`ShapeDrawn`）。
 * - `hidden`：枠も塗りも文字も無く、Excel でも見えない図形。描かない（図の数には入れる）。
 * - `fallback`：形を評価できない（`geom` が無い・表に無い `prst`・評価の失敗・上限超え）。ADR-0027 の描き方（矩形の文字枠・種類名の破線枠）にする。
 */
export type ShapeRender = ShapeDrawn | {
    kind: 'hidden';
} | {
    kind: 'fallback';
};
/** 図形が表示する文字を持つか（ランの文字が 1 文字以上ある。改行だけ・空のランは含めない） */
export declare function shapeHasText(drawing: SheetDrawing): boolean;
/** 塗りがあり、その色が見える（不透明度 0 超。色が無い＝暫定の色は見える）。下地の塗りの条件（ISSUE-96） */
export declare function filledShown(drawing: SheetDrawing): boolean;
/** 見える文字がある（文字があり、不透明度 0 超のランが 1 つ以上ある。`a:noFill` の文字は見えない）。下地の塗りの条件（ISSUE-96） */
export declare function textShown(drawing: SheetDrawing): boolean;
/**
 * 図形 1 つを、SVG に描ける数値にまとめる（ADR-0028 §決定 2・4・5・8）。
 * アンカーの矩形 → 回転前の矩形 → 形の経路（`evaluateGeom`）→ 反転 → 中心で回転 → シートの範囲に丸め → 外接矩形。
 * 外接矩形は、丸めた経路の点（制御点を含む）の外接矩形に、枠線があれば太さの半分を足して整数に広げ、アンカーの矩形と合わせたもの。
 * 枠も塗りも無く文字だけの図形は経路を作らず、外接矩形はアンカーの矩形。
 * 文字領域は形の `rect`（回転前の矩形の中）を、反転に合わせて矩形の中で鏡写しにしたもの（`flipH` なら x を w − r へ）。
 *
 * @param rect アンカーの矩形（px・整数。回転後の外接）。`preRotated` が true なら回転前の矩形そのもの
 * @param drawing 図形（`geom`・`outline`・`hasOutline`・`hasFill`・`rot`・`flipH`・`flipV`・`text`）
 * @param sheet シートの範囲
 * @param preRotated `rect` が回転前の矩形か（グループの中の図形。縦横を入れ替えない。ISSUE-89）
 */
export declare function renderShape(rect: PxRect, drawing: SheetDrawing, sheet: SheetBounds, preRotated?: boolean): ShapeRender;
