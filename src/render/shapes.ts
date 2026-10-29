// 正: InterHub docs/plan/adr/0027-xlsx-shape-text-and-connectors.md（§決定 4）
/**
 * `.xlsx` の線・矢印を SVG で描くための純ロジック（DOM 非依存・ISSUE-88）。
 * 流れ：アンカーの矩形 → 回転前の矩形 → 形の経路（prst・調整値）→ 反転 → 中心で時計回りに回転
 * → シートの範囲に丸める → 矢じり（端の向き）・線の端の縮め・外接矩形 → SVG 用の `d` 文字列。
 * 座標はすべて px（A1 の左上を原点、y は下向き）。`<marker>` は使わず、矢じりは多角形・楕円として計算する。
 * 描画（`createElementNS`）は後続の `image-layer.ts` が行い、ここは数値と数値だけの文字列を返す。
 * 図形の形（ISSUE-92・ADR-0028）も同じ写像（反転 → 回転 → 丸め）で扱う：`renderShape` が、形の評価器（geom の `evaluate.ts`）の経路を
 * 部分経路の列（閉じ＝`Z`・path ごとの塗り・線・面の濃淡）・外接矩形・文字領域にする。
 */
import type { SheetDrawing, DrawingLine, LineDash, LineEnd, ShapeGeom } from '../core/index';
import type { PxRect } from './drawings';
import type { GeomDef, PathSeg, Pt } from '../geom/index';

/** geom の入口（形の表 `preset-geom.json`・約 171 KB と式の評価器）。{@link loadGeom} で読む */
type GeomModule = typeof import('../geom/index');

/** 読み込んだ geom（読む前・失敗は null。null の間、形のある図形は代替＝種類名の破線枠で描く） */
let geomLib: GeomModule | null = null;
/** 読み込み中・読み込み済みの約束（失敗したら捨て、次の呼び出しで読み直す） */
let geomLoading: Promise<void> | null = null;

/**
 * 形の表と評価器（geom）を読む（M2-3）。表が大きいため render からは動的 import にし、形のある図形を描くシートを開く時だけ読む。
 * 何度呼んでも読むのは 1 回（約束を共有する）。読めなければ例外（呼び出し側は代替で描き、通知を出す）。
 */
export function loadGeom(): Promise<void> {
  if (!geomLoading) {
    geomLoading = import('../geom/index').then(
      (lib) => {
        geomLib = lib;
      },
      (e: unknown) => {
        geomLoading = null;
        throw e;
      },
    );
  }
  return geomLoading;
}

/** geom を読み込み済みか（未読なら `renderShape` は形のある図形も代替を返す） */
export function isGeomLoaded(): boolean {
  return geomLib !== null;
}

/** 1px あたりの EMU（`drawings.ts` の `EMU_PER_PX` と同じ値。循環 import を避けるためここにも持つ） */
const EMU_PER_PX = 9525;

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
export type ArrowShape =
  | {
      kind: 'polygon';
      /** どちらの端か（`head`＝経路の始点、`tail`＝終点） */
      end: 'head' | 'tail';
      points: Pt[];
      /** 塗る（`arrow`＝開いた V だけ false。線で描く） */
      filled: boolean;
      /** 閉じる（`Z`。開いた V だけ false） */
      closed: boolean;
    }
  | {
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

/** 調整値の既定（`adj` が無い・有限でないとき。割合 50%） */
const DEFAULT_ADJ = 50000;
/** 調整値の分母（1/100000 の割合） */
const ADJ_DEN = 100000;
/** 線の太さの範囲（px） */
const MIN_STROKE = 1;
const MAX_STROKE = 16;
/** 矢じりの大きさの下限の基準（線の太さがこれより細くても、この値に係数を掛ける）と上限（px） */
const MIN_ARROW_BASE = 2;
const MAX_ARROW = 48;
/** 図形の枠線の miterlimit（SVG の既定値。`stroke-miterlimit` に明示し、外接矩形の余白もこれで決める。ISSUE-92） */
export const SHAPE_MITER_LIMIT = 4;
/** 矢じりの大きさの係数（`sm`・`med`・`lg`） */
const ARROW_FACTOR = { sm: 2, med: 3, lg: 5 } as const;
/** `stealth` の切れ込みの位置（先端から長さの何割後ろか） */
const STEALTH_NOTCH = 0.7;
/** 1/60000 度あたりの 1 周（360 度）・各境界 */
const ROT_FULL = 360 * 60000;
const ROT_45 = 45 * 60000;
const ROT_135 = 135 * 60000;
const ROT_225 = 225 * 60000;
const ROT_315 = 315 * 60000;
/** 同じ点とみなす距離（px）。端の接線の向きを求めるとき、これ以下の差の点は飛ばす */
const SAME_POINT = 1e-6;

/** 有限でない値を 0 にする */
function fin(n: number): number {
  return Number.isFinite(n) ? n : 0;
}

/** 回転（1/60000 度）を [0, 360 度) に正規化する（有限でない値は 0） */
function normRot(rot: number | undefined): number {
  if (rot === undefined || !Number.isFinite(rot)) {
    return 0;
  }
  return ((Math.round(rot) % ROT_FULL) + ROT_FULL) % ROT_FULL;
}

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
export function preRotationRect(
  rect: PxRect,
  rot: number | undefined,
  ext?: { w: number; h: number },
): PxRect {
  const r = normRot(rot);
  let swap = (r >= ROT_45 && r < ROT_135) || (r >= ROT_225 && r < ROT_315);
  if (ext !== undefined && ext.w > 0 && ext.h > 0) {
    const keep = Math.abs(rect.w - ext.w) + Math.abs(rect.h - ext.h);
    const swapped = Math.abs(rect.w - ext.h) + Math.abs(rect.h - ext.w);
    if (Math.abs(keep - swapped) > 1) {
      swap = swapped < keep;
    }
  }
  if (!swap) {
    return { ...rect };
  }
  const cx = rect.x + rect.w / 2;
  const cy = rect.y + rect.h / 2;
  return { x: cx - rect.h / 2, y: cy - rect.w / 2, w: rect.h, h: rect.w };
}

/** 図形の回転前の大きさ（`xfrmExt`。EMU）を px にする。持たない・読めないなら undefined（`preRotationRect` の `ext` に渡す） */
export function xfrmExtPx(
  drawing: Pick<SheetDrawing, 'xfrmExt'>,
): { w: number; h: number } | undefined {
  const e = drawing.xfrmExt;
  return e === undefined ? undefined : { w: e.cx / EMU_PER_PX, h: e.cy / EMU_PER_PX };
}

/** 調整値 `i`（0 始まり）を割合（1.0＝100%）にする。無い・有限でないときは 50% */
function adjRatio(adj: readonly (number | undefined)[] | undefined, i: number): number {
  const v = adj?.[i];
  return (v !== undefined && Number.isFinite(v) ? v : DEFAULT_ADJ) / ADJ_DEN;
}

/**
 * 形（`prst`）と調整値から、回転前の矩形（幅 w・高さ h・左上が原点）の中の経路を作る。式は ECMA-376 の preset 定義に合わせる。
 * - `line`・`straightConnector1`・未知の値：(0,0)→(w,h) の直線。
 * - `bentConnector2〜5`：折れ線。`curvedConnector4・5` は同じ調整値の `bentConnector4・5` の折れ線で代替する。
 * - `curvedConnector2・3`：3 次ベジェ。
 * 調整値は割合のまま使い（範囲外も張り出しとして許す）、丸めは回転後の点に対して行う（`clampPath`）。
 *
 * @param adj 調整値 adj1〜adj3（1/100000。無い要素は 50000）
 */
export function shapePath(
  geom: string,
  adj: readonly (number | undefined)[] | undefined,
  w: number,
  h: number,
): ShapePath {
  const L = (x: number, y: number): PathSeg => ({ t: 'L', x, y });
  const a1 = adjRatio(adj, 0);
  const a2 = adjRatio(adj, 1);
  const a3 = adjRatio(adj, 2);
  switch (geom) {
    case 'bentConnector2':
      return { start: { x: 0, y: 0 }, segs: [L(w, 0), L(w, h)] };
    case 'bentConnector3': {
      const x1 = w * a1;
      return { start: { x: 0, y: 0 }, segs: [L(x1, 0), L(x1, h), L(w, h)] };
    }
    case 'bentConnector4':
    case 'curvedConnector4': {
      const x1 = w * a1;
      const y2 = h * a2;
      return { start: { x: 0, y: 0 }, segs: [L(x1, 0), L(x1, y2), L(w, y2), L(w, h)] };
    }
    case 'bentConnector5':
    case 'curvedConnector5': {
      const x1 = w * a1;
      const y2 = h * a2;
      const x3 = w * a3;
      return {
        start: { x: 0, y: 0 },
        segs: [L(x1, 0), L(x1, y2), L(x3, y2), L(x3, h), L(w, h)],
      };
    }
    case 'curvedConnector2':
      return {
        start: { x: 0, y: 0 },
        segs: [{ t: 'C', x1: w / 2, y1: 0, x2: w, y2: h / 2, x: w, y: h }],
      };
    case 'curvedConnector3': {
      const x2 = w * a1;
      return {
        start: { x: 0, y: 0 },
        segs: [
          { t: 'C', x1: x2 / 2, y1: 0, x2, y2: h / 4, x: x2, y: h / 2 },
          { t: 'C', x1: x2, y1: (h * 3) / 4, x2: (w + x2) / 2, y2: h, x: w, y: h },
        ],
      };
    }
    default:
      // line・straightConnector1・未知の値（custGeom など）は直線で代替
      return { start: { x: 0, y: 0 }, segs: [L(w, h)] };
  }
}

/** 経路の点（制御点を含む）を全部 `fn` で写した新しい経路にする */
function mapPath(path: ShapePath, fn: (p: Pt) => Pt): ShapePath {
  const out: ShapePath = {
    start: fn(path.start),
    segs: path.segs.map((s): PathSeg => {
      if (s.t === 'L') {
        const p = fn(s);
        return { t: 'L', x: p.x, y: p.y };
      }
      const c1 = fn({ x: s.x1, y: s.y1 });
      const c2 = fn({ x: s.x2, y: s.y2 });
      const e = fn(s);
      return { t: 'C', x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: e.x, y: e.y };
    }),
  };
  // 閉じた部分経路は閉じを保つ（開いた線の経路には `closed` を足さない）
  if (path.closed) {
    out.closed = true;
  }
  return out;
}

/**
 * 矩形の中の経路を、反転（`flipH`：x→w-x、`flipV`：y→h-y）してから、矩形の中心で時計回りに `rot` 回転し、
 * シートの座標（px）にする。画面座標（y 下向き）では `x' = cx + dx·cos − dy·sin`、`y' = cy + dx·sin + dy·cos`。
 *
 * @param path 矩形の左上を原点とする経路（`shapePath` の結果）
 * @param box 回転前の矩形（`preRotationRect` の結果）
 */
export function transformPath(path: ShapePath, box: PxRect, xfrm: Xfrm): ShapePath {
  const rad = (normRot(xfrm.rot) / 60000) * (Math.PI / 180);
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  return mapPath(path, (p) => {
    const lx = xfrm.flipH ? box.w - p.x : p.x;
    const ly = xfrm.flipV ? box.h - p.y : p.y;
    const dx = box.x + lx - cx;
    const dy = box.y + ly - cy;
    return { x: cx + dx * cos - dy * sin, y: cy + dx * sin + dy * cos };
  });
}

/**
 * 経路の全点（制御点を含む）をシートの範囲（0〜全幅・0〜全高）に丸める。
 * 細工した極端な調整値で、座標・`<svg>` の大きさ・表の範囲が巨大にならないための安全弁（ADR-0027 §決定 4）。
 * 範囲の中の点は変えない（Excel の U 字線の張り出しは縮まない）。
 */
export function clampPath(path: ShapePath, sheet: SheetBounds): ShapePath {
  const maxX = Math.max(0, fin(sheet.totalWidth));
  const maxY = Math.max(0, fin(sheet.totalHeight));
  return mapPath(path, (p) => ({
    x: Math.min(maxX, Math.max(0, fin(p.x))),
    y: Math.min(maxY, Math.max(0, fin(p.y))),
  }));
}

/** 経路の全点（始点・制御点・終点）を並べる */
export function pathPoints(path: ShapePath): Pt[] {
  const pts: Pt[] = [path.start];
  for (const s of path.segs) {
    if (s.t === 'C') {
      pts.push({ x: s.x1, y: s.y1 }, { x: s.x2, y: s.y2 });
    }
    pts.push({ x: s.x, y: s.y });
  }
  return pts;
}

/** 線の太さ（px）：EMU / 9525 を 1〜16 に丸める。0・未指定・有限でない値は 1px */
export function strokeWidthPx(widthEmu: number | undefined): number {
  if (widthEmu === undefined || !Number.isFinite(widthEmu) || widthEmu <= 0) {
    return MIN_STROKE;
  }
  return Math.min(MAX_STROKE, Math.max(MIN_STROKE, widthEmu / EMU_PER_PX));
}

/**
 * 線種を `stroke-dasharray` の数値にする。`solid` は空（実線）、`dot`・`sysDot` は点線 [w, 2w]、
 * それ以外は破線 [4w, 3w]（w は線の太さ px）。
 */
export function dashArray(dash: LineDash | string | undefined, strokeWidth: number): number[] {
  if (dash === undefined || dash === 'solid') {
    return [];
  }
  if (dash === 'dot' || dash === 'sysDot') {
    return [strokeWidth, strokeWidth * 2];
  }
  return [strokeWidth * 4, strokeWidth * 3];
}

/** 矢じりの大きさ（px）：max(線の太さ, 2px) × 係数（`sm` 2・`med` 3・`lg` 5。未知は `med`）。上限 48px */
export function arrowDimension(strokeWidth: number, size: string | undefined): number {
  const f = size === 'sm' || size === 'lg' ? ARROW_FACTOR[size] : ARROW_FACTOR.med;
  return Math.min(MAX_ARROW, Math.max(strokeWidth, MIN_ARROW_BASE) * f);
}

/** `a` から `b` への距離 */
function dist(a: Pt, b: Pt): number {
  return Math.hypot(b.x - a.x, b.y - a.y);
}

/**
 * 端の接線の向き（単位ベクトル）：経路の外へ向かう向き。
 * `tail`＝終点で、一つ前の（終点と違う）点から終点への向き。`head`＝始点で、次の（始点と違う）点から始点への向き。
 * 直線・折れ線は端の線分、ベジェは端に隣る制御点→端点。長さ 0 なら更に遡る。全点が同じなら (1, 0)。
 */
export function endDirection(path: ShapePath, end: 'head' | 'tail'): Pt {
  const pts = pathPoints(path);
  const tip = end === 'head' ? pts[0]! : pts[pts.length - 1]!;
  const order = end === 'head' ? pts.slice(1) : pts.slice(0, -1).reverse();
  for (const p of order) {
    const d = dist(p, tip);
    if (d > SAME_POINT) {
      return { x: (tip.x - p.x) / d, y: (tip.y - p.y) / d };
    }
  }
  return { x: 1, y: 0 };
}

/** 矢じりの型ごとの、線の端を縮める長さ（矢じりの長さに対する割合）。塗り三角は全長、stealth は切れ込みまで、他は縮めない */
function shrinkRatio(type: string): number {
  switch (type) {
    case 'arrow':
    case 'diamond':
    case 'oval':
      return 0;
    case 'stealth':
      return STEALTH_NOTCH;
    default:
      return 1;
  }
}

/** 端の矢じりの大きさ（幅・長さ。px） */
function arrowSize(end: LineEnd, strokeWidth: number): { width: number; length: number } {
  return {
    width: arrowDimension(strokeWidth, end.w),
    length: arrowDimension(strokeWidth, end.len),
  };
}

/**
 * 矢じり 1 つの形を、端点 `tip` と外向きの単位ベクトル `dir` から作る（座標はシートの原点基準）。
 * - `triangle`（と未知の値）：塗った三角。先端が端点。
 * - `stealth`：後ろが切れ込んだ塗った四角形。先端が端点。
 * - `arrow`：開いた V（線だけ）。先端が端点。
 * - `diamond`・`oval`：塗ったひし形・楕円。中心が端点。
 *
 * @param strokeWidth 線の太さ（px）。大きさ＝max(太さ, 2px)×係数（上限 48px）
 */
export function arrowShape(
  end: 'head' | 'tail',
  spec: LineEnd,
  tip: Pt,
  dir: Pt,
  strokeWidth: number,
): ArrowShape {
  const { width, length } = arrowSize(spec, strokeWidth);
  // 局所座標 (a, b)：a は外向き（先端側が正）、b は直交（左右）。(0,0) が端点
  const at = (a: number, b: number): Pt => ({
    x: tip.x + dir.x * a - dir.y * b,
    y: tip.y + dir.y * a + dir.x * b,
  });
  const hw = width / 2;
  switch (spec.type) {
    case 'stealth':
      return {
        kind: 'polygon',
        end,
        points: [at(0, 0), at(-length, hw), at(-length * STEALTH_NOTCH, 0), at(-length, -hw)],
        filled: true,
        closed: true,
      };
    case 'arrow':
      return {
        kind: 'polygon',
        end,
        points: [at(-length, hw), at(0, 0), at(-length, -hw)],
        filled: false,
        closed: false,
      };
    case 'diamond':
      return {
        kind: 'polygon',
        end,
        points: [at(length / 2, 0), at(0, hw), at(-length / 2, 0), at(0, -hw)],
        filled: true,
        closed: true,
      };
    case 'oval':
      return {
        kind: 'ellipse',
        end,
        cx: tip.x,
        cy: tip.y,
        rx: length / 2,
        ry: hw,
        angle: (Math.atan2(dir.y, dir.x) * 180) / Math.PI,
      };
    default:
      // triangle と未知の値は塗った三角
      return {
        kind: 'polygon',
        end,
        points: [at(0, 0), at(-length, hw), at(-length, -hw)],
        filled: true,
        closed: true,
      };
  }
}

/** 矢じりの座標を (dx, dy) だけ動かす */
function shiftArrow(a: ArrowShape, dx: number, dy: number): ArrowShape {
  if (a.kind === 'ellipse') {
    return { ...a, cx: a.cx + dx, cy: a.cy + dy };
  }
  return { ...a, points: a.points.map((p) => ({ x: p.x + dx, y: p.y + dy })) };
}

/** 矢じりの外接矩形（左・上・右・下） */
function arrowExtent(a: ArrowShape): { l: number; t: number; r: number; b: number } {
  if (a.kind === 'ellipse') {
    const rad = (a.angle * Math.PI) / 180;
    const ex = Math.hypot(a.rx * Math.cos(rad), a.ry * Math.sin(rad));
    const ey = Math.hypot(a.rx * Math.sin(rad), a.ry * Math.cos(rad));
    return { l: a.cx - ex, t: a.cy - ey, r: a.cx + ex, b: a.cy + ey };
  }
  const xs = a.points.map((p) => p.x);
  const ys = a.points.map((p) => p.y);
  return { l: Math.min(...xs), t: Math.min(...ys), r: Math.max(...xs), b: Math.max(...ys) };
}

/**
 * 線の端を縮めた経路を返す（塗った三角・stealth の下に線が突き出ないため）。
 * 始点側は始点と最初の区間の制御点、終点側は終点と最後の区間の制御点を、端の接線に沿って動かす。
 * 縮める長さは、その端の区間の長さ（弦）を超えない。区間が 1 つで両端を縮めるときは、合計が区間の長さに収まるよう按分する。
 *
 * @param headShrink 始点側を縮める長さ（px）
 * @param tailShrink 終点側を縮める長さ（px）
 */
export function shrinkPath(path: ShapePath, headShrink: number, tailShrink: number): ShapePath {
  if (path.segs.length === 0 || (headShrink <= 0 && tailShrink <= 0)) {
    return path;
  }
  const segs = path.segs.map((s) => ({ ...s }));
  const start = { ...path.start };
  const firstSeg = segs[0]!;
  const lastSeg = segs[segs.length - 1]!;
  const firstEnd: Pt = { x: firstSeg.x, y: firstSeg.y };
  const prevOfLast: Pt =
    segs.length >= 2
      ? { x: segs[segs.length - 2]!.x, y: segs[segs.length - 2]!.y }
      : { x: start.x, y: start.y };
  const headLen = dist(start, firstEnd);
  const tailLen = dist(prevOfLast, { x: lastSeg.x, y: lastSeg.y });
  let sh = Math.min(Math.max(0, headShrink), headLen);
  let st = Math.min(Math.max(0, tailShrink), tailLen);
  if (segs.length === 1 && sh + st > headLen) {
    const k = headLen / (sh + st);
    sh *= k;
    st *= k;
  }
  // 方向は縮める前の経路から取る
  const hd = endDirection(path, 'head');
  const td = endDirection(path, 'tail');
  if (sh > 0) {
    // 始点を、経路の内側（外向き hd の逆）へ動かす
    start.x -= hd.x * sh;
    start.y -= hd.y * sh;
    if (firstSeg.t === 'C') {
      firstSeg.x1 -= hd.x * sh;
      firstSeg.y1 -= hd.y * sh;
    }
  }
  if (st > 0) {
    lastSeg.x -= td.x * st;
    lastSeg.y -= td.y * st;
    if (lastSeg.t === 'C') {
      lastSeg.x2 -= td.x * st;
      lastSeg.y2 -= td.y * st;
    }
  }
  return { start, segs };
}

/** 数値を SVG 用の文字列にする（小数 2 桁。有限でない値は 0。`-0` は `0`） */
function fmt(n: number): string {
  return String(Number(fin(n).toFixed(2)));
}

/**
 * 経路を `<path d>` の文字列にする（数値だけから組み立てる。NaN・Infinity は出さない）。閉じた経路は末尾に `Z`。
 *
 * @param originX 座標の原点の x（この値を引く。外接矩形の左上を原点にするときに渡す）
 * @param originY 座標の原点の y
 */
export function pathData(path: ShapePath, originX = 0, originY = 0): string {
  const px = (x: number): string => fmt(x - originX);
  const py = (y: number): string => fmt(y - originY);
  const parts = [`M${px(path.start.x)} ${py(path.start.y)}`];
  for (const s of path.segs) {
    if (s.t === 'L') {
      parts.push(`L${px(s.x)} ${py(s.y)}`);
    } else {
      parts.push(`C${px(s.x1)} ${py(s.y1)} ${px(s.x2)} ${py(s.y2)} ${px(s.x)} ${py(s.y)}`);
    }
  }
  if (path.closed) {
    parts.push('Z');
  }
  return parts.join(' ');
}

/** 部分経路の列を 1 つの `<path d>` にする（図形の path 1 つ。部分経路を空白でつなぐ。空なら空文字） */
export function subpathsData(paths: readonly ShapePath[], originX = 0, originY = 0): string {
  return paths.map((p) => pathData(p, originX, originY)).join(' ');
}

/** 多角形の矢じりを `<path d>` の文字列にする（閉じる図形は `Z`）。座標はそのまま（原点の移動は `renderLine` が済ませる） */
export function arrowPolygonData(a: Extract<ArrowShape, { kind: 'polygon' }>): string {
  const body = a.points.map((p, i) => `${i === 0 ? 'M' : 'L'}${fmt(p.x)} ${fmt(p.y)}`).join(' ');
  return a.closed ? `${body} Z` : body;
}

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
export function renderLine(
  rect: PxRect,
  line: DrawingLine,
  xfrm: Xfrm,
  sheet: SheetBounds,
  preRotated = false,
): LineRender {
  const box = preRotated ? { ...rect } : preRotationRect(rect, xfrm.rot);
  const local = shapePath(line.geom, line.adj, box.w, box.h);
  const path = clampPath(transformPath(local, box, xfrm), sheet);

  const strokeWidth = strokeWidthPx(line.widthEmu);
  const dash = dashArray(line.dash, strokeWidth);

  const pts = pathPoints(path);
  const tipHead = path.start;
  const tipTail = pts[pts.length - 1]!;
  const arrows: ArrowShape[] = [];
  let headShrink = 0;
  let tailShrink = 0;
  let arrowMax = 0;
  if (line.head) {
    arrows.push(arrowShape('head', line.head, tipHead, endDirection(path, 'head'), strokeWidth));
    const sz = arrowSize(line.head, strokeWidth);
    headShrink = sz.length * shrinkRatio(line.head.type);
    arrowMax = Math.max(arrowMax, sz.width, sz.length);
  }
  if (line.tail) {
    arrows.push(arrowShape('tail', line.tail, tipTail, endDirection(path, 'tail'), strokeWidth));
    const sz = arrowSize(line.tail, strokeWidth);
    tailShrink = sz.length * shrinkRatio(line.tail.type);
    arrowMax = Math.max(arrowMax, sz.width, sz.length);
  }
  const drawn = shrinkPath(path, headShrink, tailShrink);

  // 外接矩形：経路の点の外接矩形 ＋ max(太さ/2, 矢じり/2)。矢じりの実際の範囲（先端から後ろへ長さ分）も含める
  const pad = Math.max(strokeWidth / 2, arrowMax / 2);
  let l = Math.min(...pts.map((p) => p.x)) - pad;
  let t = Math.min(...pts.map((p) => p.y)) - pad;
  let r = Math.max(...pts.map((p) => p.x)) + pad;
  let b = Math.max(...pts.map((p) => p.y)) + pad;
  for (const a of arrows) {
    const e = arrowExtent(a);
    l = Math.min(l, e.l - strokeWidth / 2);
    t = Math.min(t, e.t - strokeWidth / 2);
    r = Math.max(r, e.r + strokeWidth / 2);
    b = Math.max(b, e.b + strokeWidth / 2);
  }
  const maxX = Math.max(0, fin(sheet.totalWidth));
  const maxY = Math.max(0, fin(sheet.totalHeight));
  const x0 = Math.max(0, Math.floor(l));
  const y0 = Math.max(0, Math.floor(t));
  const x1 = Math.min(maxX, Math.ceil(r));
  const y1 = Math.min(maxY, Math.ceil(b));
  const bbox: PxRect = { x: x0, y: y0, w: Math.max(0, x1 - x0), h: Math.max(0, y1 - y0) };

  return {
    bbox,
    d: pathData(drawn, bbox.x, bbox.y),
    strokeWidth,
    dash,
    heads: arrows.map((a) => shiftArrow(a, -bbox.x, -bbox.y)),
    path,
  };
}

/** 面の濃淡の種類（path の `fill` が `darken`・`darkenLess`・`lighten`・`lightenLess`。`norm`・`none` は膜なし） */
export type ShapeShade = 'darken' | 'darkenLess' | 'lighten' | 'lightenLess';

const SHADES: ReadonlySet<string> = new Set<ShapeShade>([
  'darken',
  'darkenLess',
  'lighten',
  'lightenLess',
]);

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
export const BACKDROP_PAD = 3;

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
export type ShapeRender = ShapeDrawn | { kind: 'hidden' } | { kind: 'fallback' };

/** 図形が表示する文字を持つか（ランの文字が 1 文字以上ある。改行だけ・空のランは含めない） */
export function shapeHasText(drawing: SheetDrawing): boolean {
  return (
    (drawing.kind === 'shape' || drawing.kind === 'textbox') &&
    !!drawing.text &&
    drawing.text.paragraphs.some((p) => p.runs.some((r) => !r.br && r.text.length > 0))
  );
}

/** 色の不透明度が 0 でない（色が無い・不透明度が有限でないときは見えるとみなす。暫定の色で描かれるため） */
function opaqueColor(c: { alpha: number } | undefined): boolean {
  return c === undefined || !(c.alpha <= 0);
}

/** 塗りがあり、その色が見える（不透明度 0 超。色が無い＝暫定の色は見える）。下地の塗りの条件（ISSUE-96） */
export function filledShown(drawing: SheetDrawing): boolean {
  return drawing.hasFill === true && opaqueColor(drawing.fillColor);
}

/** 見える文字がある（文字があり、不透明度 0 超のランが 1 つ以上ある。`a:noFill` の文字は見えない）。下地の塗りの条件（ISSUE-96） */
export function textShown(drawing: SheetDrawing): boolean {
  return (
    shapeHasText(drawing) &&
    drawing.text!.paragraphs.some((p) =>
      p.runs.some((r) => !r.br && r.text.length > 0 && opaqueColor(r.color)),
    )
  );
}

/** 形（`SheetDrawing.geom`）を評価器の定義と調整値にする。表に無い `prst`・形なし・geom が未読は null */
function geomDefOf(
  geom: ShapeGeom | undefined,
): { def: GeomDef; adjust: ReadonlyMap<string, number> | undefined } | null {
  if (!geom || !geomLib) {
    return null; // geom が未読・読めなかったときも、評価できない形と同じ代替
  }
  if (geom.type === 'preset') {
    const def = geomLib.getPresetGeom(geom.prst);
    return def ? { def, adjust: new Map(geom.adj) } : null;
  }
  return { def: geom.def, adjust: undefined };
}

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
export function renderShape(
  rect: PxRect,
  drawing: SheetDrawing,
  sheet: SheetBounds,
  preRotated = false,
): ShapeRender {
  const resolved = geomDefOf(drawing.geom);
  if (!resolved) {
    return { kind: 'fallback' };
  }
  const box = preRotated ? { ...rect } : preRotationRect(rect, drawing.rot, xfrmExtPx(drawing));
  const ev = geomLib!.evaluateGeom(resolved.def, resolved.adjust, box.w, box.h);
  if (!ev.ok) {
    return { kind: 'fallback' };
  }
  const outline = drawing.hasOutline === true;
  const filled = drawing.hasFill === true;
  if (!outline && !filled && !shapeHasText(drawing)) {
    return { kind: 'hidden' };
  }

  const xfrm: Xfrm = { rot: drawing.rot, flipH: drawing.flipH, flipV: drawing.flipV };
  const strokeWidth = strokeWidthPx(drawing.outline?.widthEmu);
  const dash = dashArray(drawing.outline?.dash, strokeWidth);

  // 描く path（塗らず線も引かない path は捨てる）を、反転 → 回転 → 丸めにかける。
  // 下地（ISSUE-96）だけに要る path（塗りが無く文字のある図形の、形の塗りの path）も同じ写像で持つ（`paint` が false）
  // 下地の塗りは「見える塗り（不透明度 0 超）」か「見える文字」があるときだけ（Excel で何も見えない図形を白くしない）
  const hasText = textShown(drawing);
  const wantBase = filledShown(drawing) || hasText;
  const drawn: {
    fill: boolean;
    stroke: boolean;
    /** 下地で白く塗る（形の塗りの path） */
    base: boolean;
    paint: boolean;
    shade: ShapeShade | null;
    subpaths: ShapePath[];
  }[] = [];
  for (const p of ev.paths) {
    const fillable = p.fill !== 'none';
    const fill = filled && fillable;
    const stroke = outline && p.stroke;
    const base = wantBase && fillable;
    if (!fill && !stroke && !base) {
      continue;
    }
    drawn.push({
      fill,
      stroke,
      base,
      paint: fill || stroke,
      shade: fill && SHADES.has(p.fill) ? (p.fill as ShapeShade) : null,
      subpaths: p.subpaths.map((sp) =>
        clampPath(
          transformPath({ start: sp.start, segs: sp.segs, closed: sp.closed }, box, xfrm),
          sheet,
        ),
      ),
    });
  }
  const painted = drawn.filter((f) => f.paint);

  // 外接矩形 = 経路の外接矩形（太さの半分を足す・シートの範囲に丸める）∪ アンカーの矩形
  let l = rect.x;
  let t = rect.y;
  let r = rect.x + rect.w;
  let b = rect.y + rect.h;
  const pts = painted.flatMap((f) => f.subpaths.flatMap(pathPoints));
  if (pts.length > 0) {
    // 角は miter（miterlimit 4）：鋭角の頂点は線の太さの 2 倍（miterlimit / 2 × 太さ）まで張り出す
    const pad = outline ? (SHAPE_MITER_LIMIT / 2) * strokeWidth : 0;
    const maxX = Math.max(0, fin(sheet.totalWidth));
    const maxY = Math.max(0, fin(sheet.totalHeight));
    let pl = Infinity;
    let pt = Infinity;
    let pr = -Infinity;
    let pb = -Infinity;
    for (const p of pts) {
      pl = Math.min(pl, p.x);
      pt = Math.min(pt, p.y);
      pr = Math.max(pr, p.x);
      pb = Math.max(pb, p.y);
    }
    l = Math.min(l, Math.max(0, Math.floor(pl - pad)));
    t = Math.min(t, Math.max(0, Math.floor(pt - pad)));
    r = Math.max(r, Math.min(maxX, Math.ceil(pr + pad)));
    b = Math.max(b, Math.min(maxY, Math.ceil(pb + pad)));
  }
  const bbox: PxRect = { x: l, y: t, w: r - l, h: b - t };

  // 文字領域：反転は矩形の中で鏡写し
  const tr = ev.textRect;
  const textRect: PxRect = {
    x: drawing.flipH ? box.w - tr.r : tr.l,
    y: drawing.flipV ? box.h - tr.b : tr.t,
    w: tr.r - tr.l,
    h: tr.b - tr.t,
  };

  // 面の `d` を作り、下地の path は同じ `d` を使い回す（塗りが無く文字のある図形の塗りの path だけ追加で作る）
  const ds = new Map(painted.map((f) => [f, subpathsData(f.subpaths, bbox.x, bbox.y)] as const));
  const backdrop: ShapeBackdropPath[] = [];
  for (const f of drawn) {
    if (f.base || f.stroke) {
      backdrop.push({
        d: ds.get(f) ?? subpathsData(f.subpaths, bbox.x, bbox.y),
        fill: f.base,
        stroke: f.stroke,
      });
    }
  }
  if (hasText && !backdrop.some((b) => b.fill)) {
    // 塗りの path が無い形：回転前の矩形を、図形と同じ反転・回転にかけて塗る
    const rectPath: ShapePath = {
      start: { x: 0, y: 0 },
      segs: [
        { t: 'L', x: box.w, y: 0 },
        { t: 'L', x: box.w, y: box.h },
        { t: 'L', x: 0, y: box.h },
      ],
      closed: true,
    };
    const d = pathData(clampPath(transformPath(rectPath, box, xfrm), sheet), bbox.x, bbox.y);
    backdrop.unshift({ d, fill: true, stroke: false });
  }

  return {
    kind: 'drawn',
    rect: bbox,
    box,
    textRect,
    faces: painted.map((f) => ({
      d: ds.get(f)!,
      fill: f.fill,
      stroke: f.stroke,
      shade: f.shade,
      subpaths: f.subpaths,
    })),
    backdrop,
    strokeWidth,
    dash,
  };
}
