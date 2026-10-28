// 正: InterHub docs/plan/adr/0026-xlsx-image-overlay.md（§決定 8）
/**
 * `.xlsx` の図（画像・図形など）をセルの上に重ねるための純ロジック（DOM 非依存・ISSUE-86）。
 * - アンカー（`lib/xlsx` の `DrawingAnchor`）を、A1 の左上を原点とする px の矩形にする（geometry から出す）。
 * - 図の最下行・最右列（表をそこまで広げる）・スクロール枠の可視範囲との交差判定を出す。
 * 位置は列幅・行高と同じ geometry から計算するため、表のセルと 1px もずれない。
 * グループの中の図（ISSUE-89）は、最も外側のグループのアンカーの矩形へ、グループの子の座標系（`chOff`・`chExt`）から線形に写して矩形にする。
 */
import type {
  AnchorPoint,
  DrawingAnchor,
  SheetDrawing,
  Geometry,
  DrawingGroup,
  DrawingLine,
  GroupRect,
} from '../core/index';
import { MAX_GRID_COLS, MAX_GRID_ROWS } from './grid';
import { renderLine, renderShape, type LineRender, type ShapeDrawn } from './shapes';

/** 1px あたりの EMU（Office の英語メートル単位。px＝EMU / 9525） */
export const EMU_PER_PX = 9525;

/** 座標・大きさの上限（px）。ブラウザのレイアウトの上限（約 33,554,431px）を超えないようにする */
const MAX_PX = 33_000_000;

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

/** `SheetDrawing.line` が無い線（壊れたデータ）に使う既定：細い直線・矢じりなし */
const DEFAULT_LINE: DrawingLine = {
  geom: 'line',
  adj: [50000, 50000, 50000],
  widthEmu: 0,
  dash: 'solid',
  head: null,
  tail: null,
};

/** EMU を px（整数）にする。有限でない値は 0、負は 0、上限に丸める */
function emuToPx(emu: number): number {
  if (!Number.isFinite(emu) || emu <= 0) {
    return 0;
  }
  return Math.min(MAX_PX, Math.round(emu / EMU_PER_PX));
}

/** 0 始まりの番号を [0, max] の整数に丸める（有限でない値は 0） */
function clampIndex(n: number, max: number): number {
  if (!Number.isFinite(n)) {
    return 0;
  }
  return Math.min(max, Math.max(0, Math.trunc(n)));
}

/**
 * アンカーの点の x（px）：その列の左端＋ずれ。ずれが列の幅を超える時は列の右端で止める
 * （列幅の僅差で隣のセルへ飛び出さない。ADR-0026 §決定 8）。
 * geometry の番号は 1 始まり、アンカーの `col` は 0 始まり。
 */
export function anchorPointX(geo: AnchorGeometry, p: AnchorPoint): number {
  const c = clampIndex(p.col, MAX_GRID_COLS - 1) + 1;
  return geo.colLeft(c) + Math.min(emuToPx(p.colOff), geo.colWidth(c));
}

/** アンカーの点の y（px）：その行の上端＋ずれ（行の高さを超える時は行の下端で止める） */
export function anchorPointY(geo: AnchorGeometry, p: AnchorPoint): number {
  const r = clampIndex(p.row, MAX_GRID_ROWS - 1) + 1;
  return geo.rowTop(r) + Math.min(emuToPx(p.rowOff), geo.rowHeight(r));
}

/**
 * アンカーを A1 の左上基準の矩形（px）にする。幅・高さは 0 や負の値のままで、大きさの判定はしない
 * （線は幅か高さが 0 でも描くため。判定は `anchorRect` と `layoutDrawings` が行う）。未知の種類は null。
 */
function rawAnchorRect(geo: AnchorGeometry, anchor: DrawingAnchor): PxRect | null {
  let rect: PxRect;
  switch (anchor.type) {
    case 'twoCell': {
      const x = anchorPointX(geo, anchor.from);
      const y = anchorPointY(geo, anchor.from);
      rect = { x, y, w: anchorPointX(geo, anchor.to) - x, h: anchorPointY(geo, anchor.to) - y };
      break;
    }
    case 'oneCell':
      rect = {
        x: anchorPointX(geo, anchor.from),
        y: anchorPointY(geo, anchor.from),
        w: emuToPx(anchor.cx),
        h: emuToPx(anchor.cy),
      };
      break;
    case 'absolute':
      rect = {
        x: emuToPx(anchor.x),
        y: emuToPx(anchor.y),
        w: emuToPx(anchor.cx),
        h: emuToPx(anchor.cy),
      };
      break;
    default:
      return null;
  }
  return rect;
}

/** 小数のまま持ち回る矩形（グループの写像の途中で丸めない） */
interface RectF {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** 子の座標系（`ch`）の矩形 `r` を、親の矩形 `parent`（px）へ線形に写す。`ch` の大きさは 1 以上 */
function mapChild(parent: RectF, ch: GroupRect, r: GroupRect): RectF {
  const kx = parent.w / ch.cx;
  const ky = parent.h / ch.cy;
  return {
    x: parent.x + (r.x - ch.x) * kx,
    y: parent.y + (r.y - ch.y) * ky,
    w: r.cx * kx,
    h: r.cy * ky,
  };
}

/**
 * グループの中の図の矩形（px・整数）を出す（ADR-0027 §決定 5）。
 * `anchor` は最も外側のグループのアンカーの矩形。グループの子の座標系（`chOff`・`chExt`）から親の矩形へ、外側から順に線形に写す：
 * `x = 親.x + (子.off.x − chOff.x) × 親.w / chExt.cx`（y・幅・高さも同じ）。入れ子のグループは、内側のグループ自身の `off`・`ext` を外側の座標系から写して、その矩形を次の親にする。
 * グループ自身の回転・反転は無視する。図の `xfrm` は回転前の矩形そのものなので、縦横の入れ替えはしない。
 * 端は四捨五入して幅・高さは端の差で出す。矩形は A1 の左・上へはみ出しても切らない（切ると線の傾き・文字の折り返しが変わる。はみ出しは
 * 線の `clampPath` とレイヤー（`overflow: hidden`）が隠す）。位置は ±MAX_PX、大きさは MAX_PX までで止める。座標系が壊れている・有限でない値になるときは null。
 */
export function groupChildRect(anchor: PxRect, group: DrawingGroup): PxRect | null {
  const { levels, xfrm } = group;
  if (levels.length === 0) {
    return null;
  }
  let cur: RectF = {
    x: anchor.x,
    y: anchor.y,
    w: Math.max(0, anchor.w),
    h: Math.max(0, anchor.h),
  };
  for (let i = 0; i < levels.length; i++) {
    const lv = levels[i]!;
    if (!(lv.ch.cx > 0 && lv.ch.cy > 0)) {
      return null;
    }
    // 次の段（内側のグループ）の位置、最後は図自身の xfrm
    const target = i + 1 < levels.length ? levels[i + 1]!.own : xfrm;
    if (target === null) {
      return null;
    }
    cur = mapChild(cur, lv.ch, target);
  }
  const edge = (v: number): number => Math.min(MAX_PX, Math.max(-MAX_PX, Math.round(v)));
  const left = edge(cur.x);
  const top = edge(cur.y);
  const right = edge(cur.x + cur.w);
  const bottom = edge(cur.y + cur.h);
  if (![left, top, right, bottom].every(Number.isFinite)) {
    return null;
  }
  return {
    x: left,
    y: top,
    w: Math.min(MAX_PX, Math.max(0, right - left)),
    h: Math.min(MAX_PX, Math.max(0, bottom - top)),
  };
}

/**
 * 図の矩形（px）。トップレベルの図はアンカーの矩形、グループの中の図は最も外側のグループのアンカーから写した矩形
 * （幅・高さは 0 や負のままで、大きさの判定はしない。判定は `layoutDrawings` が行う）。
 */
export function placeDrawing(geo: AnchorGeometry, drawing: SheetDrawing): PxRect | null {
  const raw = rawAnchorRect(geo, drawing.anchor);
  return raw === null || drawing.group === undefined ? raw : groupChildRect(raw, drawing.group);
}

/**
 * アンカーを A1 の左上基準の矩形（px）にする。大きさが 0 以下の図（非表示の行・列だけにかかる・逆向きの twoCell・
 * 大きさ 0 の ext）は null（出さない）。
 * - twoCell：`from`〜`to` の差。
 * - oneCell：`from` を起点に `cx`×`cy`。
 * - absolute：シートの左上からの `x`・`y` に `cx`×`cy`。
 */
export function anchorRect(geo: AnchorGeometry, anchor: DrawingAnchor): PxRect | null {
  const rect = rawAnchorRect(geo, anchor);
  return rect && rect.w > 0 && rect.h > 0 ? rect : null;
}

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
export function layoutDrawings(
  geo: DrawingGeometry,
  drawings: readonly SheetDrawing[],
): DrawingBox[] {
  const boxes: DrawingBox[] = [];
  drawings.forEach((d, index) => {
    if (d.kind === 'line') {
      const raw = placeDrawing(geo, d);
      if (!raw || (raw.w <= 0 && raw.h <= 0)) {
        return;
      }
      const anchor: PxRect = { x: raw.x, y: raw.y, w: Math.max(0, raw.w), h: Math.max(0, raw.h) };
      const line = renderLine(
        anchor,
        d.line ?? DEFAULT_LINE,
        { rot: d.rot, flipH: d.flipH, flipV: d.flipV },
        { totalWidth: geo.totalWidth, totalHeight: geo.totalHeight },
        d.group !== undefined,
      );
      if (line.bbox.w > 0 && line.bbox.h > 0) {
        boxes.push({ index, rect: line.bbox, line });
      }
      return;
    }
    const rect = placeDrawing(geo, d);
    if (rect && rect.w > 0 && rect.h > 0) {
      if ((d.kind === 'shape' || d.kind === 'textbox') && d.geom !== undefined) {
        const shape = renderShape(
          rect,
          d,
          { totalWidth: geo.totalWidth, totalHeight: geo.totalHeight },
          d.group !== undefined,
        );
        if (shape.kind === 'hidden') {
          return;
        }
        if (shape.kind === 'drawn') {
          boxes.push({ index, rect: shape.rect, shape });
          return;
        }
      }
      boxes.push({ index, rect });
    }
  });
  return boxes;
}

/**
 * 図を収めるのに必要な表の行数・列数（最下行・最右列。1 始まり）。図が無ければ 0。
 * `gridExtent` の `minRows`・`minCols` に渡して、図のある範囲まで表を広げる。
 *
 * @param geo 行数・列数の上限を Excel の最大にした geometry（`rowAt`・`colAt` が図の端を含む行・列へ届くように）
 */
export function drawingExtent(
  geo: Pick<Geometry, 'rowAt' | 'colAt'>,
  boxes: readonly DrawingBox[],
): { rows: number; cols: number } {
  let rows = 0;
  let cols = 0;
  for (const { rect } of boxes) {
    // 端の 1px 手前（図に含まれる最後の px）を含む行・列
    rows = Math.max(rows, geo.rowAt(Math.max(0, rect.y + rect.h - 1)));
    cols = Math.max(cols, geo.colAt(Math.max(0, rect.x + rect.w - 1)));
  }
  return { rows, cols };
}

/**
 * スクロール枠の可視範囲（A1 の左上基準・px）に、上下左右 1 画面分の余白を足した範囲。
 * 表の原点（A1 の左上）はスクロール枠の内容の (`originX`, `originY`)＝行番号の列の幅・見出し行の高さ。
 * 枠の大きさが 0（まだ DOM に付いていない）なら幅・高さ 0 の範囲を返す（何も交差しない）。
 */
export function viewWindow(
  scrollLeft: number,
  scrollTop: number,
  clientWidth: number,
  clientHeight: number,
  originX: number,
  originY: number,
): PxRect {
  const w = Math.max(0, clientWidth);
  const h = Math.max(0, clientHeight);
  return {
    x: scrollLeft - originX - w,
    y: scrollTop - originY - h,
    w: w === 0 ? 0 : w * 3,
    h: h === 0 ? 0 : h * 3,
  };
}

/** 2 つの矩形が交わるか（接するだけは交わらない。幅・高さ 0 の矩形は何とも交わらない） */
export function rectsIntersect(a: PxRect, b: PxRect): boolean {
  return (
    a.w > 0 &&
    a.h > 0 &&
    b.w > 0 &&
    b.h > 0 &&
    a.x < b.x + b.w &&
    b.x < a.x + a.w &&
    a.y < b.y + b.h &&
    b.y < a.y + a.h
  );
}

/** 範囲 `win` と交わる図だけを、元の順（重ね順）で返す。図は最大 1,000 個なので毎回の線形走査でよい */
export function pickVisible(boxes: readonly DrawingBox[], win: PxRect): DrawingBox[] {
  return boxes.filter((b) => rectsIntersect(b.rect, win));
}
