// 正: docs/adr/0007-moobius-viewer-api.md §5
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
export type Zoom = number | { fit: 'width'; min?: number; max?: number };

/** 倍率の下限・上限 */
export const ZOOM_MIN = 0.5;
export const ZOOM_MAX = 2;

/** 解決した倍率の指定（数は範囲に丸めた値。fit は範囲に丸めた下限・上限） */
export type ZoomSpec = { fixed: number } | { fit: { min: number; max: number } };

const clampZoom = (v: number): number => Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, v));

/**
 * 倍率の指定を確かめて解決する。数は 0.5〜2.0 に丸める。fit の `min`・`max` は数でなければ既定（0.5・2.0）、
 * 逆転していれば入れ替える。
 * @throws RangeError 数でない値（`NaN`・`Infinity` を含む）・`{ fit: 'width' }` でないもの
 */
export function resolveZoom(zoom: Zoom): ZoomSpec {
  if (typeof zoom === 'number') {
    if (!Number.isFinite(zoom)) throw new RangeError('倍率が数でない');
    return { fixed: clampZoom(zoom) };
  }
  if (typeof zoom !== 'object' || zoom === null || zoom.fit !== 'width') {
    throw new RangeError('倍率の指定が正しくない');
  }
  const num = (v: unknown, def: number): number =>
    typeof v === 'number' && Number.isFinite(v) ? clampZoom(v) : def;
  const a = num(zoom.min, ZOOM_MIN);
  const b = num(zoom.max, ZOOM_MAX);
  return { fit: { min: Math.min(a, b), max: Math.max(a, b) } };
}

/**
 * 寸法を倍率で包む。境目（行の上端・列の左端）を整数に丸め、高さ・幅は境目の差にする。
 * `rowAt`・`colAt` は倍率で割って元の寸法で引いた後、前後の行（列）を丸めた境目と比べて合わせる。
 * 倍率 1 なら元の寸法をそのまま返す（今までと同じ道を通る）。
 */
export function zoomGeometry(g: Geometry, z: number): Geometry {
  if (z === 1) return g;
  const top = (r: number): number => Math.round(g.rowTop(r) * z);
  const left = (c: number): number => Math.round(g.colLeft(c) * z);
  const lastRow = g.rowAt(Number.MAX_SAFE_INTEGER);
  const lastCol = g.colAt(Number.MAX_SAFE_INTEGER);
  /** 位置 v を含む番号（start は元の寸法で引いた見当。境目 edge で前後を合わせる） */
  const settle = (v: number, start: number, last: number, edge: (i: number) => number): number => {
    if (last < 1) return 0;
    let i = Math.min(last, Math.max(1, start));
    while (i > 1 && edge(i) > v) i--;
    while (i < last && edge(i + 1) <= v) i++;
    return i;
  };
  return {
    rowTop: top,
    rowHeight: (r) => (r < 1 ? 0 : top(r + 1) - top(r)),
    colLeft: left,
    colWidth: (c) => (c < 1 ? 0 : left(c + 1) - left(c)),
    rowAt: (y) => (Number.isNaN(y) ? g.rowAt(y) : settle(y, g.rowAt(y / z), lastRow, top)),
    colAt: (x) => (Number.isNaN(x) ? g.colAt(x) : settle(x, g.colAt(x / z), lastCol, left)),
    totalWidth: Math.round(g.totalWidth * z),
    totalHeight: Math.round(g.totalHeight * z),
  };
}

/** 図の出し入れの判定の範囲（倍率を掛けた座標）を、図の層の中の座標（100%）に戻す */
export function unzoomRect(rect: PxRect, z: number): PxRect {
  return z === 1 ? rect : { x: rect.x / z, y: rect.y / z, w: rect.w / z, h: rect.h / z };
}

/**
 * 幅に合わせるときの中身の幅（100% の px）：データのある列の右端と、図の右端の大きいほう。
 * @param dataWidth データのある列（1〜`colCount`）の幅の合計
 */
export function contentWidth(dataWidth: number, boxes: readonly DrawingBox[]): number {
  let w = Math.max(0, dataWidth);
  for (const b of boxes) w = Math.max(w, b.rect.x + b.rect.w);
  return w;
}

/**
 * 幅に合わせる倍率。合わせ直さなくてよい（中身・枠の幅が無い、今の倍率との差が 1% 未満）ときは null。
 * @param available セルに使える表示枠の幅（px。行番号の列とスクロールバーを除く）
 * @param content 中身の幅（100% の px）
 * @param current 今の倍率
 */
export function fitZoom(
  available: number,
  content: number,
  range: { min: number; max: number },
  current: number,
): number | null {
  if (!(available > 0) || !(content > 0)) return null;
  const z = Math.min(range.max, Math.max(range.min, available / content));
  return Math.abs(z - current) < current * 0.01 ? null : z;
}

/** 表示枠の左上にあるセルと、そのセルの中の位置の割合（倍率を替えても同じセルを左上に残すため） */
export interface ScrollAnchor {
  r: number;
  fy: number;
  c: number;
  fx: number;
}

/** 今のスクロール位置（A1 の左上からの px）の左上のセル */
export function scrollAnchor(g: Geometry, scrollTop: number, scrollLeft: number): ScrollAnchor {
  const r = g.rowAt(scrollTop);
  const c = g.colAt(scrollLeft);
  const h = g.rowHeight(r);
  const w = g.colWidth(c);
  return {
    r,
    fy: h > 0 ? Math.min(1, Math.max(0, (scrollTop - g.rowTop(r)) / h)) : 0,
    c,
    fx: w > 0 ? Math.min(1, Math.max(0, (scrollLeft - g.colLeft(c)) / w)) : 0,
  };
}

/** 左上のセルを、新しい寸法でのスクロール位置にする */
export function anchorScroll(g: Geometry, a: ScrollAnchor): { top: number; left: number } {
  return {
    top: Math.round(g.rowTop(a.r) + a.fy * g.rowHeight(a.r)),
    left: Math.round(g.colLeft(a.c) + a.fx * g.colWidth(a.c)),
  };
}

/** スクロールバーの幅（px）。1 回だけ測って覚える（重なるスクロールバーの環境では 0） */
let barWidth: number | null = null;
export function scrollbarWidth(): number {
  if (barWidth !== null) return barWidth;
  const host = typeof document === 'undefined' ? null : document.body;
  if (!host) return 0;
  const probe = document.createElement('div');
  probe.style.cssText = 'position:absolute;top:-9999px;width:100px;height:100px;overflow:scroll';
  host.append(probe);
  barWidth = Math.max(0, probe.offsetWidth - probe.clientWidth);
  probe.remove();
  return barWidth;
}
