/// <reference types="node" />
import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { openXlsx, type SheetDrawing } from '../../src/core/index';
import type { LineEnd } from '../../src/core/types';
import type { PxRect } from '../../src/render/drawings';
import {
  arrowDimension,
  arrowPolygonData,
  arrowShape,
  clampPath,
  dashArray,
  endDirection,
  pathData,
  pathPoints,
  preRotationRect,
  renderLine,
  shapePath,
  shrinkPath,
  strokeWidthPx,
  transformPath,
  type ArrowShape,
  type Pt,
  type ShapePath,
} from '../../src/render/shapes';

const EMU = 9525;
const BIG = { totalWidth: 1_000_000, totalHeight: 1_000_000 };
const DEG = 60000;

/** 経路の始点・終点 */
function ends(p: ShapePath): { s: Pt; e: Pt } {
  const pts = pathPoints(p);
  return { s: pts[0]!, e: pts[pts.length - 1]! };
}

/** 経路の主要点（直線・折れ線のみ。制御点は含めない）を [x, y][] にする */
function mainPts(p: ShapePath): [number, number][] {
  return [[p.start.x, p.start.y], ...p.segs.map((s): [number, number] => [s.x, s.y])];
}

/** `M... L... C...` の文字列から数値だけを順に取り出す */
function nums(d: string): number[] {
  return (d.match(/-?\d+(?:\.\d+)?/g) ?? []).map(Number);
}

/** 回転前の矩形 box（w×h）の回転後の外接矩形（回転 90・270 は縦横入れ替え） */
function anchorOf(box: PxRect, rotDeg: number): PxRect {
  const swap = rotDeg === 90 || rotDeg === 270;
  const w = swap ? box.h : box.w;
  const h = swap ? box.w : box.h;
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  return { x: cx - w / 2, y: cy - h / 2, w, h };
}

/** 点を矩形 box の中心で時計回りに 0・90・180・270 度回す（画面座標・y 下向き） */
function rot90s(p: Pt, box: PxRect, rotDeg: number): Pt {
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  const dx = p.x - cx;
  const dy = p.y - cy;
  switch (((rotDeg % 360) + 360) % 360) {
    case 90:
      return { x: cx - dy, y: cy + dx };
    case 180:
      return { x: cx - dx, y: cy - dy };
    case 270:
      return { x: cx + dy, y: cy - dx };
    default:
      return p;
  }
}

describe('TD-88-1 回転前の矩形（preRotationRect）', () => {
  const rect: PxRect = { x: 100, y: 50, w: 40, h: 200 };
  // 中心は (120, 150)
  const cases: [number, boolean][] = [
    [0, false],
    [44, false],
    [45, true],
    [90, true],
    [134, true],
    [135, false],
    [180, false],
    [224, false],
    [225, true],
    [270, true],
    [314, true],
    [315, false],
    [359, false],
    [360, false],
    [405, true],
    [-90, true],
  ];
  for (const [deg, swap] of cases) {
    it(`${deg}° は ${swap ? '縦横を入れ替える' : '入れ替えない'}`, () => {
      const r = preRotationRect(rect, deg * DEG);
      if (swap) {
        expect(r).toEqual({ x: 20, y: 130, w: 200, h: 40 });
      } else {
        expect(r).toEqual(rect);
      }
      // 中心は常にアンカーと同じ
      expect(r.x + r.w / 2).toBeCloseTo(120);
      expect(r.y + r.h / 2).toBeCloseTo(150);
    });
  }

  it('境界は 1/60000 度単位：44.99…° は入れ替えず、45° ちょうどは入れ替える', () => {
    expect(preRotationRect(rect, 45 * DEG - 1)).toEqual(rect);
    expect(preRotationRect(rect, 45 * DEG).w).toBe(200);
    expect(preRotationRect(rect, 135 * DEG - 1).w).toBe(200);
    expect(preRotationRect(rect, 135 * DEG)).toEqual(rect);
  });

  it('ext（回転前の大きさ）があれば、境界（45° ちょうど）でも Excel のアンカーに合わせて入れ替えを決める（ISSUE-93）', () => {
    // Excel は 45° ちょうどの図形のアンカーを入れ替えない（回転前の 120×53 のまま）。値だけの規則だと入れ替えてしまう
    const anchor: PxRect = { x: 100, y: 50, w: 120, h: 53 };
    const ext = { w: 120, h: 53 };
    expect(preRotationRect(anchor, 45 * DEG).w).toBe(53);
    expect(preRotationRect(anchor, 45 * DEG, ext)).toEqual(anchor);
    // 入れ替わったアンカー（53×120）は、ext が 120×53 なら入れ替えて戻す（45° 以上 135° 未満の通常の場合）
    const swapped: PxRect = { x: 133.5, y: 16.5, w: 53, h: 120 };
    const back = preRotationRect(swapped, 90 * DEG, ext);
    expect(back).toEqual({ x: 100, y: 50, w: 120, h: 53 });
    // 135° ちょうど・225°・315° も、アンカーに合う向きを選ぶ
    expect(preRotationRect(anchor, 135 * DEG, ext)).toEqual(anchor);
    expect(preRotationRect(swapped, 135 * DEG, ext).w).toBe(120);
    expect(preRotationRect(anchor, 315 * DEG, ext)).toEqual(anchor);
    expect(preRotationRect(swapped, 225 * DEG, ext).w).toBe(120);
    // 回転が 0 に近くても、アンカーが ext の縦横を入れ替えた大きさなら入れ替える
    expect(preRotationRect(swapped, 10 * DEG, ext).w).toBe(120);
    // 中心は常にアンカーと同じ
    const r = preRotationRect(swapped, 90 * DEG, ext);
    expect(r.x + r.w / 2).toBeCloseTo(swapped.x + swapped.w / 2);
    expect(r.y + r.h / 2).toBeCloseTo(swapped.y + swapped.h / 2);
  });

  it('ext が決め手にならないとき（正方形・差が 1px 以内・0 以下・無い）は、回転の値で決める', () => {
    const sq: PxRect = { x: 0, y: 0, w: 80, h: 80 };
    expect(preRotationRect(sq, 90 * DEG, { w: 80, h: 80 })).toEqual(sq);
    // 差が 1px 以内（keep 0.6・swapped 1.4）：アンカー 100×99・ext 100×99.6 → 規則どおり（90° は入れ替え、30° は入れ替えない）
    const near: PxRect = { x: 0, y: 0, w: 100, h: 99 };
    expect(preRotationRect(near, 30 * DEG, { w: 100, h: 99.6 })).toEqual(near);
    expect(preRotationRect(near, 90 * DEG, { w: 100, h: 99.6 }).w).toBe(99);
    // ext が 0・負は無視
    expect(preRotationRect(rect, 90 * DEG, { w: 0, h: 10 }).w).toBe(200);
    expect(preRotationRect(rect, 0, { w: -5, h: 10 })).toEqual(rect);
  });

  it('rot が無い・NaN は 0 として扱う。元の矩形を壊さない', () => {
    expect(preRotationRect(rect, undefined)).toEqual(rect);
    expect(preRotationRect(rect, Number.NaN)).toEqual(rect);
    const r = preRotationRect(rect, 0);
    r.x = 999;
    expect(rect.x).toBe(100);
  });
});

describe('TD-88-2 形の経路（shapePath）', () => {
  it('line・straightConnector1・未知の geom は (0,0)→(w,h) の直線', () => {
    for (const g of ['line', 'straightConnector1', 'custGeom', '', 'unknownThing']) {
      expect(mainPts(shapePath(g, [], 80, 30))).toEqual([
        [0, 0],
        [80, 30],
      ]);
    }
  });

  it('bentConnector2：(0,0)→(w,0)→(w,h)', () => {
    expect(mainPts(shapePath('bentConnector2', [], 80, 30))).toEqual([
      [0, 0],
      [80, 0],
      [80, 30],
    ]);
  });

  it('bentConnector3：x1 = w·adj1（既定 50%）', () => {
    expect(mainPts(shapePath('bentConnector3', [], 80, 30))).toEqual([
      [0, 0],
      [40, 0],
      [40, 30],
      [80, 30],
    ]);
    expect(mainPts(shapePath('bentConnector3', [30000], 80, 30))).toEqual([
      [0, 0],
      [24, 0],
      [24, 30],
      [80, 30],
    ]);
    // 負・100% 超えは矩形の外へ張り出す
    expect(shapePath('bentConnector3', [-40000], 80, 30).segs[0]).toEqual({ t: 'L', x: -32, y: 0 });
    expect(shapePath('bentConnector3', [160000], 80, 30).segs[0]).toEqual({ t: 'L', x: 128, y: 0 });
  });

  it('adj が undefined の要素・非有限は 50000 として扱う', () => {
    const p = mainPts(shapePath('bentConnector3', [undefined, 1, 2], 80, 30));
    expect(p[1]).toEqual([40, 0]);
    expect(mainPts(shapePath('bentConnector3', [Number.NaN], 80, 30))[1]).toEqual([40, 0]);
    expect(mainPts(shapePath('bentConnector3', undefined, 80, 30))[1]).toEqual([40, 0]);
  });

  it('bentConnector4：x1 = w·adj1、y2 = h·adj2', () => {
    expect(mainPts(shapePath('bentConnector4', [-18947, 65000], 100, 200))).toEqual([
      [0, 0],
      [-18.947, 0],
      [-18.947, 130],
      [100, 130],
      [100, 200],
    ]);
  });

  it('bentConnector5：x1・y2・x3 の 5 区間', () => {
    expect(mainPts(shapePath('bentConnector5', [-16364, 50000, 116364], 100, 200))).toEqual([
      [0, 0],
      [-16.364, 0],
      [-16.364, 100],
      [116.364, 100],
      [116.364, 200],
      [100, 200],
    ]);
  });

  it('curvedConnector4・5 は同じ調整値の bentConnector4・5 の折れ線で代替', () => {
    expect(shapePath('curvedConnector4', [20000, 70000], 100, 50)).toEqual(
      shapePath('bentConnector4', [20000, 70000], 100, 50),
    );
    expect(shapePath('curvedConnector5', [10000, 60000, 90000], 100, 50)).toEqual(
      shapePath('bentConnector5', [10000, 60000, 90000], 100, 50),
    );
  });

  it('curvedConnector2：3 次ベジェ (0,0) C (w/2,0) (w,h/2) (w,h)', () => {
    expect(shapePath('curvedConnector2', [], 80, 30)).toEqual({
      start: { x: 0, y: 0 },
      segs: [{ t: 'C', x1: 40, y1: 0, x2: 80, y2: 15, x: 80, y: 30 }],
    });
  });

  it('curvedConnector3：x2 = w·adj1 の 2 つのベジェ', () => {
    expect(shapePath('curvedConnector3', [20000], 100, 40)).toEqual({
      start: { x: 0, y: 0 },
      segs: [
        { t: 'C', x1: 10, y1: 0, x2: 20, y2: 10, x: 20, y: 20 },
        { t: 'C', x1: 20, y1: 30, x2: 60, y2: 40, x: 100, y: 40 },
      ],
    });
  });

  // 12 通りの prst × 反転 4 通り × 回転 0/90/180/270 の端点（始点＝(0,0) の反転・回転、終点＝(w,h) の反転・回転）
  const geoms = [
    'line',
    'straightConnector1',
    'bentConnector2',
    'bentConnector3',
    'bentConnector4',
    'bentConnector5',
    'curvedConnector2',
    'curvedConnector3',
    'curvedConnector4',
    'curvedConnector5',
    'custGeom',
  ];
  for (const geom of geoms) {
    for (const [flipH, flipV] of [
      [false, false],
      [true, false],
      [false, true],
      [true, true],
    ] as const) {
      for (const rotDeg of [0, 90, 180, 270]) {
        it(`${geom} flipH=${flipH} flipV=${flipV} rot=${rotDeg}：端点`, () => {
          const box: PxRect = { x: 200, y: 300, w: 100, h: 40 };
          const local = shapePath(geom, [30000, 60000, 80000], box.w, box.h);
          const moved = transformPath(local, box, { rot: rotDeg * DEG, flipH, flipV });
          const sx = flipH ? box.x + box.w : box.x;
          const sy = flipV ? box.y + box.h : box.y;
          const ex = flipH ? box.x : box.x + box.w;
          const ey = flipV ? box.y : box.y + box.h;
          const s = rot90s({ x: sx, y: sy }, box, rotDeg);
          const e = rot90s({ x: ex, y: ey }, box, rotDeg);
          const got = ends(moved);
          expect(got.s.x).toBeCloseTo(s.x, 6);
          expect(got.s.y).toBeCloseTo(s.y, 6);
          expect(got.e.x).toBeCloseTo(e.x, 6);
          expect(got.e.y).toBeCloseTo(e.y, 6);
        });
      }
    }
  }

  it('反転は矩形の中で行う：flipH で x→w-x、flipV で y→h-y。ベジェの制御点も同じ変換', () => {
    const box: PxRect = { x: 10, y: 20, w: 100, h: 40 };
    const local = shapePath('curvedConnector2', [], box.w, box.h);
    const moved = transformPath(local, box, { flipH: true, flipV: true });
    expect(moved.start).toEqual({ x: 110, y: 60 });
    expect(moved.segs[0]).toEqual({ t: 'C', x1: 60, y1: 60, x2: 10, y2: 40, x: 10, y: 20 });
  });

  it('回転は時計回り（画面座標）：90° で (w,0) の右上の点は右下へ回る', () => {
    const box: PxRect = { x: 0, y: 0, w: 100, h: 100 };
    const p = transformPath({ start: { x: 100, y: 0 }, segs: [] }, box, { rot: 90 * DEG });
    expect(p.start.x).toBeCloseTo(100);
    expect(p.start.y).toBeCloseTo(100);
  });

  it('45° 回転：(0,0) は中心から見て上へ（画面座標）', () => {
    const box: PxRect = { x: 0, y: 0, w: 100, h: 100 };
    const p = transformPath({ start: { x: 0, y: 0 }, segs: [] }, box, { rot: 45 * DEG });
    expect(p.start.x).toBeCloseTo(50);
    expect(p.start.y).toBeCloseTo(50 - Math.SQRT2 * 50);
  });
});

describe('TD-88-3 範囲外の調整値と経路の点の丸め', () => {
  it('U 字（幅 12,700 EMU・adj1=1,800,000）は矩形の外へ約 22.7px 張り出し、丸めで縮まない', () => {
    const w = 12700 / EMU;
    const local = shapePath('bentConnector3', [1_800_000], w, 60);
    expect(local.segs[0]!.x).toBeCloseTo(24, 6); // x1 = 1.3333 × 18
    expect(local.segs[0]!.x - w).toBeCloseTo(22.667, 2);
    const box: PxRect = { x: 400, y: 100, w, h: 60 };
    const moved = clampPath(transformPath(local, box, {}), BIG);
    expect(Math.max(...pathPoints(moved).map((p) => p.x))).toBeCloseTo(400 + 24, 6);
    // renderLine 経由でも外接矩形は張り出しを含む
    const r = renderLine(
      { x: box.x, y: box.y, w, h: 60 },
      {
        geom: 'bentConnector3',
        adj: [1_800_000, 50000, 50000],
        widthEmu: 9525,
        dash: 'solid',
        head: null,
        tail: null,
      },
      {},
      BIG,
    );
    expect(r.bbox.x + r.bbox.w).toBeGreaterThanOrEqual(424);
  });

  it('左へ張り出す U 字（rot 180・flipV）も張り出しを保つ', () => {
    const w = 12700 / EMU;
    const box: PxRect = { x: 400, y: 100, w, h: 60 };
    const moved = clampPath(
      transformPath(shapePath('bentConnector3', [1_800_000], w, 60), box, {
        rot: 180 * DEG,
        flipV: true,
      }),
      BIG,
    );
    expect(Math.min(...pathPoints(moved).map((p) => p.x))).toBeCloseTo(400 + w - 24, 6);
  });

  for (const adj of [2147483647, -2147483647]) {
    for (const geom of [
      'bentConnector3',
      'bentConnector4',
      'bentConnector5',
      'curvedConnector3',
    ] as const) {
      it(`${geom} adj=${adj}：全点・外接矩形がシートの範囲に収まる`, () => {
        const sheet = { totalWidth: 800, totalHeight: 600 };
        const rect: PxRect = { x: 100, y: 100, w: 50, h: 30 };
        const r = renderLine(
          rect,
          {
            geom,
            adj: [adj, adj, adj],
            widthEmu: 9525,
            dash: 'solid',
            head: { type: 'triangle', w: 'lg', len: 'lg' },
            tail: null,
          },
          { rot: 90 * DEG, flipH: true },
          sheet,
        );
        for (const p of pathPoints(r.path)) {
          expect(p.x).toBeGreaterThanOrEqual(0);
          expect(p.x).toBeLessThanOrEqual(800);
          expect(p.y).toBeGreaterThanOrEqual(0);
          expect(p.y).toBeLessThanOrEqual(600);
        }
        expect(r.bbox.x).toBeGreaterThanOrEqual(0);
        expect(r.bbox.y).toBeGreaterThanOrEqual(0);
        expect(r.bbox.x + r.bbox.w).toBeLessThanOrEqual(800);
        expect(r.bbox.y + r.bbox.h).toBeLessThanOrEqual(600);
        expect(r.d).not.toMatch(/NaN|Infinity/);
      });
    }
  }

  it('clampPath：範囲内の点は変えず、範囲外・非有限は端へ丸める', () => {
    const p: ShapePath = {
      start: { x: -5, y: 10 },
      segs: [
        { t: 'L', x: 1e12, y: Number.NaN },
        { t: 'C', x1: -1e12, y1: 5, x2: 50, y2: 1e9, x: 20, y: 30 },
      ],
    };
    const c = clampPath(p, { totalWidth: 100, totalHeight: 60 });
    expect(c.start).toEqual({ x: 0, y: 10 });
    expect(c.segs[0]).toEqual({ t: 'L', x: 100, y: 0 });
    expect(c.segs[1]).toEqual({ t: 'C', x1: 0, y1: 5, x2: 50, y2: 60, x: 20, y: 30 });
  });
});

describe('TD-88-4 太さ・線種', () => {
  it('太さ：EMU/9525 を 1〜16px に丸める（0・未指定は 1px）', () => {
    expect(strokeWidthPx(0)).toBe(1);
    expect(strokeWidthPx(undefined)).toBe(1);
    expect(strokeWidthPx(Number.NaN)).toBe(1);
    expect(strokeWidthPx(-5)).toBe(1);
    expect(strokeWidthPx(6350)).toBe(1); // 0.667px → 1
    expect(strokeWidthPx(9525)).toBe(1);
    expect(strokeWidthPx(12700)).toBeCloseTo(1.3333, 3);
    expect(strokeWidthPx(25400)).toBeCloseTo(2.6667, 3);
    expect(strokeWidthPx(152400)).toBe(16);
    expect(strokeWidthPx(152400 * 10)).toBe(16);
    expect(strokeWidthPx(1e12)).toBe(16);
  });

  it('線種：solid は実線、dot・sysDot は点線、その他は破線', () => {
    expect(dashArray('solid', 2)).toEqual([]);
    expect(dashArray(undefined, 2)).toEqual([]);
    expect(dashArray('dot', 2)).toEqual([2, 4]);
    expect(dashArray('sysDot', 3)).toEqual([3, 6]);
    for (const d of [
      'dash',
      'lgDash',
      'dashDot',
      'lgDashDot',
      'lgDashDotDot',
      'sysDash',
      'sysDashDot',
      'sysDashDotDot',
    ]) {
      expect(dashArray(d, 2)).toEqual([8, 6]);
    }
  });
});

describe('TD-88-5 矢じり', () => {
  const spec = (
    type: LineEnd['type'],
    w: LineEnd['w'] = 'med',
    len: LineEnd['len'] = 'med',
  ): LineEnd => ({
    type,
    w,
    len,
  });

  it('大きさ：max(太さ,2)×係数（sm 2・med 3・lg 5）。上限 48px', () => {
    expect(arrowDimension(1, 'sm')).toBe(4);
    expect(arrowDimension(1, 'med')).toBe(6);
    expect(arrowDimension(1, 'lg')).toBe(10);
    expect(arrowDimension(4, 'sm')).toBe(8);
    expect(arrowDimension(4, 'med')).toBe(12);
    expect(arrowDimension(4, 'lg')).toBe(20);
    expect(arrowDimension(16, 'lg')).toBe(48);
    expect(arrowDimension(16, 'med')).toBe(48);
    expect(arrowDimension(10, 'lg')).toBe(48);
    expect(arrowDimension(1, undefined)).toBe(6);
    expect(arrowDimension(1, 'weird')).toBe(6);
  });

  // 右向き（外向き (1,0)）・端点 (100, 50)・太さ 1
  const tip: Pt = { x: 100, y: 50 };
  const dir: Pt = { x: 1, y: 0 };

  for (const size of ['sm', 'med', 'lg'] as const) {
    const L = { sm: 4, med: 6, lg: 10 }[size];
    const W = L;
    it(`triangle ${size}：先端が端点、後ろの 2 点が長さ ${L} 手前に幅 ${W} で開く（塗り・閉じる）`, () => {
      const a = arrowShape('tail', spec('triangle', size, size), tip, dir, 1);
      expect(a).toEqual({
        kind: 'polygon',
        end: 'tail',
        points: [
          { x: 100, y: 50 },
          { x: 100 - L, y: 50 + W / 2 },
          { x: 100 - L, y: 50 - W / 2 },
        ],
        filled: true,
        closed: true,
      });
    });
    it(`stealth ${size}：先端・後ろの 2 点・切れ込み（長さの 0.7 手前）`, () => {
      const a = arrowShape('tail', spec('stealth', size, size), tip, dir, 1);
      if (a.kind !== 'polygon') throw new Error('polygon のはず');
      expect(a.filled).toBe(true);
      expect(a.closed).toBe(true);
      expect(a.points).toHaveLength(4);
      expect(a.points[0]).toEqual({ x: 100, y: 50 });
      expect(a.points[1]!.x).toBeCloseTo(100 - L);
      expect(a.points[2]!.x).toBeCloseTo(100 - L * 0.7);
      expect(a.points[2]!.y).toBeCloseTo(50);
      expect(a.points[3]!.x).toBeCloseTo(100 - L);
      expect(a.points[1]!.y).toBeCloseTo(50 + W / 2);
      expect(a.points[3]!.y).toBeCloseTo(50 - W / 2);
    });
    it(`arrow ${size}：開いた V（塗らない・閉じない）。先端が端点`, () => {
      const a = arrowShape('tail', spec('arrow', size, size), tip, dir, 1);
      expect(a).toEqual({
        kind: 'polygon',
        end: 'tail',
        points: [
          { x: 100 - L, y: 50 + W / 2 },
          { x: 100, y: 50 },
          { x: 100 - L, y: 50 - W / 2 },
        ],
        filled: false,
        closed: false,
      });
    });
    it(`diamond ${size}：中心が端点のひし形`, () => {
      const a = arrowShape('tail', spec('diamond', size, size), tip, dir, 1);
      expect(a).toEqual({
        kind: 'polygon',
        end: 'tail',
        points: [
          { x: 100 + L / 2, y: 50 },
          { x: 100, y: 50 + W / 2 },
          { x: 100 - L / 2, y: 50 },
          { x: 100, y: 50 - W / 2 },
        ],
        filled: true,
        closed: true,
      });
    });
    it(`oval ${size}：中心が端点の楕円（rx=長さ/2・ry=幅/2）`, () => {
      const a = arrowShape('tail', spec('oval', size, size), tip, dir, 1);
      expect(a).toEqual({
        kind: 'ellipse',
        end: 'tail',
        cx: 100,
        cy: 50,
        rx: L / 2,
        ry: W / 2,
        angle: 0,
      });
    });
    it(`その他の type ${size}：塗った三角`, () => {
      const odd = arrowShape(
        'tail',
        { type: 'bogus' as LineEnd['type'], w: size, len: size },
        tip,
        dir,
        1,
      );
      expect(odd).toEqual(arrowShape('tail', spec('triangle', size, size), tip, dir, 1));
    });
  }

  it('幅と長さは別々の係数（w=lg・len=sm）。上限 48 は幅・長さそれぞれに掛かる', () => {
    const a = arrowShape('tail', spec('triangle', 'lg', 'sm'), tip, dir, 4);
    if (a.kind !== 'polygon') throw new Error('polygon のはず');
    // 太さ 4：幅 = 4×5 = 20、長さ = 4×2 = 8
    expect(a.points[1]).toEqual({ x: 92, y: 60 });
    expect(a.points[2]).toEqual({ x: 92, y: 40 });
    const big = arrowShape('tail', spec('triangle', 'lg', 'lg'), tip, dir, 16);
    if (big.kind !== 'polygon') throw new Error('polygon のはず');
    expect(big.points[1]).toEqual({ x: 100 - 48, y: 50 + 24 });
  });

  it('向き：始点（head）は外向き＝逆向き。斜めの線では端の接線に沿う', () => {
    const left = arrowShape('head', spec('triangle'), { x: 10, y: 10 }, { x: -1, y: 0 }, 1);
    if (left.kind !== 'polygon') throw new Error('polygon のはず');
    expect(left.points[1]).toEqual({ x: 16, y: 7 }); // 外向き (-1,0) の左手＝上側
    const ang = arrowShape('tail', spec('oval'), { x: 0, y: 0 }, { x: 0, y: 1 }, 1);
    if (ang.kind !== 'ellipse') throw new Error('ellipse のはず');
    expect(ang.angle).toBeCloseTo(90);
  });

  describe('端の向き（endDirection）', () => {
    it('直線：端の線分の向き（tail＝終点向き、head＝始点から外向き）', () => {
      const p = shapePath('line', [], 30, 40);
      const t = endDirection(p, 'tail');
      const h = endDirection(p, 'head');
      expect(t.x).toBeCloseTo(0.6);
      expect(t.y).toBeCloseTo(0.8);
      expect(h.x).toBeCloseTo(-0.6);
      expect(h.y).toBeCloseTo(-0.8);
    });
    it('折れ線：端の線分（bentConnector3 は右へ出て右へ入る）', () => {
      const p = shapePath('bentConnector3', [], 80, 30);
      expect(endDirection(p, 'tail')).toEqual({ x: 1, y: 0 });
      expect(endDirection(p, 'head')).toEqual({ x: -1, y: 0 });
    });
    it('ベジェ：端に隣る制御点から端点（curvedConnector2：始点は左向き、終点は下向き）', () => {
      const p = shapePath('curvedConnector2', [], 80, 30);
      expect(endDirection(p, 'head')).toEqual({ x: -1, y: 0 });
      expect(endDirection(p, 'tail')).toEqual({ x: 0, y: 1 });
    });
    it('端の線分の長さが 0 なら一つ前の点へ遡る', () => {
      // adj1=100%：x1 = w なので最後の線分 (w,h)→(w,h) が長さ 0。向きは (w,0)→(w,h) の下向き
      const p = shapePath('bentConnector3', [100000], 80, 30);
      expect(endDirection(p, 'tail')).toEqual({ x: 0, y: 1 });
      // adj1=0%：最初の線分 (0,0)→(0,0) が長さ 0。始点から外向きは (0,h)→(0,0) の上向き
      const q = shapePath('bentConnector3', [0], 80, 30);
      expect(endDirection(q, 'head')).toEqual({ x: 0, y: -1 });
    });
    it('全点が同じなら (1, 0)', () => {
      const p = shapePath('line', [], 0, 0);
      expect(endDirection(p, 'tail')).toEqual({ x: 1, y: 0 });
      expect(endDirection(p, 'head')).toEqual({ x: 1, y: 0 });
    });
  });

  describe('線の端の縮め', () => {
    /** 水平線 (10,20)→(110,20) に矢じりを付けて描く */
    const line = (head: LineEnd | null, tail: LineEnd | null, w = 100) =>
      renderLine(
        { x: 10, y: 20, w, h: 0 },
        { geom: 'line', adj: [50000, 50000, 50000], widthEmu: 9525, dash: 'solid', head, tail },
        {},
        BIG,
      );
    const xsOf = (r: ReturnType<typeof renderLine>): [number, number] => {
      const n = nums(r.d);
      return [n[0]! + r.bbox.x, n[2]! + r.bbox.x];
    };

    it('triangle（med）の終点：線は矢じりの長さ 6px 分だけ縮む', () => {
      const [s, e] = xsOf(line(null, spec('triangle')));
      expect(s).toBeCloseTo(10);
      expect(e).toBeCloseTo(104);
    });
    it('triangle の始点：始点側が 6px 縮む', () => {
      const [s, e] = xsOf(line(spec('triangle'), null));
      expect(s).toBeCloseTo(16);
      expect(e).toBeCloseTo(110);
    });
    it('stealth は切れ込みの深さ（長さの 0.7＝4.2px）だけ縮む', () => {
      const [, e] = xsOf(line(null, spec('stealth')));
      expect(e).toBeCloseTo(110 - 4.2, 1);
    });
    it('arrow・diamond・oval は縮めない', () => {
      for (const t of ['arrow', 'diamond', 'oval'] as const) {
        const [s, e] = xsOf(line(spec(t), spec(t)));
        expect(s).toBeCloseTo(10);
        expect(e).toBeCloseTo(110);
      }
    });
    it('短い線で両端を縮めても、始点が終点を追い越さない（按分）', () => {
      const [s, e] = xsOf(line(spec('triangle', 'lg', 'lg'), spec('triangle', 'lg', 'lg'), 8));
      expect(e).toBeGreaterThanOrEqual(s);
      expect(s).toBeGreaterThanOrEqual(10);
      expect(e).toBeLessThanOrEqual(18);
    });
    it('shrinkPath：ベジェは端点と隣の制御点を一緒に動かす。縮める長さは区間の長さを超えない', () => {
      const p = shapePath('curvedConnector2', [], 80, 30);
      const q = shrinkPath(p, 0, 5);
      const last = q.segs[0]!;
      if (last.t !== 'C') throw new Error('C のはず');
      expect(last.x).toBeCloseTo(80);
      expect(last.y).toBeCloseTo(25);
      expect(last.y2).toBeCloseTo(10);
      expect(q.start).toEqual(p.start);
      const over = shrinkPath(shapePath('line', [], 3, 4), 0, 100);
      expect(ends(over).e.x).toBeCloseTo(0);
      expect(ends(over).e.y).toBeCloseTo(0);
    });
    it('shrinkPath：縮める長さが 0 なら同じ経路', () => {
      const p = shapePath('line', [], 30, 40);
      expect(shrinkPath(p, 0, 0)).toBe(p);
    });
    it('縮めても矢じりの位置は元の端点のまま', () => {
      const r = line(null, spec('triangle'));
      const a = r.heads[0]!;
      if (a.kind !== 'polygon') throw new Error('polygon のはず');
      expect(a.points[0]!.x + r.bbox.x).toBeCloseTo(110);
      expect(a.points[0]!.y + r.bbox.y).toBeCloseTo(20);
    });
  });

  it('arrowPolygonData：閉じる図形は Z、開いた V は Z なし。数値だけ', () => {
    const tri = arrowShape('tail', spec('triangle'), tip, dir, 1) as Extract<
      ArrowShape,
      { kind: 'polygon' }
    >;
    expect(arrowPolygonData(tri)).toBe('M100 50 L94 53 L94 47 Z');
    const v = arrowShape('tail', spec('arrow'), tip, dir, 1) as Extract<
      ArrowShape,
      { kind: 'polygon' }
    >;
    expect(arrowPolygonData(v)).toBe('M94 53 L100 50 L94 47');
  });
});

describe('TD-88-6 外接矩形（renderLine の bbox）', () => {
  const plain = (
    geom: string,
    adj: [number, number, number] = [50000, 50000, 50000],
    widthEmu = 9525,
  ) => ({
    geom: geom as 'line',
    adj,
    widthEmu,
    dash: 'solid' as const,
    head: null,
    tail: null,
  });

  it('水平線（高さ 0）でも太さ分の高さを持つ', () => {
    const r = renderLine({ x: 10, y: 20, w: 100, h: 0 }, plain('line'), {}, BIG);
    expect(r.bbox).toEqual({ x: 9, y: 19, w: 102, h: 2 });
    expect(r.strokeWidth).toBe(1);
    expect(r.d).toBe('M1 1 L101 1');
  });

  it('垂直線（幅 0）でも太さ分の幅を持つ。太い線は太さ/2 を四方に足す', () => {
    const r = renderLine({ x: 50, y: 10, w: 0, h: 80 }, plain('line', undefined, 152400), {}, BIG);
    expect(r.strokeWidth).toBe(16);
    expect(r.bbox).toEqual({ x: 42, y: 2, w: 16, h: 96 });
  });

  it('矩形の外を通る線（U 字）：外接矩形は矩形より左右へ広い', () => {
    const w = 12700 / EMU;
    const rect = { x: 400, y: 100, w, h: 60 };
    const r = renderLine(rect, plain('bentConnector3', [1_800_000, 50000, 50000], 12700), {}, BIG);
    expect(r.bbox.x).toBeLessThanOrEqual(400);
    expect(r.bbox.x + r.bbox.w).toBeGreaterThanOrEqual(424);
    expect(r.bbox.y).toBeLessThanOrEqual(100);
    expect(r.bbox.y + r.bbox.h).toBeGreaterThanOrEqual(160);
  });

  it('回転した線：アンカーの縦横入れ替え後の経路で外接矩形を出す（90° の水平線は垂直になる）', () => {
    // 回転前 100×0 の水平線を 90° 回すと、中心 (60,50) まわりの垂直線（x=60, y 0〜100）
    const r = renderLine({ x: 60, y: 0, w: 0, h: 100 }, plain('line'), { rot: 90 * DEG }, BIG);
    expect(r.path.start.x).toBeCloseTo(60);
    expect(r.path.start.y).toBeCloseTo(0);
    expect(ends(r.path).e.y).toBeCloseTo(100);
  });

  it('矢じりの大きさの半分が太さの半分より大きければそちらを足す', () => {
    const r = renderLine(
      { x: 100, y: 100, w: 100, h: 0 },
      { ...plain('line'), tail: { type: 'triangle', w: 'lg', len: 'lg' } },
      {},
      BIG,
    );
    // 太さ 1・矢じり 10 → pad 5：x は 100−5=95〜205。y は矢じりの翼（±5）に線の太さ/2 を足して 94.5〜105.5 → 94〜106
    expect(r.bbox).toEqual({ x: 95, y: 94, w: 110, h: 12 });
  });

  it('シートの範囲に丸める（原点・全幅・全高）', () => {
    const r = renderLine(
      { x: 0, y: 5, w: 50, h: 0 },
      plain('line'),
      {},
      { totalWidth: 50, totalHeight: 100 },
    );
    expect(r.bbox.x).toBe(0);
    expect(r.bbox.x + r.bbox.w).toBe(50);
    expect(r.bbox.y).toBeGreaterThanOrEqual(0);
  });

  it('d は外接矩形の左上を原点にした数値だけ。曲線は C を含む。NaN・Infinity は出さない', () => {
    const r = renderLine({ x: 100, y: 100, w: 80, h: 30 }, plain('curvedConnector3'), {}, BIG);
    expect(r.d).toMatch(/^M[\d.-]+ [\d.-]+ C/);
    expect(r.d).not.toMatch(/NaN|Infinity/);
    const bad = renderLine(
      { x: Number.NaN, y: 0, w: Number.POSITIVE_INFINITY, h: 10 },
      plain('line'),
      {},
      BIG,
    );
    expect(bad.d).not.toMatch(/NaN|Infinity/);
    expect(Number.isFinite(bad.bbox.x + bad.bbox.y + bad.bbox.w + bad.bbox.h)).toBe(true);
  });

  it('pathData：原点を引き、小数 2 桁に丸め、-0 は 0', () => {
    const p: ShapePath = { start: { x: 10.123, y: 20 }, segs: [{ t: 'L', x: 10, y: 20.456 }] };
    expect(pathData(p)).toBe('M10.12 20 L10 20.46');
    expect(pathData(p, 10, 20)).toBe('M0.12 0 L0 0.46');
    expect(pathData({ start: { x: -0.001, y: Number.NaN }, segs: [] })).toBe('M0 0');
  });

  it('heads は始点・終点の順で、外接矩形の左上を原点にした座標', () => {
    const r = renderLine(
      { x: 100, y: 100, w: 100, h: 0 },
      {
        ...plain('line'),
        head: { type: 'oval', w: 'med', len: 'med' },
        tail: { type: 'diamond', w: 'med', len: 'med' },
      },
      {},
      BIG,
    );
    expect(r.heads.map((h) => h.end)).toEqual(['head', 'tail']);
    const [h0, h1] = r.heads;
    if (h0?.kind !== 'ellipse' || h1?.kind !== 'polygon') throw new Error('型が違う');
    expect(h0.cx + r.bbox.x).toBeCloseTo(100);
    expect(h0.angle).toBeCloseTo(180);
    expect(h1.points[0]!.x + r.bbox.x).toBeCloseTo(203); // 終点 (200) の先へ 長さ/2 (3)
  });
});

// ---------------------------------------------------------------------------
// U1：ref-shapes.xlsx（Excel 16 の COM で作成した基準ファイル）の線で、向き・反転・回転の順序を照合する
// ---------------------------------------------------------------------------

/** 接続先の図形の矩形（EMU：x, y, cx, cy。キーは `cNvPr id`）。すべて rect・回転なし */
const TARGETS: Record<number, [number, number, number, number]> = {
  92: [254000, 14097000, 889000, 304800], // srcTb
  93: [1905000, 14859000, 889000, 381000], // dstRect
  97: [3810000, 14097000, 762000, 254000], // uA
  98: [3810000, 14859000, 1270000, 254000], // uB
  101: [6096000, 14859000, 762000, 254000], // uC
  102: [7620000, 14859000, 762000, 254000], // uD
  105: [9144000, 14097000, 762000, 304800], // vTop
  106: [9906000, 14986000, 762000, 304800], // vBottom
  109: [11430000, 14859000, 762000, 304800], // wA
  110: [10922000, 14097000, 762000, 304800], // wB
  114: [254000, 15494000, 762000, 381000], // pA_bent4
  115: [1905000, 16510000, 762000, 381000], // pB_bent4
  117: [3302000, 15494000, 762000, 381000], // pA_curved4
  118: [4953000, 16510000, 762000, 381000], // pB_curved4
  120: [6350000, 15494000, 762000, 381000], // pA_bent5
  121: [8001000, 16510000, 762000, 381000], // pB_bent5
  123: [9398000, 15494000, 762000, 381000], // pA_curved5
  124: [11049000, 16510000, 762000, 381000], // pB_curved5
  126: [12446000, 15494000, 762000, 381000], // pA_curved2
  127: [14097000, 16510000, 762000, 381000], // pB_curved2
};

/** 接続線：[名前, xfrm(EMU: x, y, cx, cy), rot(1/60000°), 始点の接続先 id, 始点の site idx, 終点の接続先 id, 終点の site idx] */
const CONNECTORS: [
  string,
  [number, number, number, number],
  number,
  number,
  number,
  number,
  number,
][] = [
  ['linkElbow', [977900, 14122400, 647700, 1206500], 16200000, 92, 2, 93, 1],
  ['linkStraight', [1143000, 14249400, 1206500, 609600], 0, 92, 3, 93, 0],
  ['linkCurve', [1295400, 13804900, 457200, 1651000], 16200000, 92, 2, 93, 0],
  ['uRightRight', [4572000, 14224000, 508000, 762000], 0, 97, 3, 98, 3],
  ['uLeftLeft', [3810000, 14224000, 12700, 762000], 10800000, 97, 1, 98, 1],
  ['uTopTop', [7239000, 14097000, 12700, 1524000], 5400000, 101, 0, 102, 0],
  ['uBottomBottom', [7239000, 14351000, 12700, 1524000], 16200000, 101, 2, 102, 2],
  ['vElbow', [9613900, 14312900, 584200, 762000], 16200000, 105, 2, 106, 0],
  ['bentRightTop', [9906000, 14249400, 381000, 736600], 0, 105, 3, 106, 0],
  ['bentTopBottom', [11328400, 14376400, 457200, 508000], 16200000, 109, 0, 110, 2],
  ['bentLeftRight', [11430000, 14249400, 254000, 762000], 10800000, 109, 1, 110, 3],
  ['bentBack', [11430000, 14249400, 254000, 762000], 0, 110, 3, 109, 1],
  ['bent4', [666750, 15462250, 1206500, 1270000], 16200000, 114, 0, 115, 1],
  ['curved4', [3714750, 15462250, 1206500, 1270000], 16200000, 117, 0, 118, 1],
  ['bent5', [6858000, 15367000, 1397000, 1651000], 16200000, 120, 0, 121, 2],
  ['curved5', [9906000, 15367000, 1397000, 1651000], 16200000, 123, 0, 124, 2],
  ['curved2', [13049250, 15652750, 825500, 1270000], 16200000, 126, 2, 127, 1],
];

/** 始点・終点を指定して作った線：[名前, xfrm(EMU), rot, flipH, flipV, [x1, y1, x2, y2](pt), 作成後に掛けた反転('H'・'V'・''=指定値が最終形)] */
const SPECIFIED: [
  string,
  [number, number, number, number],
  number,
  boolean,
  boolean,
  [number, number, number, number],
  string,
][] = [
  ['lnPlain', [254000, 1524000, 1016000, 381000], 0, false, false, [20, 120, 100, 150], ''],
  ['lnFlipV', [1524000, 1524000, 1016000, 381000], 0, false, true, [120, 150, 200, 120], ''],
  ['lnFlipH', [3429000, 1524000, 635000, 381000], 0, true, false, [320, 120, 270, 150], ''],
  ['lnFlipHV', [4318000, 1524000, 1016000, 381000], 0, true, true, [420, 150, 340, 120], ''],
  ['lnRot90', [254000, 2476500, 1016000, 508000], 5400000, false, false, [20, 195, 100, 235], ''],
  [
    'lnRot180',
    [1778000, 2476500, 1016000, 508000],
    10800000,
    false,
    false,
    [140, 195, 220, 235],
    '',
  ],
  [
    'lnRot270',
    [3302000, 2476500, 1016000, 508000],
    16200000,
    false,
    false,
    [260, 195, 340, 235],
    '',
  ],
  [
    'lnFlipVRot90',
    [4826000, 2476500, 1016000, 508000],
    5400000,
    false,
    true,
    [380, 195, 460, 235],
    'V',
  ],
  [
    'lnFlipHRot180',
    [6350000, 2476500, 1016000, 508000],
    10800000,
    true,
    false,
    [500, 195, 580, 235],
    'H',
  ],
  ['lnHoriz', [7874000, 2476500, 1016000, 0], 0, false, false, [620, 195, 700, 195], ''],
  ['lnVert', [9144000, 2476500, 0, 508000], 0, false, false, [720, 195, 720, 235], ''],
  ['bent3', [254000, 2921000, 1270000, 508000], 0, false, false, [20, 230, 120, 270], ''],
  ['bent3FlipV', [2032000, 2921000, 1270000, 508000], 0, false, true, [160, 270, 260, 230], ''],
  ['bent3FlipH', [3810000, 2921000, 1270000, 508000], 0, true, false, [400, 230, 300, 270], ''],
  ['bent3Adj', [5588000, 2921000, 1270000, 508000], 0, false, false, [440, 230, 540, 270], ''],
  ['bent3AdjNeg', [7366000, 2921000, 1270000, 508000], 0, false, false, [580, 230, 680, 270], ''],
  ['bent3AdjBig', [9144000, 2921000, 1016000, 508000], 0, false, false, [720, 230, 800, 270], ''],
  ['curve', [254000, 3619500, 1270000, 508000], 0, false, false, [20, 285, 120, 325], ''],
  ['curveFlipV', [2032000, 3619500, 1270000, 508000], 0, false, true, [160, 325, 260, 285], ''],
  ['curveFlipH', [3810000, 3619500, 1270000, 508000], 0, true, false, [400, 285, 300, 325], ''],
  ['curveAdj', [5588000, 3619500, 1270000, 508000], 0, false, false, [440, 285, 540, 325], ''],
  [
    'curveRot90',
    [7366000, 3619500, 1016000, 508000],
    5400000,
    false,
    false,
    [580, 285, 660, 325],
    '',
  ],
  ['arrow_tri', [254000, 4445000, 1651000, 0], 0, false, false, [20, 350, 150, 350], ''],
  ['arrow_open', [2286000, 4445000, 1651000, 0], 0, false, false, [180, 350, 310, 350], ''],
  ['arrow_stealth', [4318000, 4445000, 1651000, 0], 0, false, false, [340, 350, 470, 350], ''],
  ['arrow_diamond', [6350000, 4445000, 1651000, 0], 0, false, false, [500, 350, 630, 350], ''],
  ['arrow_oval', [8382000, 4445000, 1651000, 0], 0, false, false, [660, 350, 790, 350], ''],
  ['arrowMed_tri', [254000, 5143500, 1651000, 0], 0, false, false, [20, 405, 150, 405], ''],
  ['arrowMed_open', [2286000, 5143500, 1651000, 0], 0, false, false, [180, 405, 310, 405], ''],
  ['arrowMed_stealth', [4318000, 5143500, 1651000, 0], 0, false, false, [340, 405, 470, 405], ''],
  ['arrowMed_diamond', [6350000, 5143500, 1651000, 0], 0, false, false, [500, 405, 630, 405], ''],
  ['arrowMed_oval', [8382000, 5143500, 1651000, 0], 0, false, false, [660, 405, 790, 405], ''],
  ['arrowShortWide', [254000, 5842000, 1651000, 0], 0, false, false, [20, 460, 150, 460], ''],
  ['arrowLongNarrow', [2286000, 5842000, 1651000, 0], 0, false, false, [180, 460, 310, 460], ''],
  ['arrowThick', [4318000, 5842000, 1651000, 0], 0, false, false, [340, 460, 470, 460], ''],
  ['dash_solid', [254000, 6413500, 1016000, 0], 0, false, false, [20, 505, 100, 505], ''],
  ['dash_squareDot', [1524000, 6413500, 1016000, 0], 0, false, false, [120, 505, 200, 505], ''],
  ['dash_roundDot', [2794000, 6413500, 1016000, 0], 0, false, false, [220, 505, 300, 505], ''],
  ['dash_dash', [4064000, 6413500, 1016000, 0], 0, false, false, [320, 505, 400, 505], ''],
  ['dash_dashDot', [5334000, 6413500, 1016000, 0], 0, false, false, [420, 505, 500, 505], ''],
  ['dash_dashDotDot', [6604000, 6413500, 1016000, 0], 0, false, false, [520, 505, 600, 505], ''],
  ['dash_longDash', [7874000, 6413500, 1016000, 0], 0, false, false, [620, 505, 700, 505], ''],
  ['dash_longDashDot', [9144000, 6413500, 1016000, 0], 0, false, false, [720, 505, 800, 505], ''],
];

describe('TD-88-7 U1：ref-shapes.xlsx の線の向き（接続先の辺の中央・指定した端点）', () => {
  let byName: Map<string, SheetDrawing>;

  beforeAll(async () => {
    const book = await openXlsx(
      readFileSync(new URL('../../fixtures/excel/ref-shapes.xlsx', import.meta.url)),
    );
    byName = new Map((await book.openDrawings(1)).drawings.map((d) => [d.name, d]));
  });

  /** EMU の矩形を px にし、アンカー（回転後の外接）と回転前の矩形を返す */
  function rectsOf(
    xf: [number, number, number, number],
    rot: number,
  ): { anchor: PxRect; box: PxRect } {
    const box: PxRect = { x: xf[0] / EMU, y: xf[1] / EMU, w: xf[2] / EMU, h: xf[3] / EMU };
    return { anchor: anchorOf(box, (rot / DEG) % 360), box };
  }

  /** 名前の線を描く（アンカーは xfrm から作る）。読み取った rot・反転・形を使う */
  function draw(
    name: string,
    xf: [number, number, number, number],
  ): { r: ReturnType<typeof renderLine>; d: SheetDrawing } {
    const d = byName.get(name);
    if (!d || d.kind !== 'line' || !d.line) throw new Error(`線が無い: ${name}`);
    const { anchor } = rectsOf(xf, d.rot ?? 0);
    const r = renderLine(anchor, d.line, { rot: d.rot, flipH: d.flipH, flipV: d.flipV }, BIG);
    return { r, d };
  }

  /** 矩形図形（rect）の接続サイト idx（0=上・1=左・2=下・3=右）の位置（px） */
  function site(id: number, idx: number): Pt {
    const t = TARGETS[id]!;
    const x = t[0] / EMU;
    const y = t[1] / EMU;
    const w = t[2] / EMU;
    const h = t[3] / EMU;
    switch (idx) {
      case 0:
        return { x: x + w / 2, y };
      case 1:
        return { x, y: y + h / 2 };
      case 2:
        return { x: x + w / 2, y: y + h };
      default:
        return { x: x + w, y: y + h / 2 };
    }
  }

  it('線の件数（接続線 17・指定して作った線 43）', () => {
    expect(CONNECTORS).toHaveLength(17);
    expect(SPECIFIED).toHaveLength(43);
  });

  for (const [name, xf, rot, stId, stIdx, enId, enIdx] of CONNECTORS) {
    it(`${name}：始点＝接続元(${stId}, site ${stIdx})の辺の中央・終点＝接続先(${enId}, site ${enIdx})の辺の中央（1px 以内）`, () => {
      const { r, d } = draw(name, xf);
      expect(d.rot ?? 0).toBe(rot);
      const { s, e } = ends(r.path);
      // uLeftLeft は Excel が幅を最小の 1pt（1.33px）にして保存するため、始点が接続元の辺から 1.33px ずれる（Excel 側の丸め）
      const tol = name === 'uLeftLeft' ? 1.5 : 1;
      const want0 = site(stId, stIdx);
      const want1 = site(enId, enIdx);
      expect(Math.abs(s.x - want0.x)).toBeLessThan(tol);
      expect(Math.abs(s.y - want0.y)).toBeLessThan(tol);
      expect(Math.abs(e.x - want1.x)).toBeLessThan(tol);
      expect(Math.abs(e.y - want1.y)).toBeLessThan(tol);
    });
  }

  for (const [name, xf, rot, flipH, flipV, [x1, y1, x2, y2], preFlip] of SPECIFIED) {
    it(`${name}：始点・終点が指定した値（pt×4/3）に 1px 以内で一致`, () => {
      const { r, d } = draw(name, xf);
      expect(d.rot ?? 0).toBe(rot);
      expect(!!d.flipH).toBe(flipH);
      expect(!!d.flipV).toBe(flipV);
      const { box } = rectsOf(xf, rot);
      const k = 4 / 3;
      let s: Pt = { x: x1 * k, y: y1 * k };
      let e: Pt = { x: x2 * k, y: y2 * k };
      // 作成後に Flip・Rotation を掛けた線は、指定値が反転・回転の前（＝反転 → 回転の順に掛ける）
      if (preFlip === 'H' || preFlip === 'V') {
        const cx = box.x + box.w / 2;
        const cy = box.y + box.h / 2;
        const f = (p: Pt): Pt =>
          preFlip === 'H' ? { x: 2 * cx - p.x, y: p.y } : { x: p.x, y: 2 * cy - p.y };
        s = f(s);
        e = f(e);
      }
      if (rot !== 0) {
        s = rot90s(s, box, rot / DEG);
        e = rot90s(e, box, rot / DEG);
      }
      const got = ends(r.path);
      expect(Math.abs(got.s.x - s.x)).toBeLessThan(1);
      expect(Math.abs(got.s.y - s.y)).toBeLessThan(1);
      expect(Math.abs(got.e.x - e.x)).toBeLessThan(1);
      expect(Math.abs(got.e.y - e.y)).toBeLessThan(1);
    });
  }

  it('U 字線（uLeftLeft・uTopTop・uBottomBottom・uRightRight）の張り出しは丸めで縮まない', () => {
    // uLeftLeft：接続元の左辺（x=400px）から左へ 22.67px 出て戻る
    const ll = draw('uLeftLeft', [3810000, 14224000, 12700, 762000]).r;
    expect(Math.min(...pathPoints(ll.path).map((p) => p.x))).toBeCloseTo(400 - 22.667, 1);
    // uTopTop：接続元 uC の上辺（y = 14859000/9525 = 1560px）から上へ 22.67px（幅 1pt の箱の中心合わせで 0.67px 内側）
    const tt = draw('uTopTop', [7239000, 14097000, 12700, 1524000]).r;
    expect(
      Math.abs(Math.min(...pathPoints(tt.path).map((p) => p.y)) - (14859000 / EMU - 22.667)),
    ).toBeLessThan(1);
    // uBottomBottom：接続元の下辺（(14859000+254000)/9525 ではなく uC の下辺）から下へ
    const bb = draw('uBottomBottom', [7239000, 14351000, 12700, 1524000]).r;
    const uCbottom = (14859000 + 254000) / EMU;
    expect(
      Math.abs(Math.max(...pathPoints(bb.path).map((p) => p.y)) - (uCbottom + 22.667)),
    ).toBeLessThan(1);
    // uRightRight：adj1=145000・幅 508000EMU（53.3px）→ x1 = 77.3px。右辺から約 24px 右へ
    const rr = draw('uRightRight', [4572000, 14224000, 508000, 762000]).r;
    expect(Math.max(...pathPoints(rr.path).map((p) => p.x))).toBeCloseTo(
      4572000 / EMU + (508000 / EMU) * 1.45,
      1,
    );
    // 外接矩形も張り出しを含む
    expect(ll.bbox.x).toBeLessThanOrEqual(400 - 22);
  });
});

/** `transformPath` を直接使う補助（未使用警告の回避と、反転→回転の順序の確認） */
describe('TD-88-8 反転 → 回転の順序', () => {
  it('flipH と 90° 回転：反転してから回す（回してから反転すると終点が違う場所になる）', () => {
    // 100×40 の直線（始点 (0,0)・終点 (100,40)）を flipH → 始点 (100,0)・終点 (0,40)
    // 中心 (50,20)。90° 時計回り：d=(50,-20) → (20,50) → (70,70)、d=(-50,20) → (-20,-50) → (30,-30)
    const box: PxRect = { x: 0, y: 0, w: 100, h: 40 };
    const p = transformPath(shapePath('line', [], 100, 40), box, { rot: 90 * DEG, flipH: true });
    const { s, e } = ends(p);
    expect(s.x).toBeCloseTo(70);
    expect(s.y).toBeCloseTo(70);
    expect(e.x).toBeCloseTo(30);
    expect(e.y).toBeCloseTo(-30);
  });
});
