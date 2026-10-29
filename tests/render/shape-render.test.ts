import { describe, expect, it } from 'vitest';
import type { SheetDrawing } from '../../src/core/index';
import type { DrawingText } from '../../src/core/types';
import {
  layoutDrawings,
  drawingExtent,
  pickVisible,
  EMU_PER_PX,
  type PxRect,
} from '../../src/render/drawings';
import {
  clampPath,
  pathData,
  pathPoints,
  renderShape,
  subpathsData,
  transformPath,
  type ShapePath,
} from '../../src/render/shapes';
import { createGeometry } from '../../src/core/index';
import { emptyLayout } from '../../src/core/types';

/**
 * 図形の形の描画（純ロジック。ISSUE-92・ADR-0028 §決定 4・5・8）：経路の型の拡張（部分経路・閉じ）、
 * 外接矩形、文字領域、代替・見えない図形の判定、配置（`layoutDrawings`）・可視判定・`drawingExtent`。
 */

const BIG = { totalWidth: 1_000_000, totalHeight: 1_000_000 };
const DEG = 60000;
const px = (n: number): number => n * EMU_PER_PX;

/** 閉じた 3 点の経路（矩形 20×10 の三角形）と、2 つ目の部分経路 */
const TRI: ShapePath = {
  start: { x: 0, y: 0 },
  segs: [
    { t: 'L', x: 20, y: 0 },
    { t: 'C', x1: 20, y1: 5, x2: 15, y2: 10, x: 10, y: 10 },
  ],
  closed: true,
};
const BOX = { x: 100, y: 200, w: 20, h: 10 };

/** 反転 → 中心で時計回りに回転した点（テストでの独立した計算） */
function expectedPoint(p: { x: number; y: number }, flipH: boolean, flipV: boolean, deg: number) {
  const lx = flipH ? BOX.w - p.x : p.x;
  const ly = flipV ? BOX.h - p.y : p.y;
  const dx = lx - BOX.w / 2;
  const dy = ly - BOX.h / 2;
  const r = (deg * Math.PI) / 180;
  return {
    x: BOX.x + BOX.w / 2 + dx * Math.cos(r) - dy * Math.sin(r),
    y: BOX.y + BOX.h / 2 + dx * Math.sin(r) + dy * Math.cos(r),
  };
}

describe('TD-92-1 経路の写像：部分経路・閉じ・反転（4 通り）× 回転（0・90・180・270°）', () => {
  for (const flipH of [false, true]) {
    for (const flipV of [false, true]) {
      for (const deg of [0, 90, 180, 270]) {
        it(`flipH=${flipH} flipV=${flipV} rot=${deg}：全点（制御点を含む）が反転 → 回転の式どおり。閉じは保たれる`, () => {
          const out = transformPath(TRI, BOX, { rot: deg * DEG, flipH, flipV });
          expect(out.closed).toBe(true);
          const got = pathPoints(out);
          const want = pathPoints(TRI).map((p) => expectedPoint(p, flipH, flipV, deg));
          expect(got).toHaveLength(want.length);
          got.forEach((g, i) => {
            expect(g.x).toBeCloseTo(want[i]!.x, 9);
            expect(g.y).toBeCloseTo(want[i]!.y, 9);
          });
        });
      }
    }
  }

  it('閉じていない経路は closed を持たない（線の結果は変えない）', () => {
    const line: ShapePath = { start: { x: 0, y: 0 }, segs: [{ t: 'L', x: 5, y: 5 }] };
    expect('closed' in transformPath(line, BOX, {})).toBe(false);
    expect('closed' in clampPath(line, BIG)).toBe(false);
  });

  it('clampPath：閉じを保ち、全点をシートの範囲に丸める', () => {
    const out = clampPath(
      { start: { x: -5, y: 3 }, segs: [{ t: 'L', x: 5e9, y: 9 }], closed: true },
      { totalWidth: 100, totalHeight: 50 },
    );
    expect(out.closed).toBe(true);
    expect(out.start).toEqual({ x: 0, y: 3 });
    expect(out.segs[0]).toEqual({ t: 'L', x: 100, y: 9 });
  });

  it('pathData：閉じた経路は末尾に Z。部分経路の列は空白でつなぐ。原点を引く', () => {
    expect(pathData(TRI)).toBe('M0 0 L20 0 C20 5 15 10 10 10 Z');
    expect(pathData(TRI, 100, 200)).toBe('M-100 -200 L-80 -200 C-80 -195 -85 -190 -90 -190 Z');
    expect(pathData({ start: { x: 1, y: 2 }, segs: [{ t: 'L', x: 3, y: 4 }] })).toBe('M1 2 L3 4');
    const open: ShapePath = { start: { x: 0, y: 0 }, segs: [{ t: 'L', x: 4, y: 4 }] };
    expect(subpathsData([TRI, open], 0, 0)).toBe('M0 0 L20 0 C20 5 15 10 10 10 Z M0 0 L4 4');
    expect(subpathsData([], 0, 0)).toBe('');
  });
});

/** 図形 1 つ（既定は枠線のある矩形 100×50） */
function shape(over: Partial<SheetDrawing> = {}): SheetDrawing {
  return {
    anchor: {
      type: 'oneCell',
      from: { col: 0, colOff: px(10), row: 0, rowOff: px(20) },
      cx: px(100),
      cy: px(50),
    },
    kind: 'shape',
    name: 's',
    descr: '',
    mediaPath: null,
    geom: { type: 'preset', prst: 'rect', adj: [] },
    hasOutline: true,
    hasFill: false,
    ...over,
  };
}

const geo = createGeometry(emptyLayout(), 1000, 1000);
/** A1 の左上基準の矩形（x=10・y=20・100×50 は上の既定のアンカー） */
const ANCHOR: PxRect = { x: 10, y: 20, w: 100, h: 50 };

const TEXT: DrawingText = {
  anchor: 't',
  vert: 'horz',
  noWrap: false,
  paragraphs: [{ align: 'l', bullet: null, runs: [{ text: 'あ' }] }],
} as unknown as DrawingText;

describe('TD-92-2 renderShape：外接矩形・経路・枠線', () => {
  it('矩形（枠線 1px・塗りなし）：経路は閉じた 1 本。外接矩形は miter の張り出し（miterlimit 4 の半分 × 太さ）を足して整数へ広げる', () => {
    const r = renderShape(ANCHOR, shape(), BIG);
    expect(r.kind).toBe('drawn');
    if (r.kind !== 'drawn') return;
    expect(r.rect).toEqual({ x: 8, y: 18, w: 104, h: 54 });
    expect(r.box).toEqual(ANCHOR);
    expect(r.faces).toHaveLength(1);
    expect(r.faces[0]!.d).toBe('M2 2 L102 2 L102 52 L2 52 Z');
    expect(r.faces[0]).toMatchObject({ fill: false, stroke: true, shade: null });
    expect(r.strokeWidth).toBe(1);
    expect(r.dash).toEqual([]);
  });

  it('枠線の太さ・線種：太さは線と同じ換算（EMU/9525・1〜16px）、dash は破線 [4w,3w]、dot は [w,2w]', () => {
    const wide = renderShape(ANCHOR, shape({ outline: { widthEmu: px(6), dash: 'dash' } }), BIG);
    if (wide.kind !== 'drawn') throw new Error('drawn でない');
    expect(wide.strokeWidth).toBe(6);
    expect(wide.dash).toEqual([24, 18]);
    // 余白は miterlimit(4)/2 × 太さ = 12px。x は 10−12 < 0 なのでシートの左端（0）で止まる
    expect(wide.rect).toEqual({ x: 0, y: 8, w: 122, h: 74 });
    const dot = renderShape(ANCHOR, shape({ outline: { widthEmu: px(2), dash: 'sysDot' } }), BIG);
    if (dot.kind !== 'drawn') throw new Error('drawn でない');
    expect(dot.dash).toEqual([2, 4]);
    const huge = renderShape(ANCHOR, shape({ outline: { widthEmu: px(500), dash: 'solid' } }), BIG);
    if (huge.kind !== 'drawn') throw new Error('drawn でない');
    expect(huge.strokeWidth).toBe(16);
  });

  it('塗りだけ（枠線なし）：外接矩形はアンカーと同じ。線は引かない', () => {
    const r = renderShape(ANCHOR, shape({ hasOutline: false, hasFill: true }), BIG);
    if (r.kind !== 'drawn') throw new Error('drawn でない');
    expect(r.rect).toEqual(ANCHOR);
    expect(r.faces[0]).toMatchObject({ fill: true, stroke: false, shade: null });
  });

  it('回転 90°（アンカーは縦横が入れ替わった外接）：回転前の矩形に戻して経路を作り、アンカーの中心で回す', () => {
    // 回転前は 100×50（中心 (60,70)）。アンカーは 50×100 で、同じ中心
    const anchor: PxRect = { x: 35, y: 20, w: 50, h: 100 };
    const r = renderShape(
      anchor,
      shape({ rot: 90 * DEG, hasOutline: false, hasFill: true }),
      BIG,
      false,
    );
    if (r.kind !== 'drawn') throw new Error('drawn でない');
    expect(r.box).toEqual({ x: 10, y: 45, w: 100, h: 50 });
    // 回転後の経路の外接 = x:35〜85・y:20〜120（枠線なしなのでアンカーと同じ）
    expect(r.rect).toEqual({ x: 35, y: 20, w: 50, h: 100 });
    const pts = r.faces[0]!.subpaths[0]!;
    expect(pts.start.x).toBeCloseTo(85, 6);
    expect(pts.start.y).toBeCloseTo(20, 6);
  });

  it('回転 45° ちょうど＋左右反転の rightArrow：アンカーは回転前の大きさのまま。xfrmExt があれば入れ替えない（細い矢印を太く描かない。ISSUE-93）', () => {
    // Excel のアンカーは 120×53（回転前の大きさ）。値だけの規則（45° 以上は入れ替え）だと 53×120 の太い矢印になる
    const anchor: PxRect = { x: 100, y: 200, w: 120, h: 53 };
    const base = {
      rot: 45 * DEG,
      flipH: true,
      hasOutline: true,
      hasFill: false,
      geom: { type: 'preset' as const, prst: 'rightArrow', adj: [] },
    };
    const withExt = renderShape(
      anchor,
      shape({ ...base, xfrmExt: { cx: px(120), cy: px(53) } }),
      BIG,
    );
    if (withExt.kind !== 'drawn') throw new Error('drawn でない');
    expect(withExt.box).toEqual(anchor);
    // ext が無ければ従来どおり（45° は入れ替える）
    const noExt = renderShape(anchor, shape(base), BIG);
    if (noExt.kind !== 'drawn') throw new Error('drawn でない');
    expect(noExt.box).toEqual({ x: 133.5, y: 166.5, w: 53, h: 120 });
    // 入れ替えたアンカー（53×120）で 90° の通常の図形は、ext が 120×53 なら従来どおり入れ替える
    const r90 = renderShape(
      { x: 133.5, y: 166.5, w: 53, h: 120 },
      shape({ ...base, rot: 90 * DEG, xfrmExt: { cx: px(120), cy: px(53) } }),
      BIG,
    );
    if (r90.kind !== 'drawn') throw new Error('drawn でない');
    expect(r90.box).toEqual(anchor);
  });

  it('グループの子（preRotated）：矩形は xfrm のまま（縦横を入れ替えない）', () => {
    const r = renderShape(ANCHOR, shape({ rot: 90 * DEG }), BIG, true);
    if (r.kind !== 'drawn') throw new Error('drawn でない');
    expect(r.box).toEqual(ANCHOR);
  });

  it('矩形の外へ出る形（吹き出しの尾）：外接矩形は経路の外接とアンカーの矩形を合わせたもの', () => {
    const r = renderShape(
      ANCHOR,
      shape({ geom: { type: 'preset', prst: 'wedgeRectCallout', adj: [] } }),
      BIG,
    );
    if (r.kind !== 'drawn') throw new Error('drawn でない');
    // 既定の adj2=62500 → 尾の先は矩形の下端より 12.5 * ... 外側（y = 45+0.625*50 = 76.25 より下）
    expect(r.rect.y).toBeLessThanOrEqual(19);
    expect(r.rect.y + r.rect.h).toBeGreaterThanOrEqual(77);
    expect(r.rect.x).toBeLessThanOrEqual(ANCHOR.x);
    expect(r.rect.x + r.rect.w).toBeGreaterThanOrEqual(ANCHOR.x + ANCHOR.w);
  });

  it('シート範囲への丸め：極端な調整値で経路が巨大になっても外接矩形は範囲の中', () => {
    const r = renderShape(
      ANCHOR,
      shape({
        geom: {
          type: 'preset',
          prst: 'wedgeRectCallout',
          adj: [
            ['adj1', 1e9],
            ['adj2', 1e9],
          ],
        },
      }),
      { totalWidth: 500, totalHeight: 400 },
    );
    if (r.kind !== 'drawn') throw new Error('drawn でない');
    for (const f of r.faces) {
      for (const sp of f.subpaths) {
        for (const p of pathPoints(sp)) {
          expect(p.x).toBeLessThanOrEqual(500);
          expect(p.y).toBeLessThanOrEqual(400);
          expect(p.x).toBeGreaterThanOrEqual(0);
          expect(p.y).toBeGreaterThanOrEqual(0);
        }
      }
    }
    expect(r.rect.x + r.rect.w).toBeLessThanOrEqual(Math.max(500, ANCHOR.x + ANCHOR.w));
    expect(r.rect.y + r.rect.h).toBeLessThanOrEqual(Math.max(400, ANCHOR.y + ANCHOR.h));
  });
});

describe('TD-92-3 renderShape：path ごとの fill・stroke と面の濃淡', () => {
  it('円柱（can）：path が複数。fill="lighten"・"none" と stroke=false を反映する', () => {
    const r = renderShape(
      ANCHOR,
      shape({ hasFill: true, geom: { type: 'preset', prst: 'can', adj: [] } }),
      BIG,
    );
    if (r.kind !== 'drawn') throw new Error('drawn でない');
    // can は 3 path：本体（norm）・上面（lighten）・上の縁（fill none）
    expect(r.faces.map((f) => [f.fill, f.stroke, f.shade])).toEqual([
      [true, false, null],
      [true, false, 'lighten'],
      [false, true, null],
    ]);
  });

  it('塗りが無い図形では濃淡の膜も付けない', () => {
    const r = renderShape(
      ANCHOR,
      shape({ hasFill: false, geom: { type: 'preset', prst: 'can', adj: [] } }),
      BIG,
    );
    if (r.kind !== 'drawn') throw new Error('drawn でない');
    expect(r.faces.every((f) => f.shade === null && !f.fill)).toBe(true);
  });

  it('何も塗らず線も引かない path は出さない', () => {
    const r = renderShape(
      ANCHOR,
      shape({ hasOutline: false, hasFill: true, geom: { type: 'preset', prst: 'can', adj: [] } }),
      BIG,
    );
    if (r.kind !== 'drawn') throw new Error('drawn でない');
    expect(r.faces).toHaveLength(2);
  });
});

describe('TD-92-4 renderShape：文字領域（回転前の矩形の左上基準）', () => {
  const textShape = (over: Partial<SheetDrawing> = {}): SheetDrawing =>
    shape({ text: TEXT, ...over });
  const textRectOf = (d: SheetDrawing, anchor: PxRect = ANCHOR, grp = false) => {
    const r = renderShape(anchor, d, BIG, grp);
    if (r.kind !== 'drawn') throw new Error('drawn でない');
    return r.textRect;
  };

  it('rect の形（矩形）：矩形全体', () => {
    expect(textRectOf(textShape())).toEqual({ x: 0, y: 0, w: 100, h: 50 });
  });

  it('楕円：内接矩形（cos45°）', () => {
    const t = textRectOf(textShape({ geom: { type: 'preset', prst: 'ellipse', adj: [] } }));
    const dx = 50 * Math.cos(Math.PI / 4);
    const dy = 25 * Math.sin(Math.PI / 4);
    expect(t.x).toBeCloseTo(50 - dx, 6);
    expect(t.y).toBeCloseTo(25 - dy, 6);
    expect(t.w).toBeCloseTo(2 * dx, 6);
    expect(t.h).toBeCloseTo(2 * dy, 6);
  });

  it('反転：flipH は x を w−r へ、flipV は y を h−b へ（楕円の内接矩形は左右対称なので不変。非対称の吹き出しで確かめる）', () => {
    const base = textRectOf(textShape({ geom: { type: 'preset', prst: 'rightArrow', adj: [] } }));
    const h = textRectOf(
      textShape({ geom: { type: 'preset', prst: 'rightArrow', adj: [] }, flipH: true }),
    );
    const v = textRectOf(
      textShape({ geom: { type: 'preset', prst: 'rightArrow', adj: [] }, flipV: true }),
    );
    expect(h.x).toBeCloseTo(100 - (base.x + base.w), 6);
    expect(h.w).toBeCloseTo(base.w, 6);
    expect(h.y).toBeCloseTo(base.y, 6);
    expect(v.y).toBeCloseTo(50 - (base.y + base.h), 6);
    expect(v.x).toBeCloseTo(base.x, 6);
  });

  it('回転：文字領域は回転前の矩形の中のまま（回転は DOM が矩形ごと掛ける）', () => {
    const anchor: PxRect = { x: 35, y: 20, w: 50, h: 100 };
    expect(textRectOf(textShape({ rot: 90 * DEG }), anchor)).toEqual({ x: 0, y: 0, w: 100, h: 50 });
  });

  it('rect が反転・幅 0 の形は矩形全体（評価器が矩形全体へ戻す）', () => {
    // 幅 0 の矩形でなくても、文字領域は常に正の幅・高さ
    const t = textRectOf(textShape({ geom: { type: 'preset', prst: 'triangle', adj: [] } }));
    expect(t.w).toBeGreaterThan(0);
    expect(t.h).toBeGreaterThan(0);
  });

  it('自由図形で rect が無い：矩形全体', () => {
    const d = textShape({
      geom: {
        type: 'custom',
        def: {
          av: [],
          gd: [],
          paths: [
            {
              w: 10,
              h: 10,
              c: [
                ['M', 0, 0],
                ['L', 10, 10],
              ],
            },
          ],
        },
      },
    });
    expect(textRectOf(d)).toEqual({ x: 0, y: 0, w: 100, h: 50 });
  });
});

describe('TD-92-5 renderShape：見えない図形・文字だけ・代替', () => {
  it('枠も塗りも文字も無い図形は hidden（描かない）', () => {
    expect(renderShape(ANCHOR, shape({ hasOutline: false, hasFill: false }), BIG).kind).toBe(
      'hidden',
    );
  });

  it('枠も塗りも無く文字だけ：面は無く、外接矩形はアンカーの矩形', () => {
    const r = renderShape(ANCHOR, shape({ hasOutline: false, hasFill: false, text: TEXT }), BIG);
    if (r.kind !== 'drawn') throw new Error('drawn でない');
    expect(r.faces).toEqual([]);
    expect(r.rect).toEqual(ANCHOR);
  });

  it('代替：表に無い prst・形を持たない図形は fallback', () => {
    const unknown = shape({ geom: { type: 'preset', prst: 'noSuchShape', adj: [] } });
    expect(renderShape(ANCHOR, unknown, BIG).kind).toBe('fallback');
    expect(renderShape(ANCHOR, { ...unknown, geom: undefined }, BIG).kind).toBe('fallback');
    // __proto__ などの名前でも例外にならず代替
    const proto = shape({ geom: { type: 'preset', prst: '__proto__', adj: [] } });
    expect(renderShape(ANCHOR, proto, BIG).kind).toBe('fallback');
  });

  it('代替：評価に失敗する自由図形（未知の名前）は fallback', () => {
    const bad = shape({
      geom: {
        type: 'custom',
        def: {
          av: [],
          gd: [],
          paths: [
            {
              c: [
                ['M', 'noSuchName', 0],
                ['L', 1, 1],
              ],
            },
          ],
        },
      },
    });
    expect(renderShape(ANCHOR, bad, BIG).kind).toBe('fallback');
  });
});

describe('TD-92-6 layoutDrawings・可視判定・drawingExtent：外接矩形とアンカーの矩形を合わせた矩形', () => {
  it('形のある図形は box.rect に外接矩形（shape）を持つ。見えない図形は出ない。代替はアンカーの矩形で shape を持たない', () => {
    const boxes = layoutDrawings(geo, [
      shape({ name: 'visible' }),
      shape({ name: 'invisible', hasOutline: false, hasFill: false }),
      shape({ name: 'fallback', geom: { type: 'preset', prst: 'noSuchShape', adj: [] } }),
      shape({ name: 'nogeom', geom: undefined }),
    ]);
    expect(boxes.map((b) => b.index)).toEqual([0, 2, 3]);
    expect(boxes[0]!.shape?.kind).toBe('drawn');
    expect(boxes[0]!.rect).toEqual({ x: 8, y: 18, w: 104, h: 54 });
    expect(boxes[1]!.shape).toBeUndefined();
    expect(boxes[1]!.rect).toEqual(ANCHOR);
    expect(boxes[2]!.shape).toBeUndefined();
  });

  it('大きさが 0 以下の図形は除く（今までどおり）', () => {
    const zero = shape({
      anchor: {
        type: 'oneCell',
        from: { col: 0, colOff: 0, row: 0, rowOff: 0 },
        cx: 0,
        cy: px(10),
      },
    });
    expect(layoutDrawings(geo, [zero])).toEqual([]);
  });

  it('吹き出しの尾：矩形の外へ出る部分も可視判定に使う（アンカーが画面の外でも尾が入れば出る）', () => {
    const d = shape({
      geom: {
        type: 'preset',
        prst: 'wedgeRectCallout',
        adj: [
          ['adj1', -20833],
          ['adj2', 300000],
        ],
      },
    });
    const [b] = layoutDrawings(geo, [d]);
    expect(b!.rect.y + b!.rect.h).toBeGreaterThan(ANCHOR.y + ANCHOR.h + 50);
    // アンカーの下端より 100px 下の窓は、アンカーの矩形とは交わらないが、外接矩形とは交わる
    const win: PxRect = { x: 0, y: ANCHOR.y + ANCHOR.h + 100, w: 500, h: 20 };
    expect(pickVisible([b!], win)).toHaveLength(1);
  });

  it('drawingExtent：外接矩形（尾の先）を含む行まで表を広げる', () => {
    const d = shape({
      geom: {
        type: 'preset',
        prst: 'wedgeRectCallout',
        adj: [
          ['adj1', -20833],
          ['adj2', 300000],
        ],
      },
    });
    const boxes = layoutDrawings(geo, [d]);
    const ext = drawingExtent(geo, boxes);
    const anchorOnly = drawingExtent(geo, [{ index: 0, rect: ANCHOR }]);
    expect(ext.rows).toBeGreaterThan(anchorOnly.rows);
  });

  it('グループの子の図形：矩形は xfrm のまま（縦横を入れ替えない）', () => {
    const d = shape({
      rot: 90 * DEG,
      group: {
        levels: [{ ch: { x: 0, y: 0, cx: px(200), cy: px(100) }, own: null }],
        xfrm: { x: px(20), y: px(10), cx: px(100), cy: px(40) },
      },
      anchor: {
        type: 'oneCell',
        from: { col: 0, colOff: 0, row: 0, rowOff: 0 },
        cx: px(200),
        cy: px(100),
      },
    });
    const [b] = layoutDrawings(geo, [d]);
    if (b?.shape?.kind !== 'drawn') throw new Error('drawn でない');
    expect(b.shape.box).toEqual({ x: 20, y: 10, w: 100, h: 40 });
  });
});
