import { describe, expect, it } from 'vitest';
import type { SheetDrawing } from '../../src/core/index';
import type { DrawingText } from '../../src/core/types';
import type { PxRect } from '../../src/render/drawings';
import { EMU_PER_PX } from '../../src/render/drawings';
import { BACKDROP_PAD, renderShape } from '../../src/render/shapes';

/**
 * ダーク表示の下地の形（純ロジック。ISSUE-96・ADR-0028 §決定 6.7）：
 * 塗りのある図形・塗りが無く文字のある図形は形の塗りの path、枠線は線の path、塗りも文字も無い図形は枠線だけ、塗りの path が無い形は回転前の矩形。
 */

const BIG = { totalWidth: 1_000_000, totalHeight: 1_000_000 };
const px = (n: number): number => n * EMU_PER_PX;
const ANCHOR: PxRect = { x: 10, y: 20, w: 100, h: 50 };

const TEXT = {
  anchor: 't',
  vert: 'horz',
  noWrap: false,
  paragraphs: [{ align: 'l', bullet: null, runs: [{ text: 'あ' }] }],
} as unknown as DrawingText;

function shape(over: Partial<SheetDrawing> = {}): SheetDrawing {
  return {
    anchor: { type: 'absolute', x: px(10), y: px(20), cx: px(100), cy: px(50) },
    kind: 'shape',
    name: 's',
    descr: '',
    mediaPath: null,
    geom: { type: 'preset', prst: 'rect', adj: [] },
    hasOutline: false,
    hasFill: false,
    ...over,
  };
}

function drawn(d: SheetDrawing, rect: PxRect = ANCHOR) {
  const r = renderShape(rect, d, BIG);
  if (r.kind !== 'drawn') throw new Error('drawn でない');
  return r;
}

describe('TD-96-1 下地の形（renderShape の backdrop）', () => {
  it('下地の余白は片側 3px', () => {
    expect(BACKDROP_PAD).toBe(3);
  });

  it('塗りのある図形：形の塗りの path を白で塗る（面と同じ d）。枠線が無ければ縁取りは無い', () => {
    const r = drawn(shape({ hasFill: true }));
    expect(r.backdrop).toEqual([{ d: r.faces[0]!.d, fill: true, stroke: false }]);
  });

  it('塗り＋枠線：同じ path が塗りと縁取りを兼ねる（1 つ）', () => {
    const r = drawn(shape({ hasFill: true, hasOutline: true }));
    expect(r.backdrop).toEqual([{ d: r.faces[0]!.d, fill: true, stroke: true }]);
  });

  it('塗りも文字も無い図形：枠線の縁取りだけ（内側は塗らない）', () => {
    const r = drawn(shape({ hasOutline: true }));
    expect(r.backdrop).toEqual([{ d: r.faces[0]!.d, fill: false, stroke: true }]);
  });

  it('塗りが無く文字のある図形：形の塗りの path を白で塗る（枠線が無くても。faces は空のまま）', () => {
    const r = drawn(shape({ text: TEXT, geom: { type: 'preset', prst: 'ellipse', adj: [] } }));
    expect(r.faces).toEqual([]);
    expect(r.backdrop).toHaveLength(1);
    expect(r.backdrop[0]).toMatchObject({ fill: true, stroke: false });
    expect(r.backdrop[0]!.d).toMatch(/C/); // 楕円は曲線
  });

  it('塗りが無く文字と枠線のある図形：塗りの path は塗り、枠線は縁取り', () => {
    const r = drawn(shape({ text: TEXT, hasOutline: true }));
    expect(r.backdrop).toEqual([{ d: r.faces[0]!.d, fill: true, stroke: true }]);
  });

  it('円柱（塗りあり・枠線あり）：塗る path（濃淡の path も）は白で塗り、fill が none で線だけの path は縁取りだけ', () => {
    const r = drawn(
      shape({ hasFill: true, hasOutline: true, geom: { type: 'preset', prst: 'can', adj: [] } }),
    );
    expect(r.backdrop.map((b) => [b.fill, b.stroke])).toEqual([
      [true, false],
      [true, false],
      [false, true],
    ]);
  });

  it('塗りの path が無い形（全 path が fill none）で文字がある：回転前の矩形を回した形で塗る', () => {
    const def = {
      type: 'custom' as const,
      def: {
        av: [],
        gd: [],
        paths: [
          {
            w: 10,
            h: 10,
            fill: 'none' as const,
            c: [
              ['M', 0, 0],
              ['L', 10, 10],
            ] as const,
          },
        ],
      },
    };
    const r = drawn(shape({ geom: def, text: TEXT, rot: 90 * 60000 }), {
      x: 35,
      y: -5 + 25,
      w: 50,
      h: 100,
    });
    expect(r.backdrop).toHaveLength(1);
    expect(r.backdrop[0]).toMatchObject({ fill: true, stroke: false });
    // 回転前 100×50 の矩形（中心 (60,70)）を 90° 回した 4 点：x は 35〜85、y は 20〜120
    expect(r.backdrop[0]!.d).toMatch(/^M[\d. -]+ L[\d. -]+ L[\d. -]+ L[\d. -]+ Z$/);
  });

  it('塗りの path が無く文字も無い図形（線だけの形）：塗りは下地にならず、枠線は縁取りだけ', () => {
    const def = {
      type: 'custom' as const,
      def: {
        av: [],
        gd: [],
        paths: [
          {
            w: 10,
            h: 10,
            fill: 'none' as const,
            c: [
              ['M', 0, 0],
              ['L', 10, 10],
            ] as const,
          },
        ],
      },
    };
    // 塗りがあっても fill none の path しか無ければ塗るものが無い（面も下地も空）
    const filledOnly = drawn(shape({ geom: def, hasFill: true }));
    expect(filledOnly.faces).toEqual([]);
    expect(filledOnly.backdrop).toEqual([]);
    const outlined = drawn(shape({ geom: def, hasOutline: true }));
    expect(outlined.backdrop).toEqual([{ d: outlined.faces[0]!.d, fill: false, stroke: true }]);
  });

  it('回転・反転は下地の path にも掛かる（面の d と同じ座標）', () => {
    const r = drawn(
      shape({
        hasFill: true,
        rot: 45 * 60000,
        flipH: true,
        geom: { type: 'preset', prst: 'rightArrow', adj: [] },
      }),
    );
    expect(r.backdrop[0]!.d).toBe(r.faces[0]!.d);
  });

  it('見えない塗り・見えない文字は下地の塗りにしない（不透明度 0）。枠線の縁取りは今のまま', () => {
    const zero = { rgb: 0xff0000, alpha: 0 };
    const hiddenText = {
      ...TEXT,
      paragraphs: [{ align: 'l', bullet: null, runs: [{ text: 'あ', color: zero }] }],
    } as unknown as DrawingText;
    expect(drawn(shape({ hasFill: true, fillColor: zero })).backdrop).toEqual([]);
    expect(drawn(shape({ text: hiddenText })).backdrop).toEqual([]);
    const outlined = drawn(
      shape({ hasFill: true, fillColor: zero, text: hiddenText, hasOutline: true }),
    );
    expect(outlined.backdrop).toEqual([{ d: outlined.faces[0]!.d, fill: false, stroke: true }]);
    // 見える塗り（不透明度 0 超）・見える文字は塗る
    expect(
      drawn(shape({ hasFill: true, fillColor: { rgb: 1, alpha: 0.1 } })).backdrop[0],
    ).toMatchObject({ fill: true });
    expect(drawn(shape({ text: TEXT })).backdrop[0]).toMatchObject({ fill: true });
  });
});
