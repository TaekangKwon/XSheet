import { describe, expect, it } from 'vitest';
import {
  CHART_THIN,
  bandGap,
  clusteredBarColumnPx,
  mergeBarColumns,
  simplifyDP,
  simplifyIndices,
  splitMarkerBands,
  thinM4,
  thinMarkerIndices,
  thinPolyline,
  thinSectorRuns,
  type BarBoxIn,
  type Pt,
  type SectorSpan,
} from '../../src/render/chart-thin';

/** 点 p から線分 ab までの距離（試験用の独立した実装） */
function distSeg(p: Pt, a: Pt, b: Pt): number {
  const dx = b[0] - a[0];
  const dy = b[1] - a[1];
  const l2 = dx * dx + dy * dy;
  const t = l2 === 0 ? 0 : Math.max(0, Math.min(1, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / l2));
  return Math.hypot(p[0] - a[0] - t * dx, p[1] - a[1] - t * dy);
}
/** 点 p から折れ線までの距離 */
function distPoly(p: Pt, poly: readonly Pt[]): number {
  let best = Infinity;
  for (let i = 0; i + 1 < poly.length; i++)
    best = Math.min(best, distSeg(p, poly[i]!, poly[i + 1]!));
  return best;
}
/** 決まった乱数（試験の再現のため） */
function rng(seed: number): () => number {
  let s = seed >>> 0;
  return () => {
    s = (Math.imul(s, 1664525) + 1013904223) >>> 0;
    return s / 2 ** 32;
  };
}

describe('TD-106-1 定数', () => {
  it('ADR-0031 §決定 11.1 の値', () => {
    expect(CHART_THIN).toEqual({
      denseSpacingPx: 1,
      gridPx: 0.5,
      epsPx: 0.25,
      markerSpacingPx: 1,
      barColumnPx: 1,
      barColumnPerSeriesPx: 0.2,
      barColumnMaxPx: 2,
      smoothSteps: 4,
    });
  });
});

describe('TD-106-2 M4（列ごとの最初・最小・最大・最後）', () => {
  it('点が 3 未満なら入力のまま', () => {
    expect(thinM4([], 0)).toEqual([]);
    expect(thinM4([[1, 2]], 0)).toEqual([[1, 2]]);
    expect(
      thinM4(
        [
          [1, 2],
          [1.1, 9],
        ],
        0,
      ),
    ).toEqual([
      [1, 2],
      [1.1, 9],
    ]);
  });

  it('同じ列に続く点は最初・最小・最大・最後だけを元の順で残す', () => {
    // 列 0（x 0〜0.5 未満）に 7 点、列 1 に 2 点
    const pts: Pt[] = [
      [0.0, 5],
      [0.05, 7],
      [0.1, 1],
      [0.2, 4],
      [0.3, 9],
      [0.35, 3],
      [0.4, 6],
      [0.6, 2],
      [0.7, 8],
    ];
    expect(thinM4(pts, 0)).toEqual([
      [0.0, 5], // 最初
      [0.1, 1], // 最小
      [0.3, 9], // 最大
      [0.4, 6], // 最後
      [0.6, 2],
      [0.7, 8],
    ]);
  });

  it('最初・最小・最大・最後が重なる時は重複しない', () => {
    const pts: Pt[] = [
      [0.0, 1],
      [0.1, 5],
      [0.2, 3],
      [0.3, 4],
      [0.4, 9],
      [0.5, 0],
    ];
    // 列 0＝先頭 5 点（最初と最小は同じ点・最大と最後は同じ点）、列 1＝[0.5, 0] の 1 点
    expect(thinM4(pts, 0)).toEqual([
      [0.0, 1],
      [0.4, 9],
      [0.5, 0],
    ]);
  });

  it('項目の並びが逆順でも列ごとにまとめる', () => {
    const pts: Pt[] = [];
    for (let i = 0; i < 100; i++) pts.push([100 - i * 0.1, Math.sin(i)]);
    const out = thinM4(pts, 0);
    expect(out[0]).toEqual(pts[0]);
    expect(out[out.length - 1]).toEqual(pts[99]);
    expect(out.length).toBeLessThan(pts.length);
    // 列ごとに、元の点の最小・最大の y が残っている
    const col = (p: Pt): number => Math.floor(p[0] / 0.5);
    const range = (ps: readonly Pt[]): Map<number, [number, number]> => {
      const m = new Map<number, [number, number]>();
      for (const p of ps) {
        const r = m.get(col(p));
        if (r === undefined) m.set(col(p), [p[1], p[1]]);
        else m.set(col(p), [Math.min(r[0], p[1]), Math.max(r[1], p[1])]);
      }
      return m;
    };
    expect(range(out)).toEqual(range(pts));
  });

  it('縦方向（項目軸が y）でも同じ', () => {
    const pts: Pt[] = [
      [3, 0.0],
      [9, 0.1],
      [1, 0.2],
      [5, 0.3],
      [4, 0.4],
    ];
    expect(thinM4(pts, 1)).toEqual([
      [3, 0.0],
      [9, 0.1],
      [1, 0.2],
      [4, 0.4],
    ]);
  });
});

describe('TD-106-3 Douglas–Peucker', () => {
  it('点が 3 未満なら入力のまま', () => {
    expect(simplifyDP([])).toEqual([]);
    expect(
      simplifyDP([
        [0, 0],
        [5, 5],
      ]),
    ).toEqual([
      [0, 0],
      [5, 5],
    ]);
    expect(
      simplifyIndices([
        [0, 0],
        [1, 1],
      ]),
    ).toEqual([0, 1]);
  });

  it('直線上の点は端点だけになる', () => {
    const pts: Pt[] = Array.from({ length: 50 }, (_, i) => [i, 2 * i]);
    expect(simplifyDP(pts)).toEqual([pts[0], pts[49]]);
  });

  it('端点を残し、元の全点が単純化した線から 0.25px 以内', () => {
    const r = rng(7);
    const pts: Pt[] = Array.from({ length: 3000 }, (_, i) => [
      i * 0.2,
      50 + 30 * Math.sin(i / 40) + r() * 2,
    ]);
    const out = simplifyDP(pts);
    expect(out[0]).toEqual(pts[0]);
    expect(out[out.length - 1]).toEqual(pts[pts.length - 1]);
    expect(out.length).toBeLessThan(pts.length);
    for (const p of pts) expect(distPoly(p, out)).toBeLessThanOrEqual(CHART_THIN.epsPx + 1e-9);
  });

  it('最悪の入力（10 万点の鋸歯）でも固まらず、ずれの上限を守る（全点が残る）', () => {
    const pts: Pt[] = Array.from({ length: 100_000 }, (_, i) => [i * 0.01, i % 2 === 0 ? 0 : 10]);
    const t0 = performance.now();
    const out = simplifyDP(pts);
    expect(performance.now() - t0).toBeLessThan(5000);
    expect(out.length).toBe(pts.length); // 振幅 10px は 0.25px を超える＝全部残る
  });

  it('thinPolyline：M4 の後に単純化する・最初と最後は残る', () => {
    const pts: Pt[] = Array.from({ length: 4000 }, (_, i) => [i * 0.1, 50 + 30 * Math.sin(i / 50)]);
    const out = thinPolyline(pts, 0);
    expect(out[0]).toEqual(pts[0]);
    expect(out[out.length - 1]).toEqual(pts[3999]);
    expect(out.length).toBeLessThan(200);
    for (const p of pts) expect(distPoly(p, out)).toBeLessThanOrEqual(0.5);
  });
});

describe('TD-106-4 円のマーカーの帯の区間分け', () => {
  it('間隔の上限は √(8R·eps)', () => {
    expect(bandGap(4.667)).toBeCloseTo(Math.sqrt(8 * 4.667 * 0.25), 9);
    expect(bandGap(0)).toBe(0);
  });

  it('続く中心の間隔が上限以下の区間（2 点以上）が帯・他は単独', () => {
    const c: Pt[] = [
      [0, 0],
      [1, 0],
      [2, 0], // 帯 1（3 点）
      [10, 0], // 単独
      [20, 0],
      [21, 1], // 帯 2（2 点）
      [40, 0], // 単独
    ];
    const { runs, singles } = splitMarkerBands(c, 3);
    expect(runs).toEqual([
      [c[0], c[1], c[2]],
      [c[4], c[5]],
    ]);
    expect(singles).toEqual([c[3], c[6]]);
  });

  it('点が無い・1 点', () => {
    expect(splitMarkerBands([], 3)).toEqual({ runs: [], singles: [] });
    expect(splitMarkerBands([[1, 1]], 3)).toEqual({ runs: [], singles: [[1, 1]] });
  });
});

describe('TD-106-5 その他のマーカーの間引きの添字', () => {
  it('点が 3 未満なら全部', () => {
    expect(
      thinMarkerIndices([
        [0, 0],
        [0.1, 0],
      ]),
    ).toEqual([0, 1]);
  });

  it('省いた点は残した点から 1px 以内・Douglas–Peucker の点は残る・端が残る', () => {
    const pts: Pt[] = Array.from({ length: 2000 }, (_, i) => [i * 0.3, 50 + 30 * Math.sin(i / 60)]);
    const keep = thinMarkerIndices(pts);
    expect(keep[0]).toBe(0);
    for (const i of simplifyIndices(pts, CHART_THIN.epsPx)) expect(keep).toContain(i);
    expect(keep.length).toBeLessThan(pts.length / 2);
    const set = new Set(keep);
    for (let i = 0; i < pts.length; i++) {
      if (set.has(i)) continue;
      const near = Math.min(
        ...keep.map((k) => Math.hypot(pts[i]![0] - pts[k]![0], pts[i]![1] - pts[k]![1])),
      );
      expect(near).toBeLessThanOrEqual(CHART_THIN.markerSpacingPx + 1e-9);
    }
  });
});

describe('TD-106-6 棒の列ごとのまとめ', () => {
  const bar = (lo: number, hi: number, a: number, b: number, neg = false): BarBoxIn => ({
    lo,
    hi,
    a,
    b,
    neg,
  });

  it('空は空・1 本はほぼそのまま（列の中に収まる）', () => {
    expect(mergeBarColumns([])).toEqual([]);
    const [one] = mergeBarColumns([bar(0.2, 0.5, 10, 30)]);
    expect(one!.lo).toBeCloseTo(0.2, 9);
    expect(one!.hi).toBeCloseTo(0.5, 9);
    expect([one!.a, one!.b]).toEqual([10, 30]);
  });

  it('同じ 1px の列の棒は 1 本になり、幅は合計・値の範囲は和集合・位置は列の中', () => {
    const out = mergeBarColumns([
      bar(0.0, 0.2, 100, 120),
      bar(0.3, 0.5, 90, 110),
      bar(0.6, 0.8, 100, 140),
      bar(1.1, 1.3, 50, 60), // 次の列
    ]);
    expect(out).toHaveLength(2);
    const [c0, c1] = out;
    expect(c0!.hi - c0!.lo).toBeCloseTo(0.6, 9);
    expect(c0!.a).toBe(90);
    expect(c0!.b).toBe(140);
    expect(c0!.lo).toBeGreaterThanOrEqual(0);
    expect(c0!.hi).toBeLessThanOrEqual(1 + 1e-9);
    expect(c1).toEqual({ lo: 1.1, hi: 1.3, a: 50, b: 60 });
  });

  it('正の側と負の側は別にまとめる', () => {
    const out = mergeBarColumns([bar(0.1, 0.3, 100, 140), bar(0.5, 0.7, 140, 160, true)]);
    expect(out).toHaveLength(2);
    expect(out.map((b) => [b.a, b.b])).toEqual([
      [100, 140],
      [140, 160],
    ]);
  });

  it('列ごとの幅の合計が保たれ、元の全部の棒の値の範囲が入る', () => {
    const r = rng(3);
    const bars: BarBoxIn[] = [];
    for (let i = 0; i < 1000; i++) {
      const lo = i * 0.4;
      const v = r() * 100;
      bars.push(bar(lo, lo + 0.3, 200 - v, 200));
    }
    const out = mergeBarColumns(bars);
    const total = (xs: readonly { lo: number; hi: number }[]): number =>
      xs.reduce((s, b) => s + (b.hi - b.lo), 0);
    expect(total(out)).toBeCloseTo(total(bars), 6);
    expect(out.length).toBeLessThanOrEqual(Math.ceil(1000 * 0.4) + 1);
    for (const b of bars) {
      const col = Math.floor((b.lo + b.hi) / 2);
      const m = out.find((o) => Math.floor((o.lo + o.hi) / 2 + 1e-9) === col)!;
      expect(m.a).toBeLessThanOrEqual(b.a);
      expect(m.b).toBeGreaterThanOrEqual(b.b);
    }
  });
});

describe('TD-111-1 集合の棒の列の幅（系列の数×0.2px を整数 px に丸める・1〜2px）', () => {
  it('系列 1・7 は 1px・8・10 は 2px・上限 2px・整数 px', () => {
    expect(clusteredBarColumnPx(1)).toBe(1);
    expect(clusteredBarColumnPx(7)).toBe(1); // 1.4 → 1
    expect(clusteredBarColumnPx(8)).toBe(2); // 1.6 → 2
    expect(clusteredBarColumnPx(10)).toBe(2);
    expect(clusteredBarColumnPx(100)).toBe(2);
    expect(clusteredBarColumnPx(0)).toBe(1);
  });

  it('列の幅を広げると矩形の数が減り、幅の合計は保たれる', () => {
    const bars: BarBoxIn[] = [];
    for (let i = 0; i < 1000; i++)
      bars.push({ lo: i * 0.4, hi: i * 0.4 + 0.3, a: 100 - (i % 7), b: 100, neg: false });
    const w1 = mergeBarColumns(bars, 1);
    const w2 = mergeBarColumns(bars, clusteredBarColumnPx(10));
    expect(w2.length).toBeLessThan(w1.length * 0.6);
    const total = (xs: readonly { lo: number; hi: number }[]): number =>
      xs.reduce((s, b) => s + (b.hi - b.lo), 0);
    expect(total(w2)).toBeCloseTo(total(bars), 6);
    expect(Math.min(...w2.map((b) => b.a))).toBe(94);
  });
});

describe('TD-111-2 円・ドーナツの扇の束ね（外周の弧が線の太さ以下の連なり）', () => {
  /** 外半径 100 で、弧の長さ（px）が arcs の順に並ぶ扇 */
  const spansOf = (arcs: number[], fixed: number[] = []): SectorSpan[] => {
    let a = 0;
    return arcs.map((len, i) => {
      const a0 = a;
      a += len / 100;
      return { a0, a1: a, fixed: fixed.includes(i) };
    });
  };

  it('弧の長さ＝線の太さは束ねる・わずかに超えたら束ねない（境目）', () => {
    expect(thinSectorRuns(spansOf([2, 2, 2]), 100, 2)).toEqual([[0, 2]]);
    expect(thinSectorRuns(spansOf([2.01, 2.01]), 100, 2)).toEqual([]);
    expect(thinSectorRuns(spansOf([2, 2.01, 2]), 100, 2)).toEqual([]);
  });

  it('1 つだけの細い扇は束ねない・太い扇で区間が切れる', () => {
    expect(thinSectorRuns(spansOf([50, 1, 50]), 100, 2)).toEqual([]);
    expect(thinSectorRuns(spansOf([1, 1, 50, 1, 1, 1, 50, 1]), 100, 2)).toEqual([
      [0, 1],
      [3, 5],
    ]);
  });

  it('fixed（c:dPt で上書きした扇）は束ねず、区間を切る', () => {
    expect(thinSectorRuns(spansOf([1, 1, 1, 1, 1], [2]), 100, 2)).toEqual([
      [0, 1],
      [3, 4],
    ]);
    expect(thinSectorRuns(spansOf([1, 1, 1], [1]), 100, 2)).toEqual([]);
  });

  it('半径が小さい（ドーナツの内側の輪）ほど束ねやすい・線が無い（太さ 0）・半径 0 は束ねない', () => {
    const spans = spansOf([2.5, 2.5, 2.5]); // 外半径 100 の弧。半径 50 なら弧は半分
    expect(thinSectorRuns(spans, 100, 2)).toEqual([]);
    expect(thinSectorRuns(spans, 50, 2)).toEqual([[0, 2]]);
    expect(thinSectorRuns(spans, 100, 0)).toEqual([]);
    expect(thinSectorRuns(spans, 0, 2)).toEqual([]);
    expect(thinSectorRuns([], 100, 2)).toEqual([]);
  });
});
