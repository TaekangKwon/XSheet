import { describe, expect, it } from 'vitest';
import { applyTransforms } from '../../src/core/color';
import { DEFAULT_THEME } from '../../src/core/theme';
import type {
  ChartAxis,
  ChartLine,
  ChartModel,
  ChartPaint,
  ChartSeries,
  ManualLayout,
} from '../../src/core/types';
import {
  CHART_COLOR_VARIATIONS,
  CHART_DEFAULTS,
  MAX_CHART_SCENE_ITEMS,
  PX_PER_PT,
  autoColor,
  estimateTextWidth,
  layoutChart,
  planCategoryLabelForm,
  truncateLine,
  wrapCategoryLabel,
  type ChartItem,
  type ChartPathItem,
  type ChartRectItem,
  type ChartScene,
  type ChartTextItem,
} from '../../src/render/chart-layout';

// ---- 試験の補助 ----

function ser(
  idx: number,
  name: string,
  values: (number | null)[],
  over: Partial<ChartSeries> = {},
): ChartSeries {
  return { idx, name, values, ...over };
}

const catAx: ChartAxis = { kind: 'cat', deleted: false, orientation: 'minMax', text: {} };
const valAx: ChartAxis = {
  kind: 'val',
  deleted: false,
  orientation: 'minMax',
  text: {},
  majorGridlines: {},
};

/** 既定は 3 項目・1 系列の縦棒 */
function mk(over: Partial<ChartModel> = {}): ChartModel {
  return {
    type: 'bar',
    barDir: 'col',
    grouping: 'clustered',
    varyColors: false,
    categories: ['A', 'B', 'C'],
    series: [ser(0, 'S1', [10, 20, 30])],
    catAxis: catAx,
    valAxis: valAx,
    autoTitleDeleted: true,
    chartArea: {},
    plotArea: {},
    text: {},
    dispBlanksAs: 'gap',
    ...over,
  };
}

const pie = (over: Partial<ChartModel> = {}): ChartModel =>
  mk({
    type: 'pie',
    barDir: undefined,
    grouping: 'standard',
    varyColors: true,
    catAxis: undefined,
    valAxis: undefined,
    ...over,
  });

/** 場面を作る（null なら試験を落とす） */
function scene(m: ChartModel, w = 400, h = 300): ChartScene {
  const s = layoutChart(m, w, h);
  if (s === null) throw new Error('場面が null');
  return s;
}

const paths = (s: ChartScene): ChartPathItem[] =>
  s.items.filter((i): i is ChartPathItem => i.kind === 'path');
const texts = (s: ChartScene): ChartTextItem[] =>
  s.items.filter((i): i is ChartTextItem => i.kind === 'text');
const rects = (s: ChartScene): ChartRectItem[] =>
  s.items.filter((i): i is ChartRectItem => i.kind === 'rect');

/** 塗りのある経路（系列の図形）：塗りの色で絞る */
const fillPaths = (s: ChartScene): ChartPathItem[] => paths(s).filter((p) => p.fill !== null);

/** `M x y L x y L x y L x y Z` の連なりを矩形（x・y・w・h）にする */
function rectsOf(d: string): { x: number; y: number; w: number; h: number }[] {
  const out: { x: number; y: number; w: number; h: number }[] = [];
  for (const part of d.split('Z').filter((p) => p !== '')) {
    const n = part.match(/-?\d+(?:\.\d+)?/g)!.map(Number);
    const xs = [n[0]!, n[2]!, n[4]!, n[6]!];
    const ys = [n[1]!, n[3]!, n[5]!, n[7]!];
    out.push({
      x: Math.min(...xs),
      y: Math.min(...ys),
      w: Math.max(...xs) - Math.min(...xs),
      h: Math.max(...ys) - Math.min(...ys),
    });
  }
  return out;
}

/** 手動のプロットを内側・edge で与え、プロットの面の矩形を読めるようにする */
const innerPlot = (x: number, y: number, w: number, h: number): Partial<ChartModel> => ({
  plotLayout: {
    target: 'inner',
    xMode: 'edge',
    yMode: 'edge',
    wMode: 'factor',
    hMode: 'factor',
    x,
    y,
    w,
    h,
  },
  plotArea: { fill: { kind: 'solid', color: { rgb: 0x123456, alpha: 1 } } },
});
const plotRect = (s: ChartScene): ChartRectItem =>
  rects(s).find((r) => r.fill?.color === '#123456')!;

const solid = (
  rgb: number,
  alpha = 1,
): { kind: 'solid'; color: { rgb: number; alpha: number } } => ({
  kind: 'solid',
  color: { rgb, alpha },
});
const hexOf = (rgb: number): string => `#${rgb.toString(16).padStart(6, '0')}`;

// ---- 文字の幅・切り詰め ----

describe('文字の幅の推定・切り詰め', () => {
  it('TD-102-20: 半角英数 0.55em・半角カナ 0.5em・全角 1em・その他 0.6em', () => {
    expect(estimateTextWidth('ab', 10)).toBeCloseTo(11, 10);
    expect(estimateTextWidth('ｱｲ', 10)).toBeCloseTo(10, 10);
    expect(estimateTextWidth('あ漢', 10)).toBeCloseTo(20, 10);
    expect(estimateTextWidth('é', 10)).toBeCloseTo(6, 10);
    expect(estimateTextWidth('', 10)).toBe(0);
  });

  it('TD-102-20: 1 行は 255 文字で切る（超えたら末尾 …）', () => {
    expect(truncateLine('a'.repeat(255))).toBe('a'.repeat(255));
    const t = truncateLine('b'.repeat(300));
    expect(Array.from(t)).toHaveLength(255);
    expect(t.endsWith('…')).toBe(true);
  });
});

// ---- 既定の色 ----

describe('既定の色（ADR-0031 §決定 4・C3）', () => {
  it('TD-102-21: 系列 1〜6 は accent1〜6', () => {
    const want = [0x4472c4, 0xed7d31, 0xa5a5a5, 0xffc000, 0x5b9bd5, 0x70ad47];
    want.forEach((rgb, i) => expect(autoColor(DEFAULT_THEME, i).rgb).toBe(rgb));
  });

  it('TD-102-21: 7 系列目は accent1 の lumMod 60%・13 系列目は lumMod 80%+lumOff 20%（変化の表どおり）', () => {
    const a1 = { rgb: 0x4472c4, alpha: 1 };
    expect(autoColor(DEFAULT_THEME, 6)).toEqual(applyTransforms(a1, '<a:lumMod val="60000"/>'));
    expect(autoColor(DEFAULT_THEME, 12)).toEqual(
      applyTransforms(a1, '<a:lumMod val="80000"/><a:lumOff val="20000"/>'),
    );
    expect(autoColor(DEFAULT_THEME, 6).rgb).not.toBe(0x4472c4);
    expect(autoColor(DEFAULT_THEME, 7).rgb).not.toBe(autoColor(DEFAULT_THEME, 1).rgb); // accent2 の変化
  });

  it('TD-102-21: 変化は 9 周で繰り返す', () => {
    expect(CHART_COLOR_VARIATIONS).toHaveLength(9);
    expect(autoColor(DEFAULT_THEME, 54)).toEqual(autoColor(DEFAULT_THEME, 0));
    expect(autoColor(DEFAULT_THEME, 60)).toEqual(autoColor(DEFAULT_THEME, 6));
  });

  it('TD-102-21: テーマの色が当たる（accent1 を変えた時）', () => {
    const theme = new Map(DEFAULT_THEME);
    theme.set('accent1', 0x112233);
    expect(autoColor(theme, 0).rgb).toBe(0x112233);
  });

  it('TD-102-22: 既定の色は配列の位置でなく ChartSeries.idx で決まる', () => {
    const m = mk({ series: [ser(3, 'X', [1, 2, 3]), ser(0, 'Y', [1, 2, 3])] });
    const fills = fillPaths(scene(m)).map((p) => p.fill!.color);
    expect(fills).toEqual([hexOf(0xffc000), hexOf(0x4472c4)]);
  });

  it('TD-102-23: varyColors が真の円・ドーナツは、点の位置の色で当たる（偽の時は TD-104-33）', () => {
    const s = scene(
      pie({
        series: [ser(0, 'S', [1, 1, 1, 1, 1, 1, 1])],
        categories: ['a', 'b', 'c', 'd', 'e', 'f', 'g'],
      }),
    );
    const f = fillPaths(s).map((p) => p.fill!.color);
    expect(f.slice(0, 6)).toEqual(
      [0x4472c4, 0xed7d31, 0xa5a5a5, 0xffc000, 0x5b9bd5, 0x70ad47].map(hexOf),
    );
    expect(f[6]).toBe(hexOf(autoColor(DEFAULT_THEME, 6).rgb));
  });

  it('TD-102-23: 棒は varyColors でも 1 系列の時だけ点の位置・2 系列なら系列の色', () => {
    const one = fillPaths(scene(mk({ varyColors: true }))).map((p) => p.fill!.color);
    expect(one).toEqual([0x4472c4, 0xed7d31, 0xa5a5a5].map(hexOf));
    const two = fillPaths(
      scene(mk({ varyColors: true, series: [ser(0, 'a', [1, 2, 3]), ser(1, 'b', [1, 2, 3])] })),
    ).map((p) => p.fill!.color);
    expect(two).toEqual([hexOf(0x4472c4), hexOf(0xed7d31)]);
    // varyColors が偽なら 1 系列は 1 色
    expect(fillPaths(scene(mk()))).toHaveLength(1);
  });

  it('TD-102-24: 明示の色が優先（系列の塗り）・c:dPt が点を上書き・noFill は描かない', () => {
    const m = mk({
      series: [
        ser(0, 'S', [10, 20, 30], {
          fill: solid(0xff0000, 0.5),
          points: new Map<number, ChartPaint>([
            [1, solid(0x00ff00)],
            [2, { kind: 'none' as const }],
          ]),
        }),
      ],
    });
    const fp = fillPaths(scene(m));
    expect(fp.map((p) => [p.fill!.color, p.fill!.opacity])).toEqual([
      ['#ff0000', 0.5],
      ['#00ff00', 1],
    ]);
    expect(fp.every((p) => rectsOf(p.d).length === 1)).toBe(true); // 点 2 は none＝経路を作らない
    expect(
      paths(scene(mk({ series: [ser(0, 'S', [1, 2, 3], { fill: { kind: 'none' } })] }))).filter(
        (p) => p.fill !== null,
      ),
    ).toHaveLength(0);
  });

  it('TD-102-25: 部品の既定（グラフの面・目盛線・軸線・文字・円の枠線・折れ線の太さ）', () => {
    const s = scene(mk({ title: { lines: ['T'], text: {}, overlay: false } }));
    const ca = rects(s)[0]!;
    expect([ca.width, ca.height]).toEqual([400, 300]);
    expect(ca.fill).toEqual({ color: '#ffffff', opacity: 1 });
    expect(ca.stroke!.color).toBe('#898989'); // ISSUE-104：c:spPr の無いグラフを Excel 16 で開いた値（0.5pt・#898989）
    expect(ca.stroke!.width).toBeCloseTo(0.5 * PX_PER_PT, 10);
    // ISSUE-105：c:spPr の無い軸は、主の目盛線・項目軸の線・数値軸の線が黒 0.5pt（Excel 16 の PNG の画素）
    const grid = paths(s).find((p) => p.fill === null && p.stroke?.color === '#000000')!;
    expect(grid.stroke!.width).toBeCloseTo(0.5 * PX_PER_PT, 10);
    const axisLines = s.items.filter((i) => i.kind === 'line');
    expect(axisLines).toHaveLength(2); // 項目軸と数値軸
    for (const l of axisLines) {
      expect(l.kind === 'line' && l.stroke.color).toBe('#000000');
      expect(l.kind === 'line' && l.stroke.width).toBeCloseTo(0.5 * PX_PER_PT, 10);
    }
    expect(texts(s).every((t) => t.color.color === '#000000')).toBe(true);
    // 円の枠線は白 1.5pt（Excel が点ごとに書く値）
    const p = paths(scene(pie()))[0]!;
    expect(p.stroke).toMatchObject({ color: '#ffffff', width: 1.5 * PX_PER_PT });
    // 折れ線は 1.5pt・端は丸・系列の色
    const ln = scene(mk({ type: 'line', barDir: undefined, grouping: 'standard' }));
    const lp = paths(ln).find((x) => x.fill === null && x.stroke?.color === '#4472c4')!;
    expect(lp.stroke).toMatchObject({ width: 1.5 * PX_PER_PT, cap: 'round', join: 'round' });
  });

  it('TD-102-25: 明示の線（太さ・色）・目盛線なし・軸線 noFill が効く', () => {
    const line: ChartLine = { paint: solid(0xff00ff), widthEmu: 19050 };
    const s = scene(
      mk({
        valAxis: { ...valAx, majorGridlines: line },
        catAxis: { ...catAx, line: { paint: { kind: 'none' } } },
      }),
    );
    expect(paths(s).some((p) => p.stroke?.color === '#ff00ff' && p.stroke.width === 2)).toBe(true);
    // 項目軸の線は none・数値軸は既定（黒の縦線 1 本）
    const lines = s.items.filter((i) => i.kind === 'line');
    expect(lines).toHaveLength(1);
    expect(lines[0]!.kind === 'line' && lines[0]!.x1 === lines[0]!.x2).toBe(true);
    const noTick = { majorTickMark: 'none', minorTickMark: 'none' } as const; // 目盛の印が無いと既定は交差（C12）
    const none = scene(
      mk({
        valAxis: { ...valAx, ...noTick, majorGridlines: undefined },
        catAxis: { ...catAx, ...noTick },
      }),
    );
    expect(paths(none).some((p) => p.fill === null && p.stroke?.color === '#000000')).toBe(false);
  });

  it('TD-102-26: backdrop（グラフの面が塗りなしで真・白の塗りで偽・既定は偽）', () => {
    expect(scene(mk({ chartArea: { fill: { kind: 'none' } } })).backdrop).toBe(true);
    expect(scene(mk({ chartArea: { fill: solid(0xffffff) } })).backdrop).toBe(false);
    expect(scene(mk()).backdrop).toBe(false);
    const nf = scene(mk({ chartArea: { fill: { kind: 'none' } } }));
    expect(rects(nf)[0]!.fill).toBeNull(); // 面の塗りは描かない
  });
});

// ---- 文字の大きさ ----

describe('文字の大きさ・太字・色（§決定 8）', () => {
  it('TD-102-27: 既定はタイトル 10pt（太字）・軸 10pt・凡例 10pt・軸タイトル 10pt（ISSUE-104：Excel 16 の照合 C4）', () => {
    const m = mk({
      title: { lines: ['T'], text: {}, overlay: false },
      legend: { pos: 'r', overlay: false, deleted: [], text: {} },
      catAxis: { ...catAx, title: { lines: ['CT'], text: {}, overlay: false } },
    });
    const t = texts(scene(m));
    const size = (s: string): number => t.find((x) => x.text === s)!.size;
    expect(size('T')).toBeCloseTo(10 * PX_PER_PT, 10);
    expect(size('A')).toBeCloseTo(10 * PX_PER_PT, 10);
    expect(size('S1')).toBeCloseTo(10 * PX_PER_PT, 10);
    expect(size('CT')).toBeCloseTo(10 * PX_PER_PT, 10);
    expect(t.find((x) => x.text === 'T')!.bold).toBe(true);
  });

  it('TD-102-27: 部品 → グラフ全体 → 既定の順・太字と色を使う', () => {
    const m = mk({
      text: { sizePt: 12, color: { rgb: 0x00ff00, alpha: 1 } },
      catAxis: { ...catAx, text: { sizePt: 8, bold: true, color: { rgb: 0xff0000, alpha: 1 } } },
    });
    const t = texts(scene(m));
    const a = t.find((x) => x.text === 'A')!;
    expect([a.size, a.bold, a.color.color]).toEqual([8 * PX_PER_PT, true, '#ff0000']);
    const v = t.find((x) => x.text === '10')!;
    expect([v.size, v.color.color]).toEqual([12 * PX_PER_PT, '#00ff00']);
  });
});

// ---- 配置 ----

describe('manualLayout（§決定 6）', () => {
  const W = 400;
  const H = 300;

  it('TD-102-30: inner・edge・factor：x・y は左上の位置の割合、w・h は大きさの割合', () => {
    const r = plotRect(scene(mk(innerPlot(0.1, 0.2, 0.5, 0.6)), W, H));
    expect([r.x, r.y, r.width, r.height]).toEqual([40, 60, 200, 180]);
  });

  it('TD-102-30: wMode・hMode が edge なら右下の位置', () => {
    const m = mk({
      plotLayout: {
        target: 'inner',
        xMode: 'edge',
        yMode: 'edge',
        wMode: 'edge',
        hMode: 'edge',
        x: 0.1,
        y: 0.2,
        w: 0.9,
        h: 0.8,
      },
      plotArea: { fill: solid(0x123456) },
    });
    const r = plotRect(scene(m, W, H));
    expect([r.x, r.y, r.width, r.height]).toEqual([40, 60, 320, 180]);
  });

  it('TD-102-30: xMode・yMode が factor（既定）なら自動の位置からのずれ', () => {
    const auto = plotRect(
      scene(
        mk({
          plotLayout: { target: 'inner', w: 0.5, h: 0.5 },
          plotArea: { fill: solid(0x123456) },
        }),
        W,
        H,
      ),
    );
    const shifted = plotRect(
      scene(
        mk({
          plotLayout: { target: 'inner', x: 0.1, y: -0.05, w: 0.5, h: 0.5 },
          plotArea: { fill: solid(0x123456) },
        }),
        W,
        H,
      ),
    );
    expect(shifted.x).toBeCloseTo(auto.x + 40, 5);
    expect(shifted.y).toBeCloseTo(auto.y - 15, 5);
    expect([shifted.width, shifted.height]).toEqual([200, 150]);
  });

  it('TD-102-30: outer（既定）は軸の文字を含む外側＝内側は小さくなる', () => {
    const layout: ManualLayout = { xMode: 'edge', yMode: 'edge', x: 0.1, y: 0.1, w: 0.8, h: 0.8 };
    const r = plotRect(
      scene(mk({ plotLayout: layout, plotArea: { fill: solid(0x123456) } }), W, H),
    );
    expect(r.x).toBeGreaterThan(40);
    expect(r.y).toBeGreaterThanOrEqual(30);
    expect(r.x + r.width).toBeLessThanOrEqual(360 + 1e-9);
    expect(r.y + r.height).toBeLessThan(270);
    const inner = plotRect(
      scene(
        mk({ plotLayout: { ...layout, target: 'inner' }, plotArea: { fill: solid(0x123456) } }),
        W,
        H,
      ),
    );
    expect([inner.x, inner.y, inner.width, inner.height]).toEqual([40, 30, 320, 240]);
    expect(r.width).toBeLessThan(inner.width);
  });

  it('TD-102-31: タイトル・凡例の manualLayout（edge の x・y）', () => {
    const t = texts(
      scene(
        mk({
          title: {
            lines: ['T'],
            text: {},
            overlay: false,
            layout: { xMode: 'edge', yMode: 'edge', x: 0.5, y: 0.5 },
          },
        }),
        W,
        H,
      ),
    ).find((x) => x.text === 'T')!;
    expect(t.y).toBeGreaterThan(150);
    const lg = texts(
      scene(
        mk({
          legend: {
            pos: 'r',
            overlay: false,
            deleted: [],
            text: {},
            layout: { xMode: 'edge', yMode: 'edge', x: 0.1, y: 0.1 },
          },
        }),
        W,
        H,
      ),
    ).find((x) => x.text === 'S1')!;
    expect(lg.x).toBeGreaterThan(40);
    expect(lg.x).toBeLessThan(150);
    expect(lg.y).toBeLessThan(80);
  });
});

describe('自動配置（タイトル・凡例・プロット）', () => {
  const base = (m: Partial<ChartModel>): ChartRectItem =>
    plotRect(scene(mk({ plotArea: { fill: solid(0x123456) }, ...m })));
  const legend = (
    pos: 'r' | 'l' | 't' | 'b' | 'tr',
    overlay = false,
  ): NonNullable<ChartModel['legend']> => ({ pos, overlay, deleted: [], text: {} });

  it('TD-102-32: 縁の余白は 7pt（プロットは余白の内側）', () => {
    const r = base({});
    expect(r.x).toBeGreaterThan(7 * PX_PER_PT);
    expect(r.y).toBeGreaterThanOrEqual(7 * PX_PER_PT - 1e-9);
    expect(r.x + r.width).toBeLessThanOrEqual(400 - 7 * PX_PER_PT + 1e-9);
    expect(r.y + r.height).toBeLessThan(300 - 7 * PX_PER_PT);
  });

  it('TD-102-33: タイトルがあるとプロットが下がる・overlay なら下がらない', () => {
    const none = base({});
    const withT = base({ title: { lines: ['T'], text: {}, overlay: false } });
    const over = base({ title: { lines: ['T'], text: {}, overlay: true } });
    expect(withT.y).toBeGreaterThan(none.y);
    expect(withT.height).toBeLessThan(none.height);
    expect(over.y).toBe(none.y);
    expect(over.height).toBe(none.height);
  });

  it('TD-102-33: 文字の無い c:title は系列が 1 つなら系列名・2 つ以上なら出さない（C10）', () => {
    const one = texts(scene(mk({ title: { lines: [], text: {}, overlay: false } }))).map(
      (t) => t.text,
    );
    expect(one).toContain('S1');
    const two = texts(
      scene(
        mk({
          title: { lines: [], text: {}, overlay: false },
          series: [ser(0, 'a', [1, 2, 3]), ser(1, 'b', [1, 2, 3])],
        }),
      ),
    ).map((t) => t.text);
    expect(two).not.toContain('a');
    // c:title が無く autoTitleDeleted が真なら出さない
    expect(texts(scene(mk())).map((t) => t.text)).not.toContain('S1');
  });

  it('TD-104-20: c:title が無くても autoTitleDeleted が偽で系列が 1 つなら系列名がタイトルになる（Excel 16 の照合 C10）。2 系列なら出さない', () => {
    const one = scene(mk({ autoTitleDeleted: false }));
    const t = texts(one).find((x) => x.text === 'S1')!;
    expect(t).toBeDefined();
    expect(t.anchor).toBe('middle');
    expect(t.size).toBeCloseTo(10 * PX_PER_PT, 10); // c:txPr の無い既定のタイトル（10pt・太字）
    expect(t.bold).toBe(true);
    const two = mk({
      autoTitleDeleted: false,
      series: [ser(0, 'a', [1, 2, 3]), ser(1, 'b', [1, 2, 3])],
    });
    expect(texts(scene(two)).map((x) => x.text)).not.toContain('a');
  });

  it('TD-104-21: タイトルの上端は面の上から 6pt・プロットの上はタイトルの下から 11pt（Excel 16 の照合 C5）', () => {
    const s = scene(
      mk({ title: { lines: ['T'], text: { sizePt: 14 }, overlay: false } }),
      480,
      288,
    );
    const t = texts(s).find((x) => x.text === 'T')!;
    const h = 14 * PX_PER_PT * 1.45; // タイトルの高さ
    expect(t.y).toBeCloseTo(6 * PX_PER_PT + h / 2, 5);
    const grid = paths(s).find((p) => p.fill === null && p.stroke?.color === '#000000')!;
    const top = Math.min(
      ...[...grid.d.matchAll(/[ML]([\d.]+) ([\d.]+)/g)].map((m) => Number(m[2])),
    );
    expect(top).toBeCloseTo(6 * PX_PER_PT + h + 11 * PX_PER_PT, 1);
  });

  it('TD-102-33: タイトルの複数行は行ごとに 1 つの text', () => {
    const t = texts(scene(mk({ title: { lines: ['L1', 'L2'], text: {}, overlay: false } }))).filter(
      (x) => x.text.startsWith('L'),
    );
    expect(t).toHaveLength(2);
    expect(t[1]!.y).toBeGreaterThan(t[0]!.y);
    expect(t[0]!.anchor).toBe('middle');
  });

  it('TD-102-34: 凡例の位置 5 種（r・l・t・b・tr）で凡例の文字の位置とプロットの矩形が変わる', () => {
    const none = base({});
    for (const pos of ['r', 'l', 't', 'b', 'tr'] as const) {
      const m = mk({ plotArea: { fill: solid(0x123456) }, legend: legend(pos) });
      const s = scene(m);
      const r = plotRect(s);
      const t = texts(s).find((x) => x.text === 'S1')!;
      if (pos === 'r' || pos === 'tr') {
        expect(t.x).toBeGreaterThan(r.x + r.width);
        expect(r.width).toBeLessThan(none.width);
        if (pos === 'tr') expect(t.y).toBeLessThan(80);
        else expect(Math.abs(t.y - 150)).toBeLessThan(15); // 縦に中央
      } else if (pos === 'l') {
        expect(t.x).toBeLessThan(r.x);
        expect(r.x).toBeGreaterThan(none.x);
      } else if (pos === 't') {
        expect(t.y).toBeLessThan(r.y);
        expect(r.height).toBeLessThan(none.height);
      } else {
        expect(t.y).toBeGreaterThan(r.y + r.height);
        expect(r.height).toBeLessThan(none.height);
      }
    }
  });

  it('TD-102-34: 凡例の overlay ではプロットの矩形が変わらない', () => {
    const none = base({});
    for (const pos of ['r', 'l', 't', 'b'] as const) {
      const r = base({ legend: legend(pos, true) });
      expect([r.x, r.y, r.width, r.height]).toEqual([none.x, none.y, none.width, none.height]);
    }
  });

  it('TD-102-35: 凡例の項目と記号：棒は四角・折れ線は線＋マーカー・Series n・c:legendEntry の削除', () => {
    const m = mk({
      legend: { pos: 'r', overlay: false, deleted: [1], text: {} },
      series: [ser(0, 'Alpha', [1, 2, 3]), ser(1, 'Beta', [1, 2, 3]), ser(2, '', [1, 2, 3])],
    });
    const s = scene(m);
    const tx = texts(s).map((t) => t.text);
    expect(tx).toContain('Alpha');
    expect(tx).not.toContain('Beta');
    expect(tx).toContain('Series 3');
    const squares = rects(s).filter((r) => r.width < 12 && r.width === r.height);
    expect(squares.map((r) => r.fill!.color)).toEqual([hexOf(0x4472c4), hexOf(0xa5a5a5)]);

    const ln = scene(
      mk({
        type: 'line',
        barDir: undefined,
        grouping: 'standard',
        lineMarkers: true,
        legend: { pos: 'b', overlay: false, deleted: [], text: {} },
      }),
    );
    expect(
      ln.items.some(
        (i) =>
          i.kind === 'line' && i.stroke.color === '#4472c4' && i.stroke.width === 1.5 * PX_PER_PT,
      ),
    ).toBe(true);
    // 凡例の記号のマーカー（小さい円の経路）がある
    expect(paths(ln).some((p) => p.d.includes('A') && p.fill?.color === '#4472c4')).toBe(true);
  });

  it('TD-102-35: 円・ドーナツと varyColors の 1 系列の棒の凡例は項目', () => {
    const p = scene(pie({ legend: { pos: 'r', overlay: false, deleted: [], text: {} } }));
    expect(texts(p).map((t) => t.text)).toEqual(expect.arrayContaining(['A', 'B', 'C']));
    const b = scene(
      mk({ varyColors: true, legend: { pos: 'r', overlay: false, deleted: [], text: {} } }),
    );
    expect(texts(b).filter((t) => ['A', 'B', 'C'].includes(t.text)).length).toBe(6); // 軸の文字 3 ＋凡例 3
    expect(texts(b).some((t) => t.text === 'S1')).toBe(false);
  });

  it('TD-102-35: 入らない項目は出さない（低いグラフ）', () => {
    const many = Array.from({ length: 30 }, (_, i) => ser(i, `N${i}`, [1, 2, 3]));
    const s = scene(
      mk({ series: many, legend: { pos: 'r', overlay: false, deleted: [], text: {} } }),
      400,
      120,
    );
    const names = texts(s).filter((t) => /^N\d+$/.test(t.text));
    expect(names.length).toBeGreaterThan(0);
    expect(names.length).toBeLessThan(30);
  });

  it('TD-102-36: 横並びの凡例は入らなければ折り返す（行が増える）', () => {
    const many = Array.from({ length: 12 }, (_, i) => ser(i, `Series name ${i}`, [1, 2, 3]));
    const s = scene(
      mk({ series: many, legend: { pos: 'b', overlay: false, deleted: [], text: {} } }),
      300,
      300,
    );
    const ys = new Set(
      texts(s)
        .filter((t) => t.text.startsWith('Series name'))
        .map((t) => Math.round(t.y)),
    );
    expect(ys.size).toBeGreaterThan(1);
  });
});

// ---- 軸の文字・向き ----

describe('軸の文字（数値軸・項目軸・軸タイトル）', () => {
  it('TD-102-37: 縦棒の数値軸の文字は左に右揃え・目盛の値の書式つき', () => {
    const s = scene(mk());
    const t = texts(s).filter((x) => /^\d+$/.test(x.text));
    expect(t.map((x) => x.text)).toEqual(['0', '5', '10', '15', '20', '25', '30', '35']);
    expect(t.every((x) => x.anchor === 'end' && x.baseline === 'middle')).toBe(true);
    expect(t[0]!.y).toBeGreaterThan(t.at(-1)!.y); // 0 が下
  });

  it('TD-102-37: percentStacked の数値軸の文字は 0%…100%', () => {
    const m = mk({
      grouping: 'percentStacked',
      series: [ser(0, 'a', [1, 2, 3]), ser(1, 'b', [3, 2, 1])],
    });
    const tx = texts(scene(m)).map((t) => t.text);
    expect(tx).toContain('0%');
    expect(tx).toContain('100%');
    expect(tx).toContain('50%');
  });

  it('TD-102-37: 数値軸の numFmt（#,##0）', () => {
    const m = mk({
      series: [ser(0, 'a', [1000, 2000, 3000])],
      valAxis: { ...valAx, numFmt: { code: '#,##0', sourceLinked: false } },
    });
    expect(texts(scene(m)).map((t) => t.text)).toContain('3,000');
  });

  it('TD-102-38: 項目軸の文字は項目の下・帯の中央', () => {
    const s = scene(mk());
    const [a, b, c] = ['A', 'B', 'C'].map((k) => texts(s).find((t) => t.text === k)!);
    expect(a!.x).toBeLessThan(b!.x);
    expect(b!.x).toBeLessThan(c!.x);
    expect(b!.x - a!.x).toBeCloseTo(c!.x - b!.x, 5);
    expect(a!.baseline).toBe('hanging');
    expect(a!.y).toBe(b!.y);
  });

  it('TD-102-38: 軸の文字は deleted・tickLblPos none で出ない（余白も取らない）', () => {
    const withAx = plotRect(scene(mk({ plotArea: { fill: solid(0x123456) } })));
    const m = mk({
      plotArea: { fill: solid(0x123456) },
      catAxis: { ...catAx, deleted: true },
      valAxis: { ...valAx, tickLblPos: 'none' },
    });
    const s = scene(m);
    expect(texts(s)).toHaveLength(0);
    expect(plotRect(s).width).toBeGreaterThan(withAx.width);
    expect(plotRect(s).height).toBeGreaterThan(withAx.height);
  });

  it('TD-102-39: 軸タイトル：縦の数値軸は −90°・左に置く／横軸は下', () => {
    const m = mk({
      valAxis: { ...valAx, title: { lines: ['Value'], text: {}, overlay: false } },
      catAxis: { ...catAx, title: { lines: ['Cat'], text: {}, overlay: false } },
    });
    const s = scene(m);
    const v = texts(s).find((t) => t.text === 'Value')!;
    const c = texts(s).find((t) => t.text === 'Cat')!;
    expect(v.rotate).toBe(-90);
    expect(v.x).toBeLessThan(30);
    expect(c.rotate).toBe(0);
    expect(c.y).toBeGreaterThan(270);
    // 横棒では数値軸が横（回さない）・項目軸が縦（−90°）
    const h = scene({ ...m, barDir: 'bar' });
    expect(texts(h).find((t) => t.text === 'Value')!.rotate).toBe(0);
    expect(texts(h).find((t) => t.text === 'Cat')!.rotate).toBe(-90);
  });

  it('TD-102-40: 横棒は先頭の項目が下・maxMin で上', () => {
    const y = (s: ChartScene, k: string): number => texts(s).find((t) => t.text === k)!.y;
    const h = scene(mk({ barDir: 'bar' }));
    expect(y(h, 'A')).toBeGreaterThan(y(h, 'C'));
    const r = scene(mk({ barDir: 'bar', catAxis: { ...catAx, orientation: 'maxMin' } }));
    expect(y(r, 'A')).toBeLessThan(y(r, 'C'));
    // 項目の文字は左に右揃え
    expect(texts(h).find((t) => t.text === 'A')!.anchor).toBe('end');
  });

  it('TD-102-40: 縦棒の項目軸 maxMin は先頭が右', () => {
    const x = (s: ChartScene, k: string): number => texts(s).find((t) => t.text === k)!.x;
    const r = scene(mk({ catAxis: { ...catAx, orientation: 'maxMin' } }));
    expect(x(r, 'A')).toBeGreaterThan(x(r, 'C'));
  });

  it('TD-102-41: 項目の文字は tickLblSkip の明示で間引く・推定の幅が帯に入らなければ自動で間引く', () => {
    const cats = Array.from({ length: 10 }, (_, i) => `C${i}`);
    const vals = cats.map((_, i) => i);
    const lbl = (m: ChartModel, w = 400): string[] =>
      texts(scene(m, w))
        .filter((t) => /^C\d$/.test(t.text))
        .map((t) => t.text);
    expect(
      lbl(
        mk({
          categories: cats,
          series: [ser(0, 's', vals)],
          catAxis: { ...catAx, tickLblSkip: 3 },
        }),
      ),
    ).toEqual(['C0', 'C3', 'C6', 'C9']);
    expect(lbl(mk({ categories: cats, series: [ser(0, 's', vals)] }))).toHaveLength(10);
    // 面の高さが低く、回した帯が入らない小さい面は、回さず間引く（ISSUE-107 の安全弁。回せる大きさは TD-107-7・12）
    const long = Array.from({ length: 10 }, (_, i) => `Category-name-${i}`);
    const out = texts(
      scene(mk({ categories: long, series: [ser(0, 's', vals)] }), 300, 100),
    ).filter((t) => t.text.startsWith('Category'));
    expect(out.length).toBeLessThan(10);
    expect(out[0]!.text).toBe('Category-name-0');
    expect(out.every((t) => t.rotate === 0)).toBe(true);
  });

  it('TD-102-41: a:bodyPr rot の明示は回転に従う', () => {
    const s = scene(mk({ catAxis: { ...catAx, text: { rotDeg: -45 } } }));
    const a = texts(s).find((t) => t.text === 'A')!;
    expect(a.rotate).toBe(-45);
    expect(a.anchor).toBe('end');
  });

  it('TD-102-42: 項目軸の線は 0 の位置（負を含む時は中ほど）・crosses=max は上', () => {
    const lineY = (m: ChartModel): number => {
      const l = scene(m).items.find((i) => i.kind === 'line')!;
      return l.kind === 'line' ? l.y1 : NaN;
    };
    const m = mk({ series: [ser(0, 's', [-10, 20, 30])] });
    const mid = lineY(m);
    const bottom = lineY(mk());
    expect(mid).toBeLessThan(bottom);
    const top = lineY({ ...m, catAxis: { ...catAx, crosses: 'max' } });
    expect(top).toBeLessThan(mid);
  });
});

// ---- 棒 ----

describe('棒の図形（§決定 6）', () => {
  const inner = innerPlot(0.1, 0.1, 0.75, 0.6); // 400×300 → プロット 300×180・帯 100

  it('TD-102-50: 集合の棒の幅 w＝B ÷ (n − (n−1)·o/100 + g/100)（n=2・g=150・o=0 → 100/3.5）', () => {
    const m = mk({ ...inner, series: [ser(0, 'a', [10, 20, 30]), ser(1, 'b', [10, 20, 30])] });
    const s = scene(m);
    const rs = fillPaths(s).map((p) => rectsOf(p.d));
    expect(rs[0]).toHaveLength(3);
    expect(rs[0]![0]!.w).toBeCloseTo(100 / 3.5, 1);
    // 系列 b は a の右隣（隙間なし・overlap 0）
    expect(rs[1]![0]!.x).toBeCloseTo(rs[0]![0]!.x + rs[0]![0]!.w, 1);
    // 帯の中央に集まる（隙間 g/2 ずつ）
    const left = rs[0]![0]!.x - 40;
    expect(left).toBeCloseTo(((100 / 3.5) * 1.5) / 2, 1);
  });

  it('TD-102-50: gapWidth・overlap（overlap=100 なら系列が重なり w＝B ÷ (1 + g/100)）', () => {
    const two = [ser(0, 'a', [10, 20, 30]), ser(1, 'b', [10, 20, 30])];
    const rs = fillPaths(scene(mk({ ...inner, series: two, gapWidth: 100, overlap: 100 }))).map(
      (p) => rectsOf(p.d),
    );
    expect(rs[0]![0]!.w).toBeCloseTo(50, 1);
    expect(rs[1]![0]!.x).toBeCloseTo(rs[0]![0]!.x, 1);
    const r0 = fillPaths(scene(mk({ ...inner, series: two, gapWidth: 0, overlap: 0 }))).map((p) =>
      rectsOf(p.d),
    );
    expect(r0[0]![0]!.w).toBeCloseTo(50, 1);
    const neg = fillPaths(scene(mk({ ...inner, series: two, gapWidth: 100, overlap: -50 }))).map(
      (p) => rectsOf(p.d),
    );
    expect(neg[0]![0]!.w).toBeCloseTo(100 / (2 + 0.5 + 1), 1);
    expect(neg[1]![0]!.x - (neg[0]![0]!.x + neg[0]![0]!.w)).toBeCloseTo(neg[0]![0]!.w * 0.5, 1);
  });

  it('TD-102-50: 積み上げは n=1 とみなす（w＝100/2.5）・正の値は上へ積み上がる', () => {
    const m = mk({
      ...inner,
      grouping: 'stacked',
      series: [ser(0, 'a', [10, 20, 30]), ser(1, 'b', [10, 20, 30])],
    });
    const rs = fillPaths(scene(m)).map((p) => rectsOf(p.d));
    expect(rs[0]![0]!.w).toBeCloseTo(40, 1);
    expect(rs[1]![0]!.x).toBeCloseTo(rs[0]![0]!.x, 1);
    expect(rs[1]![0]!.y + rs[1]![0]!.h).toBeCloseTo(rs[0]![0]!.y, 1); // b の底は a の頂
    expect(rs[1]![0]!.h).toBeCloseTo(rs[0]![0]!.h, 1);
  });

  it('TD-102-51: 正負の積み上げ：負は 0 から下へ別に積む', () => {
    const m = mk({
      ...inner,
      grouping: 'stacked',
      series: [
        ser(0, 'a', [10, 10, 10]),
        ser(1, 'b', [-10, -10, -10]),
        ser(2, 'c', [-10, -10, -10]),
      ],
    });
    const rs = fillPaths(scene(m)).map((p) => rectsOf(p.d)[0]!);
    const zeroY = rs[0]!.y + rs[0]!.h; // a の底が 0
    expect(rs[1]!.y).toBeCloseTo(zeroY, 1);
    expect(rs[2]!.y).toBeCloseTo(rs[1]!.y + rs[1]!.h, 1);
  });

  it('TD-102-51: 100% 積み上げは各項目の合計が軸全体を埋める', () => {
    const m = mk({
      ...inner,
      grouping: 'percentStacked',
      series: [ser(0, 'a', [1, 2, 3]), ser(1, 'b', [3, 2, 1])],
    });
    const [a, b] = fillPaths(scene(m)).map((p) => rectsOf(p.d));
    for (let i = 0; i < 3; i++) expect(a![i]!.h + b![i]!.h).toBeCloseTo(180, 1);
  });

  it('TD-102-52: null の点は棒を作らない・正の値の棒は 0 の線から上へ・負は下へ', () => {
    const m = mk({ ...inner, series: [ser(0, 'a', [10, null, -10])] });
    const r = rectsOf(fillPaths(scene(m))[0]!.d);
    expect(r).toHaveLength(2);
    expect(r[0]!.y).toBeLessThan(r[1]!.y); // 正は上・負は下
  });

  it('TD-102-52: 横棒は左から右へ伸び、系列は下から並ぶ', () => {
    const m = mk({
      ...inner,
      barDir: 'bar',
      series: [ser(0, 'a', [10, 20, 30]), ser(1, 'b', [10, 20, 30])],
    });
    const rs = fillPaths(scene(m)).map((p) => rectsOf(p.d));
    expect(rs[0]![0]!.w).toBeLessThan(rs[0]![2]!.w);
    expect(rs[0]![0]!.y).toBeGreaterThan(rs[0]![2]!.y); // 先頭の項目が下
    expect(rs[0]![0]!.y).toBeGreaterThan(rs[1]![0]!.y); // 系列 a は b の下
    expect(rs[0]![0]!.h).toBeCloseTo(60 / 3.5, 0);
  });

  it('TD-102-52: 明示の min で範囲外の棒は端に収まる（はみ出さない）', () => {
    const m = mk({ ...inner, valAxis: { ...valAx, min: 0, max: 15 } });
    const r = rectsOf(fillPaths(scene(m))[0]!.d);
    const top = 40 * 0 + 30; // プロットの上端 y=30
    expect(Math.min(...r.map((x) => x.y))).toBeGreaterThanOrEqual(top - 1e-9);
  });
});

// ---- 折れ線 ----

describe('折れ線（§決定 6）', () => {
  const line = (over: Partial<ChartModel> = {}): ChartModel =>
    mk({
      type: 'line',
      barDir: undefined,
      grouping: 'standard',
      categories: ['A', 'B', 'C', 'D'],
      ...over,
    });
  const linePath = (s: ChartScene): ChartPathItem =>
    paths(s).find((p) => p.fill === null && p.stroke?.width === 1.5 * PX_PER_PT)!;
  const subpaths = (d: string): number => (d.match(/M/g) ?? []).length;

  it('TD-102-60: 空の点 3 種：gap は線が切れる・zero は 0 を通る・span はつなぐ', () => {
    const vals = [10, null, 30, 40];
    const gap = linePath(scene(line({ series: [ser(0, 's', vals)], dispBlanksAs: 'gap' })));
    expect(subpaths(gap.d)).toBe(1); // 区間は 1 点＋ 2 点：2 点の区間だけ線になる
    const zero = linePath(scene(line({ series: [ser(0, 's', vals)], dispBlanksAs: 'zero' })));
    expect(subpaths(zero.d)).toBe(1);
    expect((zero.d.match(/L/g) ?? []).length).toBe(3);
    const span = linePath(scene(line({ series: [ser(0, 's', vals)], dispBlanksAs: 'span' })));
    expect(subpaths(span.d)).toBe(1);
    expect((span.d.match(/L/g) ?? []).length).toBe(2);
    // gap：2 区間（1 点＋2 点）で、2 点の区間だけ線になる（L は 1）
    expect((gap.d.match(/L/g) ?? []).length).toBe(1);
    const gap2 = linePath(
      scene(
        line({
          series: [ser(0, 's', [10, 20, null, 30, 40])],
          categories: ['A', 'B', 'C', 'D', 'E'],
          dispBlanksAs: 'gap',
        }),
      ),
    );
    expect(subpaths(gap2.d)).toBe(2);
  });

  it('TD-102-60: 点が帯の中央にある（既定）・midCat なら目盛の上（端から端）', () => {
    const inner = innerPlot(0.1, 0.1, 0.8, 0.6); // プロット幅 320
    const xs = (d: string): number[] =>
      (d.match(/[ML](-?\d+(?:\.\d+)?)/g) ?? []).map((t) => Number(t.slice(1)));
    const mid = xs(linePath(scene(line({ ...inner, series: [ser(0, 's', [1, 2, 3, 4])] }))).d);
    expect(mid[0]).toBeCloseTo(40 + 40, 1);
    const tick = xs(
      linePath(
        scene(
          line({
            ...inner,
            series: [ser(0, 's', [1, 2, 3, 4])],
            valAxis: { ...valAx, crossBetween: 'midCat' },
          }),
        ),
      ).d,
    );
    expect(tick[0]).toBeCloseTo(40, 1);
    expect(tick.at(-1)).toBeCloseTo(360, 1);
  });

  it('TD-102-61: 積み上げは累積した値の位置に点が来る', () => {
    const m = line({
      grouping: 'stacked',
      ...innerPlot(0.1, 0.1, 0.8, 0.6),
      series: [ser(0, 'a', [10, 10, 10, 10]), ser(1, 'b', [10, 10, 10, 10])],
    });
    const ys = (d: string): number[] =>
      (d.match(/[ML]-?[\d.]+ (-?[\d.]+)/g) ?? []).map((t) => Number(t.split(' ')[1]));
    const [pa, pb] = paths(scene(m)).filter(
      (p) => p.fill === null && p.stroke?.width === 1.5 * PX_PER_PT,
    );
    expect(ys(pb!.d)[0]).toBeLessThan(ys(pa!.d)[0]!); // b は a の上
  });

  it('TD-102-62: smooth は 3 次ベジェ（C）・そうでなければ直線', () => {
    const smooth = linePath(scene(line({ series: [ser(0, 's', [1, 3, 2, 4], { smooth: true })] })));
    expect(smooth.d).toContain('C');
    expect(smooth.d).not.toContain('L');
    expect(linePath(scene(line({ series: [ser(0, 's', [1, 3, 2, 4])] }))).d).not.toContain('C');
  });

  it('TD-102-63: マーカー：自動は円・大きさ 6pt・系列の色で塗る／symbol none では出ない・グループの c:marker（lineMarkers）は見ない（C13）', () => {
    const withM = scene(line({ lineMarkers: true, series: [ser(0, 's', [1, 2, 3, 4])] }));
    const mp = paths(withM).find((p) => p.fill?.color === '#4472c4' && p.d.includes('A'))!;
    expect((mp.d.match(/M/g) ?? []).length).toBe(4);
    expect(CHART_DEFAULTS.markerSizePt).toBe(6);
    expect(mp.d).toContain('A4 4'); // 6pt × 96/72 ÷ 2
    // グループの c:marker が無い・0 でも、系列に c:marker が無ければ自動の円（Excel 16 の COM の MarkerStyle が自動・C13）
    for (const lineMarkers of [undefined, false]) {
      const m = scene(line({ lineMarkers, series: [ser(0, 's', [1, 2, 3, 4])] }));
      expect(paths(m).some((p) => p.d.includes('A'))).toBe(true);
    }
    const none = scene(
      line({
        lineMarkers: true,
        series: [ser(0, 's', [1, 2, 3, 4], { marker: { symbol: 'none' } })],
      }),
    );
    expect(paths(none).some((p) => p.d.includes('A'))).toBe(false);
    const sq = scene(
      line({
        series: [
          ser(0, 's', [1, 2, 3, 4], { marker: { symbol: 'square', fill: solid(0x00ff00) } }),
        ],
      }),
    );
    const sp = paths(sq).find((p) => p.fill?.color === '#00ff00')!;
    expect((sp.d.match(/Z/g) ?? []).length).toBe(4);
  });

  it('TD-102-63: x・plus などの線だけの形は塗らない', () => {
    const s = scene(line({ series: [ser(0, 's', [1, 2, 3, 4], { marker: { symbol: 'x' } })] }));
    const mp = paths(s).find(
      (p) =>
        p.fill === null &&
        p.stroke?.color === '#4472c4' &&
        p.stroke.width === 0.75 * PX_PER_PT &&
        p.d.includes('L'),
    )!;
    expect((mp.d.match(/M/g) ?? []).length).toBe(8); // 4 点 × 線 2 本
    expect(mp.stroke!.color).toBe('#4472c4');
  });

  it('TD-102-64: 線が none（a:ln noFill）なら線は描かずマーカーだけ', () => {
    const s = scene(
      line({
        lineMarkers: true,
        series: [ser(0, 's', [1, 2, 3, 4], { line: { paint: { kind: 'none' } } })],
      }),
    );
    expect(paths(s).some((p) => p.fill === null && p.stroke?.width === 1.5 * PX_PER_PT)).toBe(
      false,
    );
    expect(paths(s).some((p) => p.d.includes('A'))).toBe(true);
  });
});

// ---- 円・ドーナツ ----

describe('円・ドーナツ（§決定 6・C9）', () => {
  const sectors = (s: ChartScene): ChartPathItem[] => fillPaths(s);
  const inner = (over: Partial<ChartModel> = {}): ChartModel =>
    pie({
      ...innerPlot(0.25, 0.25, 0.5, 0.5),
      categories: ['a', 'b', 'c', 'd'],
      series: [ser(0, 'S', [1, 1, 1, 1])],
      ...over,
    });
  // プロット 200×150 → 中心 (200, 150)・半径 75

  it('TD-102-70: 扇は点ごとに 1 つ・null と 0 は作らない・負は絶対値', () => {
    const s = scene(inner({ series: [ser(0, 'S', [1, null, 0, -1])] }));
    expect(sectors(s).filter((p) => p.d.startsWith('M200 150'))).toHaveLength(2);
    // 負の 1 は 1 と同じ大きさ（半周ずつ）
    const [a, b] = sectors(s).filter((p) => p.d.startsWith('M200 150'));
    expect(a!.d).toContain('A75 75 0 0 1');
    expect(b!.d).toContain('A75 75 0 0 1');
  });

  it('TD-102-70: 円は最初の系列だけ', () => {
    const s = scene(inner({ series: [ser(0, 'A', [1, 1, 1, 1]), ser(1, 'B', [1, 1, 1, 1])] }));
    expect(sectors(s).filter((p) => p.d.startsWith('M200 150'))).toHaveLength(4);
  });

  it('TD-102-71: 12 時から時計回り：最初の扇は 12 時（200,75）から 3 時（275,150）へ', () => {
    const s = scene(inner());
    expect(sectors(s)[0]!.d).toBe('M200 150L200 75A75 75 0 0 1 275 150Z');
    expect(sectors(s)[1]!.d).toBe('M200 150L275 150A75 75 0 0 1 200 225Z');
  });

  it('TD-102-71: firstSliceAng=90 なら最初の扇は 3 時から始まる・270 では 9 時から', () => {
    expect(sectors(scene(inner({ firstSliceAng: 90 })))[0]!.d.startsWith('M200 150L275 150')).toBe(
      true,
    );
    expect(sectors(scene(inner({ firstSliceAng: 270 })))[0]!.d.startsWith('M200 150L125 150')).toBe(
      true,
    );
  });

  it('TD-102-71: 半周を超える扇は large-arc・全周 1 つなら円', () => {
    const big = sectors(
      scene(inner({ series: [ser(0, 'S', [3, 1])], categories: ['a', 'b'] })),
    )[0]!;
    expect(big.d).toContain('A75 75 0 1 1');
    const full = sectors(scene(inner({ series: [ser(0, 'S', [5])], categories: ['a'] })))[0]!;
    expect(full.d).toBe('M125 150A75 75 0 1 1 275 150A75 75 0 1 1 125 150Z');
  });

  it('TD-102-71: 全部 0 なら扇なし・円は plot の内側に収まる最大（短い辺の半分）', () => {
    expect(sectors(scene(inner({ series: [ser(0, 'S', [0, 0, 0, 0])] })))).toHaveLength(0);
    const r = fillPaths(
      scene(pie({ series: [ser(0, 'S', [1, 1])], categories: ['a', 'b'] }), 400, 200),
    )[0]!;
    expect(r.d).toContain('A'); // 自動配置でも描く
  });

  it('TD-102-72: ドーナツは系列ごとの同心の輪・穴は holeSize（既定 50%）', () => {
    const m = inner({
      type: 'doughnut',
      series: [ser(0, 'A', [1, 1, 1, 1]), ser(1, 'B', [1, 1, 1, 1])],
    });
    const s = scene(m);
    const ds = sectors(s).map((p) => p.d);
    expect(ds).toHaveLength(8);
    // 内輪（最初の系列 A）：外半径 56.25・内半径 37.5（穴 50%。Excel は最初の系列が内側）
    expect(ds[0]).toContain('A56.25 56.25');
    expect(ds[0]).toContain('A37.5 37.5');
    expect(ds[0]).not.toContain('A75 75');
    // 外輪（系列 B）：外半径 75・内半径 56.25（穴 37.5 と外 75 の間を 2 等分）
    expect(ds[4]).toContain('A75 75');
    expect(ds[4]).toContain('A56.25 56.25');
    const h = scene(inner({ type: 'doughnut', holeSize: 70 }));
    expect(sectors(h)[0]!.d).toContain('A52.5 52.5');
  });

  it('TD-102-72: ドーナツの全周 1 つは外円と逆回りの内円（穴）', () => {
    const m = inner({ type: 'doughnut', series: [ser(0, 'S', [5])], categories: ['a'] });
    const d = sectors(scene(m))[0]!.d;
    expect(d.match(/Z/g)).toHaveLength(2);
    expect(d).toContain('0 1 0'); // 内円は逆回り
  });

  it('TD-102-73: 点ごとの色（c:dPt）と円の枠線', () => {
    const m = inner({
      series: [ser(0, 'S', [1, 1, 1, 1], { points: new Map([[2, solid(0xabcdef)]]) })],
    });
    const s = scene(m);
    expect(sectors(s)[2]!.fill!.color).toBe('#abcdef');
    expect(sectors(s)[0]!.stroke).toMatchObject({ color: '#ffffff' });
  });
});

// ---- 場面の安全性・上限 ----

describe('場面の属性（ファイル由来の文字列は text にだけ入る）・上限', () => {
  const evil = '"><script>alert(1)</script>';

  const FIXED_WORDS = new Set([
    'solid',
    'dot',
    'dash',
    'lgDash',
    'dashDot',
    'lgDashDot',
    'lgDashDotDot',
    'sysDash',
    'sysDot',
    'sysDashDot',
    'sysDashDotDot',
    'butt',
    'round',
    'miter',
    'start',
    'middle',
    'end',
    'alphabetic',
    'hanging',
  ]);

  /** text 以外の全属性が数値・#rrggbb・固定の語だけで、d が数値と固定の命令だけ */
  function assertSafe(s: ChartScene): void {
    expect(typeof s.backdrop).toBe('boolean');
    const check = (v: unknown, key: string): void => {
      if (v === null || typeof v === 'boolean') return;
      if (typeof v === 'number') {
        expect(Number.isFinite(v)).toBe(true);
        return;
      }
      if (typeof v === 'string') {
        if (key === 'd') expect(v).toMatch(/^[MLCAZ0-9 .-]*$/);
        else if (key === 'color') expect(v).toMatch(/^#[0-9a-f]{6}$/);
        else expect(FIXED_WORDS.has(v) || ['rect', 'path', 'line', 'text'].includes(v)).toBe(true);
        return;
      }
      if (typeof v === 'object') for (const [k, x] of Object.entries(v)) check(x, k);
    };
    for (const it of s.items as ChartItem[]) {
      for (const [k, v] of Object.entries(it)) {
        if (it.kind === 'text' && k === 'text') continue;
        check(v, k);
      }
    }
  }

  const hostile = (): ChartModel =>
    mk({
      categories: [evil, 'B', 'C'],
      series: [ser(0, evil, [1, 2, 3])],
      title: { lines: [evil], text: {}, overlay: false },
      legend: { pos: 'r', overlay: false, deleted: [], text: {} },
      catAxis: { ...catAx, title: { lines: [evil], text: {}, overlay: false } },
      valAxis: { ...valAx, numFmt: { code: evil, sourceLinked: false } },
    });

  it('TD-102-80: 棒・折れ線・円・ドーナツのどれでも text 以外の属性は数値・#rrggbb・固定の語だけ', () => {
    const models: ChartModel[] = [
      hostile(),
      { ...hostile(), barDir: 'bar' },
      { ...hostile(), type: 'line', grouping: 'standard', lineMarkers: true },
      {
        ...hostile(),
        type: 'pie',
        grouping: 'standard',
        catAxis: undefined,
        valAxis: undefined,
        varyColors: true,
      },
      {
        ...hostile(),
        type: 'doughnut',
        grouping: 'standard',
        catAxis: undefined,
        valAxis: undefined,
        varyColors: true,
      },
    ];
    for (const m of models) {
      const s = scene(m);
      assertSafe(s);
      // 悪意のある文字列は text の文字にだけ現れる
      const nonText = JSON.stringify(s.items.filter((i) => i.kind !== 'text'));
      expect(nonText).not.toContain('script');
      expect(nonText).not.toContain('alert');
      expect(texts(s).some((t) => t.text.includes('script'))).toBe(true);
    }
  });

  it('TD-102-80: 数値は NaN・Infinity を含まない（壊れた寸法・値でも）', () => {
    const m = mk({
      series: [ser(0, 's', [Number.NaN, Infinity, 1e308])],
      gapWidth: Number.NaN,
      overlap: Number.NaN,
    });
    assertSafe(scene(m));
    assertSafe(scene(mk(), 3, 3));
  });

  it('TD-102-81: 大きさが不正（0・負・NaN・Infinity）なら null', () => {
    for (const [w, h] of [
      [0, 100],
      [100, 0],
      [-1, 100],
      [Number.NaN, 100],
      [100, Infinity],
    ] as const) {
      expect(layoutChart(mk(), w, h)).toBeNull();
    }
  });

  it('TD-102-82: 場面の要素が MAX_CHART_SCENE_ITEMS を超えたら null（項目の文字を 1 個おきでなく全部出す場合）', () => {
    expect(MAX_CHART_SCENE_ITEMS).toBe(5000);
    const n = 6000;
    const cats = Array.from({ length: n }, (_, i) => `c${i}`);
    const m = mk({
      categories: cats,
      series: [
        ser(
          0,
          's',
          cats.map((_, i) => i),
        ),
      ],
      catAxis: { ...catAx, tickLblSkip: 1 },
    });
    expect(layoutChart(m, 800, 300)).toBeNull();
    // 既定（自動の間引き）なら出せる
    const auto = layoutChart(
      mk({
        categories: cats,
        series: [
          ser(
            0,
            's',
            cats.map((_, i) => i),
          ),
        ],
      }),
      800,
      300,
    );
    expect(auto).not.toBeNull();
    expect(auto!.items.length).toBeLessThanOrEqual(MAX_CHART_SCENE_ITEMS);
  });

  it('TD-102-82: 円の扇が上限を超える（系列×項目 > 5,000）と null', () => {
    const cats = Array.from({ length: 5001 }, (_, i) => `c${i}`);
    expect(
      layoutChart(
        pie({
          categories: cats,
          series: [
            ser(
              0,
              's',
              cats.map(() => 1),
            ),
          ],
        }),
        400,
        300,
      ),
    ).toBeNull();
    const ok = Array.from({ length: 300 }, (_, i) => `c${i}`);
    expect(
      layoutChart(
        pie({
          categories: ok,
          series: [
            ser(
              0,
              's',
              ok.map(() => 1),
            ),
          ],
        }),
        400,
        300,
      ),
    ).not.toBeNull();
  });

  it('TD-102-83: 1 つの場面で全部の要素が上限以内・同じ入力は同じ出力（純関数）', () => {
    const a = scene(hostile());
    const b = scene(hostile());
    expect(a).toEqual(b);
    expect(a.items.length).toBeLessThanOrEqual(MAX_CHART_SCENE_ITEMS);
  });
});

// ---- ISSUE-102 の独立レビューの指摘（TD-102-90〜） ----

describe('レビュー指摘の修正（ドーナツの輪の順・負の 100% 積み上げ・nextTo・軸タイトル・上限・横軸の目盛）', () => {
  const plotted = (m: Partial<ChartModel>): ChartModel =>
    mk({ plotArea: { fill: solid(0x123456) }, ...m });
  const labelOf = (s: ChartScene, t: string): ChartTextItem => texts(s).find((x) => x.text === t)!;
  const axisLine = (s: ChartScene): { x1: number; y1: number } => {
    const l = s.items.find((i) => i.kind === 'line')!;
    if (l.kind !== 'line') throw new Error('line ではない');
    return { x1: l.x1, y1: l.y1 };
  };
  const TICK_GAP = 5 * PX_PER_PT + 3; // 目盛の長さ 5pt + 文字との間隔 3px

  it('TD-102-90: 負を含む 100% 積み上げ：負の割合の棒が出て、数値軸の最小は -1（ISSUE-104：Excel 16 の照合 C2）', () => {
    const m = plotted({
      ...innerPlot(0.1, 0.1, 0.8, 0.6),
      grouping: 'percentStacked',
      categories: ['A'],
      series: [ser(0, 'a', [-30]), ser(1, 'b', [70])],
    });
    const s = scene(m);
    const tx = texts(s).map((t) => t.text);
    expect(tx).toContain('-100%');
    expect(tx).toContain('100%');
    const [neg, pos] = fillPaths(s).map((p) => rectsOf(p.d)[0]!);
    // 範囲 −1〜1（2）：プロットの高さ 180 に対し、負（−0.3）は 0.3/2・正（0.7）は 0.7/2
    expect(neg!.h).toBeCloseTo((180 * 0.3) / 2, 1);
    expect(pos!.h).toBeCloseTo((180 * 0.7) / 2, 1);
    expect(neg!.y).toBeCloseTo(pos!.y + pos!.h, 1);
  });

  it('TD-102-91: nextTo（無指定）の項目軸の文字は軸の線の隣（負を含む縦棒）・low は下端・high は上端', () => {
    const base = { series: [ser(0, 's', [-10, 20, 30])] };
    const nx = scene(plotted(base));
    const lineY = axisLine(nx).y1;
    const plot = plotRect(nx);
    const a = labelOf(nx, 'A');
    expect(a.y).toBeCloseTo(lineY + TICK_GAP, 5);
    expect(a.y).toBeLessThan(plot.y + plot.height); // プロットの内側
    expect(a.baseline).toBe('hanging');
    const low = scene(plotted({ ...base, catAxis: { ...catAx, tickLblPos: 'low' } }));
    const lowPlot = plotRect(low);
    expect(labelOf(low, 'A').y).toBeCloseTo(lowPlot.y + lowPlot.height + TICK_GAP, 5);
    const high = scene(plotted({ ...base, catAxis: { ...catAx, tickLblPos: 'high' } }));
    const highPlot = plotRect(high);
    expect(labelOf(high, 'A').y).toBeCloseTo(highPlot.y - TICK_GAP, 5);
    expect(labelOf(high, 'A').baseline).toBe('alphabetic');
  });

  it('TD-102-91: nextTo は線が中ほどにある時は余白を取らない（low より縦に広い）', () => {
    const base = { series: [ser(0, 's', [-10, 20, 30])] };
    const nx = plotRect(scene(plotted(base)));
    const low = plotRect(scene(plotted({ ...base, catAxis: { ...catAx, tickLblPos: 'low' } })));
    expect(nx.height).toBeGreaterThan(low.height);
    // 線が下端にある（正だけ）なら nextTo でも余白を取り、low と同じ
    const pos = { series: [ser(0, 's', [10, 20, 30])] };
    const a = plotRect(scene(plotted(pos)));
    const b = plotRect(scene(plotted({ ...pos, catAxis: { ...catAx, tickLblPos: 'low' } })));
    expect(a.height).toBeCloseTo(b.height, 5);
  });

  it('TD-102-91: 負を含む横棒：項目軸の文字は軸の線の隣（左に右揃え）・low は左端・high は右端', () => {
    const base = { barDir: 'bar' as const, series: [ser(0, 's', [-10, 20, 30])] };
    const nx = scene(plotted(base));
    const lineX = axisLine(nx).x1;
    const a = labelOf(nx, 'A');
    expect(a.x).toBeCloseTo(lineX - TICK_GAP, 5);
    expect(a.anchor).toBe('end');
    const plot = plotRect(nx);
    expect(a.x).toBeGreaterThan(plot.x);
    const low = scene(plotted({ ...base, catAxis: { ...catAx, tickLblPos: 'low' } }));
    expect(labelOf(low, 'A').x).toBeCloseTo(plotRect(low).x - TICK_GAP, 5);
    const high = scene(plotted({ ...base, catAxis: { ...catAx, tickLblPos: 'high' } }));
    const hp = plotRect(high);
    expect(labelOf(high, 'A').x).toBeCloseTo(hp.x + hp.width + TICK_GAP, 5);
    expect(labelOf(high, 'A').anchor).toBe('start');
    expect(plotRect(nx).width).toBeGreaterThan(plotRect(low).width);
  });

  it('TD-102-92: 軸タイトルの文字の大きさは軸ごと（数値軸 14pt・項目軸は既定 10pt）', () => {
    const m = plotted({
      valAxis: { ...valAx, title: { lines: ['V'], text: { sizePt: 14 }, overlay: false } },
      catAxis: { ...catAx, title: { lines: ['CT'], text: {}, overlay: false } },
    });
    const s = scene(m);
    expect(labelOf(s, 'V').size).toBeCloseTo(14 * PX_PER_PT, 10);
    expect(labelOf(s, 'CT').size).toBeCloseTo(10 * PX_PER_PT, 10);
    // 数値軸のタイトルが大きいほど、左の帯が広がってプロットが狭い
    const small = plotted({
      valAxis: { ...valAx, title: { lines: ['V'], text: { sizePt: 10 }, overlay: false } },
    });
    expect(plotRect(s).x).toBeGreaterThan(plotRect(scene(small)).x);
    // 横棒では縦の軸が項目軸
    const h = scene({
      ...m,
      barDir: 'bar',
      catAxis: { ...catAx, title: { lines: ['CT'], text: { sizePt: 16 }, overlay: false } },
    });
    expect(labelOf(h, 'CT').size).toBeCloseTo(16 * PX_PER_PT, 10);
    expect(labelOf(h, 'V').size).toBeCloseTo(14 * PX_PER_PT, 10);
  });

  it('TD-102-93: 円の上限判定は系列数を 1 として数える（ドーナツは全系列）', () => {
    const cats = Array.from({ length: 3000 }, (_, i) => `c${i}`);
    const ones = cats.map(() => 1);
    const two = [ser(0, 'a', ones), ser(1, 'b', ones)];
    expect(layoutChart(pie({ categories: cats, series: two }), 400, 300)).not.toBeNull();
    expect(
      layoutChart(pie({ type: 'doughnut', categories: cats, series: two }), 400, 300),
    ).toBeNull();
  });

  it('TD-102-94: 横棒の数値軸の目盛の数は最長の目盛の文字の幅で決まる（縦棒より少ない）', () => {
    const big = [ser(0, 's', [10_000_000, 20_000_000, 30_000_000])];
    const numFmt = { code: '#,##0', sourceLinked: false };
    const count = (m: ChartModel): number =>
      texts(scene(m)).filter((t) => /^\d{1,3}(,\d{3})+$/.test(t.text)).length;
    const horizontal = count(mk({ barDir: 'bar', series: big, valAxis: { ...valAx, numFmt } }));
    const vertical = count(mk({ series: big, valAxis: { ...valAx, numFmt } }));
    expect(horizontal).toBeGreaterThanOrEqual(2);
    expect(horizontal).toBeLessThanOrEqual(5);
    expect(vertical).toBeGreaterThan(horizontal);
  });

  it('TD-102-95: 横棒で項目軸が maxMin：帯の中の系列は上から並ぶ（既定は下から）', () => {
    const two = [ser(0, 'a', [10, 20, 30]), ser(1, 'b', [10, 20, 30])];
    const inner = innerPlot(0.1, 0.1, 0.75, 0.6);
    const ys = (m: ChartModel): number[] => fillPaths(scene(m)).map((p) => rectsOf(p.d)[0]!.y);
    const normal = ys(mk({ ...inner, barDir: 'bar', series: two }));
    expect(normal[0]).toBeGreaterThan(normal[1]!); // a は b の下
    const rev = ys(
      mk({ ...inner, barDir: 'bar', series: two, catAxis: { ...catAx, orientation: 'maxMin' } }),
    );
    expect(rev[0]).toBeLessThan(rev[1]!); // a は b の上
  });

  it('TD-102-96: 積み上げの折れ線で gap の null：その系列の線が切れる', () => {
    const m = mk({
      type: 'line',
      barDir: undefined,
      grouping: 'stacked',
      categories: ['A', 'B', 'C', 'D', 'E'],
      dispBlanksAs: 'gap',
      series: [ser(0, 'a', [10, 10, 10, 10, 10]), ser(1, 'b', [10, 10, null, 10, 10])],
    });
    const lines = paths(scene(m)).filter(
      (p) => p.fill === null && p.stroke?.width === 1.5 * PX_PER_PT,
    );
    expect((lines[0]!.d.match(/M/g) ?? []).length).toBe(1);
    expect((lines[1]!.d.match(/M/g) ?? []).length).toBe(2); // 2 区間
  });

  it('TD-102-97: 数値軸の crosses=max と項目軸の maxMin の組み合わせで数値軸の線の側が決まる', () => {
    const line = { paint: solid(0x112233) };
    const lineX = (valCross: boolean, catRev: boolean): { x: number; plot: ChartRectItem } => {
      const s = scene(
        plotted({
          valAxis: { ...valAx, line, ...(valCross ? { crosses: 'max' as const } : {}) },
          catAxis: { ...catAx, orientation: catRev ? 'maxMin' : 'minMax' },
        }),
      );
      const l = s.items.find((i) => i.kind === 'line' && i.stroke.color === '#112233')!;
      if (l.kind !== 'line') throw new Error('line ではない');
      return { x: l.x1, plot: plotRect(s) };
    };
    const a = lineX(false, false);
    expect(a.x).toBeCloseTo(a.plot.x, 5); // 左
    const b = lineX(true, false);
    expect(b.x).toBeCloseTo(b.plot.x + b.plot.width, 5); // 右
    const c = lineX(false, true);
    expect(c.x).toBeCloseTo(c.plot.x + c.plot.width, 5); // 反転で右
    const d = lineX(true, true);
    expect(d.x).toBeCloseTo(d.plot.x, 5); // 両方で左に戻る
  });

  it('TD-102-98: 補助の目盛線（minorGridlines）は主の目盛と重ならない位置に薄い色で引く', () => {
    const s = scene(mk({ valAxis: { ...valAx, minorGridlines: {} } }));
    const minor = paths(s).find((p) => p.stroke?.color === '#bcbcbc')!;
    expect(minor.stroke!.width).toBeCloseTo(0.5 * PX_PER_PT, 10);
    // 0〜35・主 5・補助 1：36 本から主の 8 本を引いた 28 本
    expect((minor.d.match(/M/g) ?? []).length).toBe(28);
    const major = paths(s).find((p) => p.stroke?.color === '#000000')!;
    expect(s.items.indexOf(minor)).toBeLessThan(s.items.indexOf(major)); // 補助が下
    expect(paths(scene(mk())).some((p) => p.stroke?.color === '#bcbcbc')).toBe(false);
  });

  it('TD-102-99: 円の凡例で削除の idx は項目の位置（c:legendEntry c:idx＝点の位置）', () => {
    const s = scene(pie({ legend: { pos: 'r', overlay: false, deleted: [1], text: {} } }));
    const tx = texts(s).map((t) => t.text);
    expect(tx).toContain('A');
    expect(tx).not.toContain('B');
    expect(tx).toContain('C');
  });
});

describe('ISSUE-104 レビュー指摘（範囲外の値・自動タイトルの帯・円の varyColors・横棒の右の余白）', () => {
  const inner = innerPlot(0.1, 0.1, 0.75, 0.6); // 400×300 → プロット x 40..340・y 30..210（300×180・帯 100）
  const numbers = (d: string): number[] => (d.match(/-?\d+(?:\.\d+)?/g) ?? []).map(Number);
  const line = (over: Partial<ChartModel> = {}): ChartModel =>
    mk({
      type: 'line',
      barDir: undefined,
      grouping: 'standard',
      categories: ['A', 'B', 'C', 'D'],
      ...inner,
      valAxis: { ...valAx, min: 0, max: 15 },
      ...over,
    });
  const isLine = (p: ChartPathItem): boolean =>
    p.fill === null && p.stroke?.width === 1.5 * PX_PER_PT;
  const linePath = (s: ChartScene): ChartPathItem => paths(s).find(isLine)!;

  it('TD-104-30: 軸の範囲を出る折れ線は、プロットの矩形で切る（端に寄せて傾きを変えない）・外を通る区間で複数の経路になる', () => {
    // 点 A=10（中）・B=30（上へ外れる）・C=10・D=5。プロット y 30..210・帯の中心 x=77.5, 152.5, 227.5, 302.5
    const d = linePath(scene(line({ series: [ser(0, 's', [10, 30, 10, 5])] }))).d;
    const ns = numbers(d);
    // すべての座標がプロットの内側
    for (let i = 0; i < ns.length; i += 2) {
      expect(ns[i]!).toBeGreaterThanOrEqual(40 - 1e-6);
      expect(ns[i]!).toBeLessThanOrEqual(340 + 1e-6);
      expect(ns[i + 1]!).toBeGreaterThanOrEqual(30 - 1e-6);
      expect(ns[i + 1]!).toBeLessThanOrEqual(210 + 1e-6);
    }
    expect((d.match(/M/g) ?? []).length).toBe(2); // A→上端で切れる／上端から入り直して C→D
    // A(77.5, 90) → B(152.5, -150) の傾きのまま、y=30 で切れる：t=0.25 → x = 77.5 + 75 × 0.25 = 96.25
    expect(ns.slice(0, 4)).toEqual([77.5, 90, 96.25, 30]);
    // 入り直す点：B(152.5, -150) → C(227.5, 90)。y=30 は t=0.75 → x = 208.75
    expect(ns.slice(4, 8)).toEqual([208.75, 30, 227.5, 90]);
  });

  it('TD-104-30: 範囲内なら 1 本の経路のまま・プロットの外の点のマーカーは出さない', () => {
    const inside = linePath(scene(line({ series: [ser(0, 's', [10, 12, 10, 5])] })));
    expect((inside.d.match(/M/g) ?? []).length).toBe(1);
    const s = scene(line({ lineMarkers: true, series: [ser(0, 's', [10, 30, 10, 5])] }));
    const mp = paths(s).find((p) => p.fill?.color === '#4472c4' && p.d.includes('A'))!;
    expect((mp.d.match(/M/g) ?? []).length).toBe(3); // B は範囲外
  });

  it('TD-104-30: 全部が範囲外なら線の経路を作らない・滑らかな線も切る（折れ線に直す）', () => {
    expect(paths(scene(line({ series: [ser(0, 's', [20, 30, 40, 50])] }))).some(isLine)).toBe(
      false,
    );
    const smooth = linePath(
      scene(line({ series: [ser(0, 's', [10, 30, 10, 5], { smooth: true })] })),
    );
    const ns = numbers(smooth.d);
    expect(smooth.d).not.toContain('C');
    for (let i = 0; i < ns.length; i += 2) {
      expect(ns[i + 1]!).toBeGreaterThanOrEqual(30 - 1e-6);
      expect(ns[i + 1]!).toBeLessThanOrEqual(210 + 1e-6);
    }
    expect((smooth.d.match(/M/g) ?? []).length).toBe(2);
  });

  it('TD-104-31: 範囲から完全に外れた棒は出さない・かかる棒は端に収める', () => {
    // 積み上げ 2 系列（各 10）・軸 0〜8：系列 1 は 0..10 が端に収まり、系列 2 は 10..20 で完全に外れる
    const m = mk({
      ...inner,
      grouping: 'stacked',
      valAxis: { ...valAx, min: 0, max: 8 },
      series: [ser(0, 'a', [10, 10, 10]), ser(1, 'b', [10, 10, 10])],
    });
    const fp = fillPaths(scene(m)).filter((p) => p.fill!.color !== '#123456');
    expect(fp).toHaveLength(1);
    const r = rectsOf(fp[0]!.d);
    expect(r).toHaveLength(3);
    expect(Math.min(...r.map((x) => x.y))).toBeGreaterThanOrEqual(30 - 1e-9);
    expect(r.every((x) => Math.abs(x.h - 180) < 1e-6)).toBe(true); // 端まで
  });

  it('TD-104-32: 文字の無い c:title・系列 2 つ以上は、帯だけ取り文字は描かない', () => {
    const plotY = (m: Partial<ChartModel>): number =>
      plotRect(scene(mk({ plotArea: { fill: solid(0x123456) }, ...m }))).y;
    const two = [ser(0, 'a', [1, 2, 3]), ser(1, 'b', [1, 2, 3])];
    const blank = scene(mk({ title: { lines: [], text: {}, overlay: false }, series: two }));
    expect(
      texts(blank).filter((t) => t.anchor === 'middle' && t.baseline === 'middle'),
    ).toHaveLength(0);
    const none = plotY({ series: two });
    const withBlank = plotY({ title: { lines: [], text: {}, overlay: false }, series: two });
    const withT = plotY({ title: { lines: ['T'], text: {}, overlay: false }, series: two });
    expect(withBlank).toBeGreaterThan(none);
    expect(withBlank).toBeCloseTo(withT, 10); // 1 行の高さの帯
    // overlay なら帯を取らない
    expect(plotY({ title: { lines: [], text: {}, overlay: true }, series: two })).toBeCloseTo(
      none,
      10,
    );
  });

  it('TD-104-33: 円・ドーナツで varyColors が偽なら 1 色（系列の色）・c:dPt の明示は優先・凡例は系列', () => {
    const cats = ['a', 'b', 'c', 'd'];
    const flat = pie({ varyColors: false, categories: cats, series: [ser(0, 'S', [1, 1, 1, 1])] });
    expect(fillPaths(scene(flat)).map((p) => p.fill!.color)).toEqual(
      Array<string>(4).fill(hexOf(0x4472c4)),
    );
    const dpt = scene({
      ...flat,
      series: [ser(0, 'S', [1, 1, 1, 1], { points: new Map([[2, solid(0xabcdef)]]) })],
    });
    expect(fillPaths(dpt).map((p) => p.fill!.color)).toEqual(
      [0x4472c4, 0x4472c4, 0xabcdef, 0x4472c4].map(hexOf),
    );
    // ドーナツ：系列ごとに 1 色（輪ごと）
    const ring = scene(
      pie({
        type: 'doughnut',
        varyColors: false,
        categories: ['a', 'b', 'c'],
        series: [ser(0, 'S1', [1, 1, 1]), ser(1, 'S2', [1, 1, 1])],
      }),
    );
    const colors = fillPaths(ring).map((p) => p.fill!.color);
    expect(colors.slice(0, 3)).toEqual(Array<string>(3).fill(hexOf(0x4472c4)));
    expect(colors.slice(3)).toEqual(Array<string>(3).fill(hexOf(0xed7d31)));
    // 凡例：ドーナツは偽なら系列名の 1 項目・真なら項目ごと。円は偽でも項目ごと（C3。Excel 16 の COM は 5 項目を返す）
    const legend = { pos: 'r' as const, overlay: false, deleted: [], text: {} };
    const entries = (m: ChartModel): string[] =>
      texts(scene(m))
        .filter((t) => t.anchor === 'start' && t.baseline === 'middle')
        .map((t) => t.text);
    expect(entries({ ...flat, legend })).toEqual(cats); // 円は varyColors が偽でも項目ごと（色は全部が系列の色）
    expect(
      entries({
        ...flat,
        type: 'doughnut',
        legend,
      }),
    ).toEqual(['S']); // ドーナツは偽なら系列ごと
    expect(entries(pie({ categories: cats, legend }))).toEqual(cats);
  });

  it('TD-104-34: 横棒で項目軸が右（crosses=max）のとき、右の余白は項目名の帯と目盛の半幅の大きい方（足さない）', () => {
    const m = mk({
      barDir: 'bar',
      plotArea: { fill: solid(0x123456) },
      catAxis: { ...catAx, crosses: 'max' },
    });
    const r = plotRect(scene(m));
    const band = estimateTextWidth('A', 10 * PX_PER_PT) + 5 * PX_PER_PT + 3; // 項目名の幅＋目盛＋隙間
    expect(r.x + r.width).toBeCloseTo(400 - 11 * PX_PER_PT - band, 5);
  });
});

describe('ISSUE-105 照合の修正（目盛の印の既定・凡例の並び・全部が負の項目軸・マーカー）', () => {
  const lg = (pos: 'r' | 'l' | 't' | 'b' | 'tr', deleted: number[] = []): ChartModel['legend'] => ({
    pos,
    overlay: false,
    deleted,
    text: {},
  });
  const two = [ser(0, 'a', [1, 2, 3]), ser(1, 'b', [3, 2, 1])];
  /** 凡例の文字（見た目の順） */
  const legendTexts = (m: ChartModel): string[] =>
    texts(scene(m))
      .filter((t) => t.anchor === 'start' && t.baseline === 'middle')
      .map((t) => t.text);

  it('TD-105-1: c:majorTickMark が無い項目軸は交差の印（項目の境目の本数×両側）。none なら印なし（C12）', () => {
    const ticks = (m: ChartModel): number =>
      paths(scene(m))
        .filter((p) => p.fill === null && p.stroke?.color === '#000000' && !p.d.includes('A'))
        .map((p) => (p.d.match(/M/g) ?? []).length)
        .filter((n) => n === 8).length; // 3 項目＝境目 4 つ×両側 2
    expect(
      ticks(
        mk({
          valAxis: {
            ...valAx,
            majorGridlines: undefined,
            minorTickMark: 'none',
            majorTickMark: 'none',
          },
        }),
      ),
    ).toBe(1);
    expect(
      ticks(
        mk({
          catAxis: { ...catAx, majorTickMark: 'none' },
          valAxis: {
            ...valAx,
            majorGridlines: undefined,
            minorTickMark: 'none',
            majorTickMark: 'none',
          },
        }),
      ),
    ).toBe(0);
  });

  it('TD-105-2: 凡例の並びは、縦並びで積み上げの縦棒・集合の横棒だけ逆（積み上げの横棒・集合の縦棒・下の凡例は系列の順）。削除の idx は表示の並びの位置（C11）', () => {
    const col = { type: 'bar', barDir: 'col', series: two } as const;
    const bar = { type: 'bar', barDir: 'bar', series: two } as const;
    expect(legendTexts(mk({ ...col, grouping: 'clustered', legend: lg('r') }))).toEqual(['a', 'b']);
    expect(legendTexts(mk({ ...col, grouping: 'stacked', legend: lg('r') }))).toEqual(['b', 'a']);
    expect(legendTexts(mk({ ...col, grouping: 'percentStacked', legend: lg('l') }))).toEqual([
      'b',
      'a',
    ]);
    expect(legendTexts(mk({ ...bar, grouping: 'clustered', legend: lg('tr') }))).toEqual([
      'b',
      'a',
    ]);
    expect(legendTexts(mk({ ...bar, grouping: 'stacked', legend: lg('r') }))).toEqual(['a', 'b']);
    expect(legendTexts(mk({ ...col, grouping: 'stacked', legend: lg('b') }))).toEqual(['a', 'b']); // 横並びは逆にしない
    expect(legendTexts(mk({ ...col, grouping: 'stacked', legend: lg('r', [0]) }))).toEqual(['a']);
  });

  it('TD-105-3: 全部が負の 100% 積み上げ：項目軸（プロットの上端）の文字は軸の線の下（内側）に置き、上に余白を取らない（C2）', () => {
    const neg = mk({
      grouping: 'percentStacked',
      plotArea: { fill: solid(0x123456) },
      series: [ser(0, 'a', [-1, -2, -3]), ser(1, 'b', [-2, -2, -2])],
    });
    const pos = mk({
      grouping: 'percentStacked',
      plotArea: { fill: solid(0x123456) },
      series: [ser(0, 'a', [1, 2, 3]), ser(1, 'b', [2, 2, 2])],
    });
    const rn = plotRect(scene(neg));
    const rp = plotRect(scene(pos));
    expect(rn.y).toBeCloseTo(rp.y, 5); // 上に帯を取らない
    const label = texts(scene(neg)).find((t) => t.text === 'A')!;
    expect(label.y).toBeGreaterThan(rn.y);
    expect(label.y).toBeLessThan(rn.y + rn.height);
    // 数値軸の反転で端に来る時は従来どおり外側（上）
    const rev = mk({
      plotArea: { fill: solid(0x123456) },
      valAxis: { ...valAx, orientation: 'maxMin' },
    });
    const lr = texts(scene(rev)).find((t) => t.text === 'A')!;
    expect(lr.y).toBeLessThan(plotRect(scene(rev)).y);
  });

  it('TD-105-4: 折れ線の系列にマーカーの指定が無ければ、グループの c:marker が無くても 0 でも自動の円（symbol none だけ出ない・C13）', () => {
    const line = (over: Partial<ChartModel>): ChartModel =>
      mk({ type: 'line', barDir: undefined, grouping: 'standard', ...over });
    const has = (m: ChartModel): boolean => paths(scene(m)).some((p) => p.d.includes('A'));
    expect(has(line({ lineMarkers: undefined, series: [ser(0, 's', [1, 2, 3])] }))).toBe(true);
    expect(has(line({ lineMarkers: false, series: [ser(0, 's', [1, 2, 3])] }))).toBe(true);
    expect(has(line({ series: [ser(0, 's', [1, 2, 3], { marker: { symbol: 'none' } })] }))).toBe(
      false,
    );
  });
});

// ---- 点の多いグラフの間引き（ISSUE-106・ADR-0031 §決定 11.1） ----
describe('点の多いグラフの間引き（ISSUE-106・§決定 11.1）', () => {
  const W = 600;
  const H = 360;
  const DENSE_POINTS = 1000; // 600px に 1,000 点＝項目の間隔 約 0.5px
  const cats = (n: number): string[] => Array.from({ length: n }, (_, i) => `c${i}`);
  const wave = (s: number, n: number): number[] =>
    Array.from({ length: n }, (_, i) => 50 + 30 * Math.sin((i + s * 7) / 40) + s * 3);
  const denseLine = (
    over: Partial<ChartModel> = {},
    serOver: Partial<ChartSeries> = {},
  ): ChartModel =>
    mk({
      type: 'line',
      barDir: undefined,
      grouping: 'standard',
      categories: cats(DENSE_POINTS),
      series: Array.from({ length: 10 }, (_, s) => ser(s, `s${s}`, wave(s, DENSE_POINTS), serOver)),
      ...over,
    });
  const denseBar = (over: Partial<ChartModel> = {}): ChartModel =>
    mk({
      categories: cats(DENSE_POINTS),
      series: Array.from({ length: 10 }, (_, s) => ser(s, `s${s}`, wave(s, DENSE_POINTS))),
      ...over,
    });
  const off = (m: ChartModel): ChartScene => {
    const s = layoutChart(m, W, H, undefined, { thin: false });
    if (s === null) throw new Error('場面が null');
    return s;
  };
  /** 系列の線（round の端・塗りなし・太さ 1.5pt）の経路 */
  const lineItems = (s: ChartScene): ChartPathItem[] =>
    paths(s).filter(
      (p) => p.fill === null && p.stroke?.cap === 'round' && p.stroke.width === 1.5 * PX_PER_PT,
    );
  /** 経路の数値の組を点列にする（`M x y L x y …`） */
  const points = (d: string): [number, number][] => {
    const out: [number, number][] = [];
    for (const m of d.matchAll(/[ML](-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?)/g))
      out.push([Number(m[1]), Number(m[2])]);
    return out;
  };
  const distToPoly = (p: readonly [number, number], poly: readonly [number, number][]): number => {
    let best = Infinity;
    for (let i = 0; i + 1 < poly.length; i++) {
      const [ax, ay] = poly[i]!;
      const [bx, by] = poly[i + 1]!;
      const dx = bx - ax;
      const dy = by - ay;
      const l2 = dx * dx + dy * dy;
      const t = l2 === 0 ? 0 : Math.max(0, Math.min(1, ((p[0] - ax) * dx + (p[1] - ay) * dy) / l2));
      best = Math.min(best, Math.hypot(p[0] - ax - t * dx, p[1] - ay - t * dy));
    }
    return best;
  };
  const onlyKnownCommands = (d: string): boolean => /^[MLCAZ0-9. -]*$/.test(d);

  it('TD-106-10: 密でないグラフは thin: false と場面が完全に同じ（棒・折れ線・滑らか・マーカー・円）', () => {
    const lineN = (
      n: number,
      over: Partial<ChartSeries> = {},
      mo: Partial<ChartModel> = {},
    ): ChartModel =>
      mk({
        type: 'line',
        barDir: undefined,
        grouping: 'standard',
        categories: cats(n),
        series: [ser(0, 's', wave(0, n), over), ser(1, 't', wave(1, n), over)],
        ...mo,
      });
    const models: ChartModel[] = [
      mk(),
      mk({ categories: cats(200), series: [ser(0, 's', wave(0, 200))] }), // 間隔 約 2.5px
      mk({ barDir: 'bar', categories: cats(100), series: [ser(0, 's', wave(0, 100))] }),
      lineN(150),
      lineN(150, { smooth: true }),
      lineN(150, { marker: { symbol: 'circle' } }),
      lineN(150, { marker: { symbol: 'square', size: 4 } }),
      pie({ categories: cats(5), series: [ser(0, 's', [1, 2, 3, 4, 5])] }),
    ];
    for (const m of models) {
      expect(layoutChart(m, W, H, undefined, { thin: true })).toEqual(
        layoutChart(m, W, H, undefined, { thin: false }),
      );
      expect(layoutChart(m, W, H)).toEqual(layoutChart(m, W, H, undefined, { thin: false }));
    }
  });

  it('TD-106-11: 円のグラフは扇が細く続かなければ（太い扇）、thin の有無で同じ', () => {
    const m = pie({ categories: cats(20), series: [ser(0, 's', wave(0, 20))] });
    expect(layoutChart(m, W, H)).toEqual(layoutChart(m, W, H, undefined, { thin: false }));
  });

  it('TD-106-12: 密な折れ線 10×1,000：最初と最後が残り、元の全点が間引いた線から 0.5px 以内・d が大きく減る', () => {
    const m = denseLine();
    const before = lineItems(off(m));
    const after = lineItems(scene(m, W, H));
    expect(before).toHaveLength(10);
    expect(after).toHaveLength(10);
    let nBefore = 0;
    let nAfter = 0;
    for (let k = 0; k < 10; k++) {
      const a = points(before[k]!.d);
      const b = points(after[k]!.d);
      nBefore += a.length;
      nAfter += b.length;
      expect(b[0]).toEqual(a[0]);
      expect(b[b.length - 1]).toEqual(a[a.length - 1]);
      for (const p of a) expect(distToPoly(p, b)).toBeLessThanOrEqual(0.5);
      expect(after[k]!.stroke).toEqual(before[k]!.stroke);
    }
    expect(nBefore).toBe(10_000);
    expect(nAfter).toBeLessThan(nBefore / 4);
  });

  it('TD-106-13: 正弦波は頂点が大きく減る（1 系列 1,000 点 → 1/5 以下）', () => {
    const m = denseLine({ series: [ser(0, 's', wave(0, DENSE_POINTS))] });
    const a = points(lineItems(off(m))[0]!.d);
    const b = points(lineItems(scene(m, W, H))[0]!.d);
    expect(a).toHaveLength(1000);
    expect(b.length).toBeLessThan(200); // 128 前後
    expect(b.length).toBeGreaterThan(5);
  });

  it('TD-106-14: 項目が逆順（c:orientation maxMin）でも最初と最後が残る', () => {
    for (const m of [
      denseLine({ catAxis: { ...catAx, orientation: 'maxMin' } }),
      denseLine({
        catAxis: { ...catAx, orientation: 'maxMin' },
        series: [ser(0, 's', wave(0, DENSE_POINTS))],
      }),
    ]) {
      const a = points(lineItems(off(m))[0]!.d);
      const b = points(lineItems(scene(m, W, H))[0]!.d);
      expect(b[0]).toEqual(a[0]);
      expect(b[b.length - 1]).toEqual(a[a.length - 1]);
      expect(b.length).toBeLessThan(a.length);
      for (const p of a) expect(distToPoly(p, b)).toBeLessThanOrEqual(0.5);
    }
  });

  it('TD-106-15: 滑らかな線も密なら折れ線（C を出さない）で間引く・密でない時は C のまま', () => {
    const dense = denseLine({}, { smooth: true });
    for (const p of lineItems(scene(dense, W, H))) {
      expect(p.d).not.toContain('C');
      expect(onlyKnownCommands(p.d)).toBe(true);
    }
    const sparse = mk({
      type: 'line',
      barDir: undefined,
      grouping: 'standard',
      series: [ser(0, 's', [10, 30, 20], { smooth: true })],
    });
    expect(lineItems(scene(sparse))[0]!.d).toContain('C');
  });

  it('TD-106-16: 空の点で切れた線は部分ごとに間引く（部分の数は変わらない）', () => {
    const vals = wave(0, DENSE_POINTS).map((v, i) => (i >= 400 && i < 410 ? null : v));
    const m = denseLine({ series: [ser(0, 's', vals)] });
    const a = lineItems(off(m))[0]!.d;
    const b = lineItems(scene(m, W, H))[0]!.d;
    expect((a.match(/M/g) ?? []).length).toBe(2);
    expect((b.match(/M/g) ?? []).length).toBe(2);
    expect(points(b).length).toBeLessThan(points(a).length / 4);
  });

  it('TD-106-17: 円のマーカー（既定の見た目）は太さ 2R・round の線 1 本の帯になる', () => {
    const size = 5;
    const m = denseLine({}, { marker: { symbol: 'circle', size } });
    const r = (size * PX_PER_PT) / 2;
    const R = r + (CHART_DEFAULTS.thinLinePt * PX_PER_PT) / 2;
    const before = off(m);
    const after = scene(m, W, H);
    // 帯：太さ 2R・round の塗りなしの線（系列の線と見分けるため太さで探す）
    const isBand = (p: ChartPathItem): boolean =>
      p.fill === null && p.stroke !== null && Math.abs(p.stroke.width - 2 * R) < 1e-9;
    const bands = paths(after).filter(isBand);
    expect(bands).toHaveLength(10); // 系列ごとに 1 本
    for (const b of bands) {
      expect(b.stroke!.cap).toBe('round');
      expect(b.stroke!.join).toBe('round');
      expect(b.stroke!.opacity).toBe(1);
      expect((b.d.match(/M/g) ?? []).length).toBe(1);
      expect(onlyKnownCommands(b.d)).toBe(true);
    }
    // 円の経路（A は 1 つの円に 2 つ）は帯に置き換わる
    const circles = (s: ChartScene): number =>
      paths(s).reduce((n, p) => n + (p.d.match(/A/g) ?? []).length / 2, 0);
    expect(circles(before)).toBe(10_000);
    expect(circles(after)).toBe(0);
  });

  it('TD-106-18: 四角のマーカーは省いた点が残した点から 1px 以内・間引かれる', () => {
    const m = denseLine({
      series: [ser(0, 's', wave(0, DENSE_POINTS), { marker: { symbol: 'square', size: 4 } })],
    });
    const squares = (s: ChartScene): { x: number; y: number }[] =>
      paths(s)
        .filter((p) => p.fill !== null)
        .flatMap((p) => rectsOf(p.d))
        .map((q) => ({ x: q.x + q.w / 2, y: q.y + q.h / 2 }));
    const a = squares(off(m));
    const b = squares(scene(m, W, H));
    expect(a).toHaveLength(1000);
    expect(b.length).toBeLessThan(a.length); // 急な傾きでは 1px の間隔が詰まらず、あまり減らない
    expect(b.length).toBeGreaterThan(10);
    for (const p of a) {
      const near = Math.min(...b.map((q) => Math.hypot(p.x - q.x, p.y - q.y)));
      expect(near).toBeLessThanOrEqual(1 + 0.02); // 経路の丸め（小数 2 桁）の分
    }
  });

  it('TD-106-18: 平らな系列の四角のマーカーは約半分になる', () => {
    const m = denseLine({
      series: [
        ser(
          0,
          's',
          Array.from({ length: DENSE_POINTS }, () => 50),
          { marker: { symbol: 'square', size: 4 } },
        ),
      ],
    });
    const count = (s: ChartScene): number =>
      paths(s)
        .filter((p) => p.fill !== null)
        .reduce((n, p) => n + rectsOf(p.d).length, 0);
    expect(count(off(m))).toBe(1000);
    expect(count(scene(m, W, H))).toBeLessThan(600);
  });

  it('TD-106-19: c:dPt で上書きした点のマーカーは間引かない（帯にも入れない）', () => {
    const color = solid(0x123456);
    const m = denseLine({
      series: [
        ser(0, 's', wave(0, DENSE_POINTS), {
          marker: { symbol: 'circle', size: 5, fill: color },
          points: new Map([
            [500, color],
            [501, color],
          ]),
        }),
      ],
    });
    const before = paths(off(m)).filter((p) => p.d.includes('A'));
    expect(before).toHaveLength(1);
    const parts = before[0]!.d
      .split('Z')
      .filter((x) => x !== '')
      .map((x) => `${x}Z`);
    const kept = paths(scene(m, W, H)).filter((p) => p.d.includes('A'));
    expect(kept).toHaveLength(1);
    expect(kept[0]!.d).toContain(parts[500]!);
    expect(kept[0]!.d).toContain(parts[501]!);
    expect(kept[0]!.d.split('Z').length - 1).toBeLessThan(10); // 他の点は帯へ（単独に残るのは帯に入らない点だけ）
  });

  it('TD-106-20: 密な棒 10×1,000：系列ごとの矩形の数は列の数以下・幅の合計と値の範囲が保たれる', () => {
    const m = denseBar();
    const bars = (s: ChartScene): ChartPathItem[] =>
      fillPaths(s).filter((p) => p.stroke === null && rectsOf(p.d).length > 3);
    const b0 = bars(off(m));
    const b1 = bars(scene(m, W, H));
    expect(b0).toHaveLength(10);
    expect(b1).toHaveLength(10);
    const plotW = 600; // 高々 600px の幅＝列の数の上限
    for (let k = 0; k < 10; k++) {
      const ra = rectsOf(b0[k]!.d);
      const rb = rectsOf(b1[k]!.d);
      expect(ra).toHaveLength(1000);
      expect(rb.length).toBeLessThanOrEqual(plotW);
      expect(b1[k]!.fill).toEqual(b0[k]!.fill);
      const wa = ra.reduce((s, r) => s + r.w, 0);
      const wb = rb.reduce((s, r) => s + r.w, 0);
      expect(Math.abs(wa - wb) / wa).toBeLessThan(0.03);
      const top = (rs: { y: number }[]): number => Math.min(...rs.map((r) => r.y));
      const bottom = (rs: { y: number; h: number }[]): number =>
        Math.max(...rs.map((r) => r.y + r.h));
      expect(top(rb)).toBeCloseTo(top(ra), 1);
      expect(bottom(rb)).toBeCloseTo(bottom(ra), 1);
    }
  });

  it('TD-106-21: 密な縦棒・横棒・積み上げ棒は矩形が減る（間隔 0.5px の縦棒で約半分・横棒は約 1/3）・正負の棒は列ごとに 2 つまで', () => {
    const neg = (s: number): number[] =>
      wave(s, DENSE_POINTS).map((v, i) => (i % 2 === 0 ? v : -v));
    const count = (s: ChartScene): number =>
      fillPaths(s).reduce((n, p) => n + (p.d.match(/Z/g) ?? []).length, 0);
    for (const m of [denseBar(), denseBar({ barDir: 'bar' }), denseBar({ grouping: 'stacked' })]) {
      expect(count(scene(m, W, H))).toBeLessThan(count(off(m)) / 1.5);
    }
    // 正と負の値は列の中でも別の矩形（0 をまたいで 1 つにしない）：列ごとに 正と負の最大 2 つ
    const m = denseBar({ series: [ser(0, 's', neg(0))] });
    const rs = fillPaths(scene(m, W, H)).flatMap((p) => rectsOf(p.d));
    expect(rs.length).toBeGreaterThan(300);
    expect(rs.length).toBeLessThanOrEqual(2 * 600);
    // 値の方向：正の棒は 0 の線より上・負の棒は下（またいだ矩形が無い）
    const zeroY = rectsOf(fillPaths(off(m))[0]!.d).find((_, i) => i === 0)!;
    const base = zeroY.y + zeroY.h; // 最初の棒（正）の下端＝値 0 の線
    for (const r of rs) expect(r.y + r.h <= base + 0.02 || r.y >= base - 0.02).toBe(true);
  });

  it('TD-106-22: 間引いた場面の経路 d は数値と M L C A Z だけ', () => {
    for (const m of [
      denseLine(),
      denseLine({}, { smooth: true, marker: { symbol: 'circle' } }),
      denseLine({}, { marker: { symbol: 'diamond' } }),
      denseBar(),
      denseBar({ barDir: 'bar', grouping: 'stacked' }),
    ])
      for (const p of paths(scene(m, W, H))) expect(onlyKnownCommands(p.d)).toBe(true);
  });

  // ---- ISSUE-111：円・ドーナツの扇の束ね・集合の棒の列の幅 ----
  /** 大きな扇 1 つ（値 1000）の後に細い扇が n 個（値 1 ずつ） */
  const bigThenTiny = (n: number, serOver: Partial<ChartSeries> = {}): ChartModel =>
    pie({
      categories: cats(n + 1),
      series: [ser(0, 's', [1000, ...Array.from({ length: n }, () => 1)], serOver)],
    });
  const sectorsOf = (s: ChartScene): ChartPathItem[] =>
    fillPaths(s).filter((p) => p.d.startsWith('M') && p.d.includes('A'));
  const sameScene = (m: ChartModel): void =>
    expect(layoutChart(m, W, H)).toEqual(layoutChart(m, W, H, undefined, { thin: false }));

  it('TD-111-3: 細い扇の連なりは線の色で塗った扇 1 つ（同じ線付き）に束ねる・大きな扇はそのまま', () => {
    const m = bigThenTiny(500);
    const before = sectorsOf(off(m));
    const after = sectorsOf(scene(m, W, H));
    expect(before).toHaveLength(501);
    expect(after).toHaveLength(2);
    expect(after[0]).toEqual(before[0]);
    const merged = after[1]!;
    const line = before[1]!.stroke!;
    expect(merged.stroke).toEqual(line);
    expect(merged.fill).toEqual({ color: line.color, opacity: line.opacity });
    // 区間の開始は最初の細い扇の開始・終わりは最後の細い扇の終わり
    const first = before[1]!.d.match(/^M[\d.-]+ [\d.-]+L([\d.-]+ [\d.-]+)A/)![1];
    const last = before[500]!.d.match(/A[\d. ]+ [01] 1 ([\d.-]+ [\d.-]+)Z$/)![1];
    expect(merged.d).toContain(`L${first}A`);
    expect(merged.d.endsWith(` ${last}Z`)).toBe(true);
  });

  it('TD-111-4: 線が半透明・破線・線なし・太さ 0 の円は束ねない・c:dPt の扇は束ねず区間を切る', () => {
    const solid = { kind: 'solid', color: { rgb: 0xffffff, alpha: 1 } } as const;
    const lineOf = (over: object): Partial<ChartSeries> => ({
      line: { paint: solid, widthEmu: 19050, ...over },
    });
    sameScene(
      bigThenTiny(60, lineOf({ paint: { kind: 'solid', color: { rgb: 0xffffff, alpha: 0.5 } } })),
    );
    sameScene(bigThenTiny(60, lineOf({ dash: 'dash' })));
    sameScene(bigThenTiny(60, { line: { paint: { kind: 'none' } } }));
    sameScene(bigThenTiny(60, lineOf({ widthEmu: 0 })));
    // c:dPt で上書きした扇（添字 30）は残り、前後の区間は別々に束ねる
    const dpt = bigThenTiny(60, {
      points: new Map([[30, { kind: 'solid', color: { rgb: 0xff0000, alpha: 1 } }]]),
    });
    const sc = sectorsOf(scene(dpt, W, H));
    expect(sc).toHaveLength(4); // 大・束（1〜29）・上書き（30）・束（31〜60）
    expect(sc[2]!.fill).toEqual({ color: '#ff0000', opacity: 1 });
  });

  it('TD-111-5: ドーナツの輪の扇も束ねる（輪の形の経路のまま）', () => {
    const m = pie({
      type: 'doughnut',
      holeSize: 50,
      categories: cats(301),
      series: [ser(0, 's', [1000, ...Array.from({ length: 300 }, () => 1)])],
    });
    const before = sectorsOf(off(m));
    const after = sectorsOf(scene(m, W, H));
    expect(before).toHaveLength(301);
    expect(after).toHaveLength(2);
    expect(after[1]!.d.match(/A/g)).toHaveLength(2); // 外の弧と内の弧
  });

  it('TD-111-6: 扇が 1,000 個の円は、細い扇が続く所を束ねて経路が大きく減る', () => {
    const m = pie({ categories: cats(1000), series: [ser(0, 's', wave(0, 1000))] });
    expect(sectorsOf(off(m))).toHaveLength(1000);
    expect(sectorsOf(scene(m, W, H)).length).toBeLessThan(40);
  });

  it('TD-111-7: 集合の棒（10 系列）の列は 2px・積み上げ・重なり 100 は 1px のまま', () => {
    const count = (s: ChartScene): number =>
      fillPaths(s).reduce((n, p) => n + (p.d.match(/Z/g) ?? []).length, 0);
    // 系列 10×列（600px/2px＝300）以下
    expect(count(scene(denseBar(), W, H))).toBeLessThanOrEqual(10 * 300 + 10);
    // 積み上げ・重なり 100 は 1px の列のまま（列は 2px の列より多い）
    expect(count(scene(denseBar({ grouping: 'stacked' }), W, H))).toBeGreaterThan(10 * 300 + 10);
    expect(count(scene(denseBar({ overlap: 100 }), W, H))).toBeGreaterThan(10 * 300 + 10);
  });
});

describe('ISSUE-108 照合の修正（右上の凡例でプロットを縮める・横並びの等幅の格子・縦の項目数の上限）', () => {
  const lg = (pos: 'r' | 'l' | 't' | 'b' | 'tr'): ChartModel['legend'] => ({
    pos,
    overlay: false,
    deleted: [],
    text: {},
  });
  const sz = CHART_DEFAULTS.legendPt * PX_PER_PT;
  /** 凡例の文字（anchor start・baseline middle）。見た目の順（行 → 左から） */
  const legendItems = (s: ChartScene): ChartTextItem[] =>
    texts(s)
      .filter((t) => t.anchor === 'start' && t.baseline === 'middle')
      .sort((a, b) => a.y - b.y || a.x - b.x);
  /** 一番上の数値軸の文字の y（プロットの上端） */
  const plotTop = (s: ChartScene): number =>
    Math.min(
      ...texts(s)
        .filter((t) => t.anchor === 'end' && t.baseline === 'middle')
        .map((t) => t.y),
    );
  const col = (legend: ChartModel['legend']): ChartModel =>
    mk({
      categories: ['A', 'B', 'C', 'D', 'E'],
      series: [
        ser(0, 'S1', [10, 20, 30, 40, 50]),
        ser(1, 'S2', [5, 6, 7, 8, 9]),
        ser(2, 'S3', [1, 2, 3, 4, 5]),
      ],
      legend,
    });
  const labels13 = Array.from({ length: 13 }, (_, i) => 'P' + (i + 1));
  const pie13 = (pos: 'r' | 't' | 'b'): ChartModel =>
    pie({
      categories: labels13,
      series: [
        ser(
          0,
          's',
          labels13.map((_, i) => i + 1),
        ),
      ],
      legend: lg(pos),
    });
  const distinct = (xs: number[]): number[] => [
    ...new Set(xs.map((x) => Math.round(x * 100) / 100)),
  ];

  it('TD-108-1: 右上の凡例は、プロットの上端を凡例の下端 + 5pt まで縮める（標準の面・C5）', () => {
    const s = scene(col(lg('tr')), 480, 288);
    const items = legendItems(s);
    expect(items).toHaveLength(3);
    const rowH = sz * (1.2 + 0.675);
    const legendBottom = items[2]!.y + rowH / 2;
    expect(plotTop(s)).toBeCloseTo(legendBottom + 5 * PX_PER_PT, 1);
    // 右の凡例（縮めない）より上端がずっと低い
    expect(plotTop(s)).toBeGreaterThan(plotTop(scene(col(lg('r')), 480, 288)) + rowH);
  });

  it('TD-108-2: 縮めた後のプロットが 6 割に満たない小さい面では縮めず、凡例とプロットを 6pt だけ下げる（C5）', () => {
    const s = scene(col(lg('tr')), 288, 192);
    const items = legendItems(s);
    const r = scene(col(lg('r')), 288, 192);
    expect(plotTop(s)).toBeCloseTo(plotTop(r) + 6 * PX_PER_PT, 1);
    expect(plotTop(s)).toBeLessThan(items[2]!.y);
    // 凡例も 6pt 下がる（右の凡例は縦に中央。右上は面の上の縁から）
    expect(items[0]!.y).toBeCloseTo(2 * 6 * PX_PER_PT + (sz * (1.2 + 0.675)) / 2, 1);
  });

  it('TD-108-3: 13 項目の横並びで 1 行に収まらない時は 2 行・7 桁の等幅の格子（桁の幅は最も長い項目）・収まる幅では 1 行（C6）', () => {
    const narrow = legendItems(scene(pie13('t'), 266, 288)); // 幅 200pt
    expect(narrow).toHaveLength(13);
    expect(distinct(narrow.map((t) => t.y))).toHaveLength(2);
    const first = narrow.slice(0, 7);
    const second = narrow.slice(7);
    expect(second).toHaveLength(6);
    // 2 行目の x は 1 行目の最初の 6 桁と同じ
    expect(second.map((t) => t.x)).toEqual(first.slice(0, 6).map((t) => t.x));
    // 桁の幅が一定
    const pitch = first.slice(1).map((t, i) => t.x - first[i]!.x);
    expect(Math.max(...pitch) - Math.min(...pitch)).toBeLessThan(1e-6);
    // 下の凡例でも同じ並び
    const bottom = legendItems(scene(pie13('b'), 266, 288));
    expect(distinct(bottom.map((t) => t.y))).toHaveLength(2);
    expect(distinct(bottom.map((t) => t.x))).toHaveLength(7);
    // 広い面（360pt）は 1 行
    const wide = legendItems(scene(pie13('t'), 480, 288));
    expect(distinct(wide.map((t) => t.y))).toHaveLength(1);
  });

  it('TD-108-4: 格子の桁数は「行数 = ceil(項目数 / 入る桁数)・桁数 = ceil(項目数 / 行数)」で均す（280pt は 11 桁入るが 2 行 7 桁・C6）', () => {
    const mid = legendItems(scene(pie13('t'), 373, 288));
    expect(distinct(mid.map((t) => t.y))).toHaveLength(2);
    expect(distinct(mid.map((t) => t.x))).toHaveLength(7);
    // 凡例は面の中央（記号の左端から最後の桁の文字の右端までの中央が面の中央から 1 桁以内）
    const xs = mid.map((t) => t.x);
    const mid2 = (Math.min(...xs) + Math.max(...xs)) / 2;
    expect(Math.abs(mid2 - 373 / 2)).toBeLessThan(sz * 2);
  });

  it('TD-108-5: 縦並びの凡例は、面の上下の縁との間（6pt ずつ）に収まる項目だけ描く（高さ 216pt・13 項目で 10 項目・C3）', () => {
    expect(legendItems(scene(pie13('r'), 480, 288))).toHaveLength(10);
    expect(legendItems(scene(pie13('r'), 480, 288 + 3 * 22.5))).toHaveLength(13);
    expect(legendItems(scene(pie({ ...pie13('r'), legend: lg('r') }), 480, 150))).toHaveLength(5);
  });
});

describe('ISSUE-110 照合の修正（横棒の刻みは、自動配置の外側の幅で決める）', () => {
  /** 横棒の積み上げ（正の和の最大 20・負の和の最小 -25＝Excel の c1sw_barStacked_neg_*。軸の文字は 9pt）。項目名は長さを変えられる */
  const stacked = (catLen: number): ChartModel =>
    mk({
      barDir: 'bar',
      grouping: 'stacked',
      catAxis: { ...catAx, text: { sizePt: 9 } },
      valAxis: { ...valAx, text: { sizePt: 9 } },
      categories: ['C1', 'C2', 'C3', 'C4', 'C5'].map((c) => c.padEnd(catLen, 'x')),
      series: [
        ser(0, 'S1', [12, -5, 7, -15, 9]),
        ser(1, 'S2', [-8, 14, -11, 6, -3]),
        ser(2, 'S3', [5, -9, 13, -10, 4]),
      ],
    });
  /** 数値軸の目盛の文字（数字だけの文字）。左から */
  const ticks = (s: ChartScene): number[] =>
    texts(s)
      .filter((t) => /^-?\d+$/.test(t.text))
      .map((t) => Number(t.text));

  it('TD-110-1: 面の幅 360pt では 10 刻み・380pt では 5 刻み（Excel の c1sw_barStacked_neg_360／380）', () => {
    expect(ticks(scene(stacked(2), 360 * PX_PER_PT, 288))).toEqual([-30, -20, -10, 0, 10, 20, 30]);
    expect(ticks(scene(stacked(2), 380 * PX_PER_PT, 288))).toEqual([
      -30, -25, -20, -15, -10, -5, 0, 5, 10, 15, 20, 25,
    ]);
  });

  it('TD-110-2: 項目名が長くてプロットが細くなっても、刻みは外側の幅で決まる（項目軸の文字の帯で刻みを粗くしない）', () => {
    // 項目名 2 文字と 20 文字：プロットの内側の幅は違うが、面の幅が同じなら刻みは同じ
    const short = ticks(scene(stacked(2), 380 * PX_PER_PT, 288));
    const long = ticks(scene(stacked(20), 380 * PX_PER_PT, 288));
    expect(long).toEqual(short);
  });
});

// ---- 項目軸の文字の形（ISSUE-107・ADR-0031 §決定 5） ----

describe('項目軸の文字の形（水平・折り返し・45°・90°・間引き）', () => {
  /** 文字 9pt（12px）・1 文字 6.6px の推定幅（ASCII）。面 480×288px（360×216pt） */
  const PX = 12;
  const CH = 6.6;
  const catAx9: ChartAxis = { ...catAx, text: { sizePt: 9 } };
  const colModel = (cats: string[], over: Partial<ChartModel> = {}): ChartModel =>
    mk({
      categories: cats,
      series: [
        ser(
          0,
          's',
          cats.map((_, i) => i + 1),
        ),
      ],
      catAxis: catAx9,
      valAxis: { ...valAx, text: { sizePt: 9 } },
      plotArea: { fill: solid(0x123456) },
      ...over,
    });
  /** `C01xxxx…`（長さ len）の項目名を n 個 */
  const names = (n: number, len: number): string[] =>
    Array.from({ length: n }, (_, i) =>
      `C${String(i + 1).padStart(2, '0')}${'x'.repeat(len)}`.slice(0, len),
    );
  const catTexts = (s: ChartScene, cats: string[]): ChartTextItem[] =>
    texts(s).filter((t) => cats.includes(t.text));

  it('TD-107-1: 水平は、どの項目名の幅も項目の間隔以下（境目は ≤）', () => {
    const w = estimateTextWidth('abcd', PX);
    expect(planCategoryLabelForm(['abcd', 'abc'], PX, w).form).toBe('horizontal');
    expect(planCategoryLabelForm(['abcd', 'abc'], PX, w - 0.01).form).toBe('rotate45');
    expect(planCategoryLabelForm(['abcd', 'abc'], PX, 10).form).toBe('rotate90');
    expect(planCategoryLabelForm([], PX, 10).form).toBe('horizontal');
  });

  it('TD-107-2: wrapCategoryLabel は語を左から幅いっぱいまで詰める・語が入らなければ null・1 行に入ればそのまま', () => {
    const t = 'C01 alpha beta';
    expect(wrapCategoryLabel(t, PX, 14 * CH)).toEqual([t]);
    expect(wrapCategoryLabel(t, PX, 10 * CH)).toEqual(['C01 alpha', 'beta']);
    expect(wrapCategoryLabel(t, PX, 5 * CH)).toEqual(['C01', 'alpha', 'beta']);
    expect(wrapCategoryLabel(t, PX, 4 * CH)).toBeNull(); // alpha（5 文字）が入らない
    expect(wrapCategoryLabel('abcdefghij', PX, 5 * CH)).toBeNull(); // 空白が無い
  });

  it('TD-107-3: 折り返しは、幅が超える名前が全部 空白を含み 語が入る時だけ。入らない名前が 1 つでもあれば回す', () => {
    const slot = 10 * CH;
    expect(planCategoryLabelForm(['C01 alpha beta', 'C02 alpha beta'], PX, slot)).toEqual({
      form: 'wrap',
      rows: 2,
    });
    expect(planCategoryLabelForm(['C01 alpha beta', 'C02 abcdefghijklm'], PX, slot).form).not.toBe(
      'wrap',
    );
    expect(planCategoryLabelForm(['C01 alpha beta', 'short'], PX, slot).form).toBe('wrap'); // 入る名前はそのまま
    expect(planCategoryLabelForm(['alphabetagamma delta'], PX, 15).form).toBe('rotate90'); // 語が入らない
  });

  it('TD-107-4: 45° と 90° の境目は、項目の間隔が文字の大きさの 1.6 倍（9pt で 14.4pt）以上かどうか', () => {
    const lbl = ['abcdefghijklmnopqrst'];
    const t = 1.6 * PX;
    expect(planCategoryLabelForm(lbl, PX, t).form).toBe('rotate45');
    expect(planCategoryLabelForm(lbl, PX, t - 0.01).form).toBe('rotate90');
  });

  it('TD-107-5: 水平：全部を回さず出す。帯は 1 行（プロットの下に 余白 5pt ＋ 18pt）', () => {
    const cats = names(5, 8);
    const s = scene(colModel(cats), 480, 288);
    const ts = catTexts(s, cats);
    expect(ts).toHaveLength(5);
    expect(ts.every((t) => t.rotate === 0 && t.anchor === 'middle')).toBe(true);
    const p = plotRect(s);
    expect(288 - (p.y + p.height)).toBeCloseTo(23 * PX_PER_PT, 0);
  });

  it('TD-107-6: 空白を含む長い名前は折り返す：行ごとの文字が一定の送りで下へ並び、帯が厚くなる', () => {
    const cats = Array.from(
      { length: 10 },
      (_, i) => `C${String(i + 1).padStart(2, '0')} alpha be`,
    );
    const s = scene(colModel(cats), 480, 288);
    const x0 = texts(s).find((t) => t.text === 'C01')!.x;
    const first = texts(s).filter((t) => t.x === x0 && ['C01', 'alpha', 'be'].includes(t.text));
    expect(first.map((t) => t.text)).toEqual(['C01', 'alpha', 'be']);
    expect(first.every((t) => t.rotate === 0 && t.anchor === 'middle')).toBe(true);
    const pitch = first[1]!.y - first[0]!.y;
    expect(pitch).toBeCloseTo(PX * (11 / 9), 5);
    expect(first[2]!.y - first[1]!.y).toBeCloseTo(pitch, 5);
    // 水平 1 行（短い名前）の時よりプロットが 2 行分低い
    const flat = plotRect(scene(colModel(names(10, 3)), 480, 288));
    expect(flat.height - plotRect(s).height).toBeCloseTo(2 * pitch, 0);
  });

  it('TD-107-7: 45°：右上がり（−45°・末尾を目盛に揃える）で全項目を出す。帯は (幅＋行の高さ)·sin45° ＋ 6.5pt', () => {
    const cats = names(10, 20);
    const s = scene(colModel(cats), 480, 288);
    const ts = catTexts(s, cats);
    expect(ts).toHaveLength(10);
    expect(ts.every((t) => t.rotate === -45 && t.anchor === 'end')).toBe(true);
    const p = plotRect(s);
    const band = (20 * CH + PX * 1.2) * Math.SQRT1_2 + 6.5 * PX_PER_PT;
    expect(288 - (p.y + p.height)).toBeCloseTo(5 * PX_PER_PT + band, 0);
  });

  it('TD-107-8: 45°：最初の項目の文字が左へ出る分だけプロットの左が内側へ（(幅＋行の高さ)·cos45° − 最初の項目の中心 ＋ 2pt）', () => {
    const cats = names(10, 20);
    const p = plotRect(scene(colModel(cats), 480, 288));
    const reach = (20 * CH + PX * 1.2) * Math.SQRT1_2 - p.width / 10 / 2 + 2 * PX_PER_PT;
    expect(p.x).toBeCloseTo(reach, 0);
    // 水平の時（短い名前）の左（数値軸の文字の分だけ）より大きい
    expect(p.x).toBeGreaterThan(plotRect(scene(colModel(names(10, 3)), 480, 288)).x + 10);
  });

  it('TD-107-9: 90°：間隔が狭い時は −90° に回し、行の高さ＋余白 1pt が入る個数に間引く（30 項目で 2 個おき）', () => {
    const cats = names(30, 12);
    const ts = catTexts(scene(colModel(cats), 480, 288), cats);
    expect(ts.every((t) => t.rotate === -90 && t.anchor === 'end')).toBe(true);
    expect(ts).toHaveLength(15);
    expect(ts[0]!.text).toBe(cats[0]);
    expect(ts[1]!.text).toBe(cats[2]);
  });

  it('TD-107-10: 明示の a:bodyPr rot は従来どおり（自動の形を選ばない。0° の明示も回さず間引く）', () => {
    const cats = names(10, 20);
    const rot = (deg: number): ChartModel =>
      colModel(cats, { catAxis: { ...catAx9, text: { sizePt: 9, rotDeg: deg } } });
    const s = scene(rot(-30), 480, 288);
    expect(catTexts(s, cats).every((t) => t.rotate === -30)).toBe(true);
    const auto = plotRect(scene(colModel(cats), 480, 288));
    expect(plotRect(s).height).toBeGreaterThan(auto.height); // 30° の帯は 45° より薄い
    expect(plotRect(s).x).toBeLessThan(auto.x); // 左へ張り出さない（従来）
    const z = catTexts(scene(rot(0), 480, 288), cats);
    expect(z.every((t) => t.rotate === 0)).toBe(true);
    expect(z.length).toBeLessThan(10);
  });

  it('TD-107-11: 明示の c:tickLblSkip は最優先（形は 間隔 × skip で選ぶ。90° でも skip を守る）', () => {
    const cats = names(10, 20);
    const ts = catTexts(
      scene(colModel(cats, { catAxis: { ...catAx9, tickLblSkip: 3 } }), 480, 288),
      cats,
    );
    expect(ts.map((t) => t.text)).toEqual([cats[0], cats[3], cats[6], cats[9]]);
    expect(ts.every((t) => t.rotate === -45)).toBe(true); // 間隔 × 3（約 120px）は 20 文字の幅（132px）に足りない → 45°
    const c12 = names(10, 12);
    const h = catTexts(
      scene(colModel(c12, { catAxis: { ...catAx9, tickLblSkip: 3 } }), 480, 288),
      c12,
    );
    expect(h).toHaveLength(4);
    expect(h.every((t) => t.rotate === 0)).toBe(true); // 12 文字（79px）は間隔 × 3 に入る → 水平
    const c30 = names(30, 12);
    const s90 = scene(colModel(c30, { catAxis: { ...catAx9, tickLblSkip: 1 } }), 480, 288);
    expect(catTexts(s90, c30)).toHaveLength(30);
  });

  it('TD-107-12: 帯が面の高さの 0.4 を超える小さい面は回さず、従来どおり間引く（安全弁）', () => {
    const cats = names(10, 20);
    const ts = catTexts(scene(colModel(cats), 480, 120), cats);
    expect(ts.every((t) => t.rotate === 0)).toBe(true);
    expect(ts.length).toBeLessThan(10);
  });

  it('TD-107-13: 横棒の項目軸（縦の軸）は回さず折り返さない（従来どおり水平で間引く）', () => {
    const cats = Array.from(
      { length: 15 },
      (_, i) => `C${String(i + 1).padStart(2, '0')} alpha beta gamma`,
    );
    const s = scene(colModel(cats, { barDir: 'bar' }), 480, 288);
    const ts = catTexts(s, cats);
    expect(ts.length).toBeGreaterThan(0);
    expect(ts.every((t) => t.rotate === 0 && t.anchor === 'end')).toBe(true);
  });

  it('TD-107-14: 折れ線も同じ規則', () => {
    const cats = names(10, 20);
    const s = scene(colModel(cats, { type: 'line', barDir: undefined }), 480, 288);
    const ts = catTexts(s, cats);
    expect(ts).toHaveLength(10);
    expect(ts.every((t) => t.rotate === -45)).toBe(true);
  });

  it('TD-107-15: 項目が 0 個でも落ちない・項目軸を消す（deleted）と文字を出さず回さない', () => {
    expect(() => scene(colModel([]), 480, 288)).not.toThrow();
    const cats = names(10, 20);
    const s = scene(colModel(cats, { catAxis: { ...catAx9, deleted: true } }), 480, 288);
    expect(catTexts(s, cats)).toHaveLength(0);
  });
});

describe('ISSUE-109 既定のスタイル（c14:style・c:style 41〜48・c:roundedCorners）', () => {
  /** 目盛線・軸線・目盛の印（塗りのない経路と線）の色の集合 */
  const lineColors = (s: ChartScene): Set<string> =>
    new Set([
      ...paths(s)
        .filter((p) => p.fill === null && p.stroke !== null)
        .map((p) => p.stroke!.color),
      ...s.items.flatMap((i) => (i.kind === 'line' ? [i.stroke.color] : [])),
    ]);
  const title = (s: ChartScene): ChartTextItem => texts(s).find((t) => t.text === 'T')!;
  const withTitle = { title: { lines: ['T'], text: {}, overlay: false } };
  const hex6 = (rgb: number): string => `#${rgb.toString(16).padStart(6, '0')}`;

  it('TD-109-2: c14Style を持つ（c:spPr 無し）と、主の目盛線・軸線・目盛の印が #898989 0.5pt・タイトルが 18pt 太字', () => {
    const s = scene(mk({ ...withTitle, c14Style: 102 }));
    expect([...lineColors(s)]).toEqual(['#898989']);
    const ln = s.items.find((i) => i.kind === 'line')!;
    expect(ln.kind === 'line' && ln.stroke.width).toBeCloseTo(0.5 * PX_PER_PT, 6);
    expect(title(s)).toMatchObject({
      size: 18 * PX_PER_PT,
      bold: true,
      color: { color: '#000000' },
    });
    // 持たなければ今までどおり（黒・タイトル 10pt）
    const plain = scene(mk(withTitle));
    expect([...lineColors(plain)]).toEqual(['#000000']);
    expect(title(plain).size).toBe(10 * PX_PER_PT);
    // c14Style の値は見ない（持つかだけ）
    expect([...lineColors(scene(mk({ c14Style: 1 })))]).toEqual(['#898989']);
  });

  it('TD-109-2: c14Style でも補助の目盛線・軸と凡例の文字・面・折れ線の太さは変えない。明示（c:spPr・c:txPr）が優先', () => {
    const s = scene(mk({ c14Style: 102, valAxis: { ...valAx, minorGridlines: {} } }));
    expect(lineColors(s)).toEqual(new Set(['#898989', '#bcbcbc']));
    const ca = rects(s)[0]!;
    expect(ca.fill?.color).toBe('#ffffff');
    expect(ca.stroke).toMatchObject({ color: '#898989' });
    expect(ca.stroke!.width).toBeCloseTo(0.5 * PX_PER_PT, 6);
    expect(texts(s)[0]!).toMatchObject({ size: 10 * PX_PER_PT, color: { color: '#000000' } });
    const ln = scene(mk({ c14Style: 102, type: 'line', barDir: undefined, grouping: 'standard' }));
    const lp = paths(ln).find((x) => x.fill === null && x.stroke?.color === '#4472c4')!;
    expect(lp.stroke!.width).toBeCloseTo(1.5 * PX_PER_PT, 6);
    // 明示が優先：目盛線の色・タイトルの大きさ
    const ex = scene(
      mk({
        c14Style: 102,
        title: { lines: ['T'], text: { sizePt: 12 }, overlay: false },
        valAxis: { ...valAx, majorGridlines: { paint: solid(0xff00ff) } },
      }),
    );
    expect(lineColors(ex).has('#ff00ff')).toBe(true);
    expect(title(ex).size).toBe(12 * PX_PER_PT);
  });

  it('TD-109-3: c:style 41〜48 は面 黒（枠線 黒 0.5pt）・文字 白 10pt・プロット #3F3F3F。40・49 や style 無しは今までどおり', () => {
    for (const style of [41, 42, 45, 48]) {
      const s = scene(mk({ ...withTitle, style }));
      const ca = rects(s)[0]!;
      expect(ca.fill?.color).toBe('#000000');
      expect(ca.stroke).toMatchObject({ color: '#000000' });
      expect(ca.stroke!.width).toBeCloseTo(0.5 * PX_PER_PT, 6);
      expect(rects(s).some((r) => r.fill?.color === '#3f3f3f')).toBe(true);
      const ts = texts(s);
      expect(ts.length).toBeGreaterThan(1);
      expect(ts.every((t) => t.color.color === '#ffffff')).toBe(true);
      expect(title(s)).toMatchObject({ size: 10 * PX_PER_PT, bold: true });
      // c14 無しの暗い背景は目盛線・軸線が黒（c14:style とは独立）
      expect([...lineColors(s)]).toEqual(['#000000']);
    }
    for (const style of [undefined, 2, 40, 49]) {
      const s = scene(mk({ style }));
      expect(rects(s)[0]!.fill?.color).toBe('#ffffff');
      expect(rects(s).some((r) => r.fill?.color === '#3f3f3f')).toBe(false);
      expect(texts(s).every((t) => t.color.color === '#000000')).toBe(true);
    }
  });

  it('TD-109-3: 暗い背景×c14Style は目盛線・軸線 #898989・タイトル 18pt 白（2 つは独立に効く）', () => {
    const s = scene(mk({ ...withTitle, style: 42, c14Style: 142 }));
    expect([...lineColors(s)]).toEqual(['#898989']);
    expect(title(s)).toMatchObject({
      size: 18 * PX_PER_PT,
      bold: true,
      color: { color: '#ffffff' },
    });
    expect(rects(s)[0]!.fill?.color).toBe('#000000');
  });

  it('TD-109-3: 円・ドーナツのプロットは塗りなし（黒い面が見える）。明示の面・プロット・文字の色が優先', () => {
    const p = scene(pie({ style: 42 }));
    expect(rects(p)).toHaveLength(1);
    expect(rects(p)[0]!.fill?.color).toBe('#000000');
    const ex = scene(
      mk({
        style: 42,
        chartArea: { fill: solid(0x112233) },
        plotArea: { fill: solid(0x445566) },
        text: { color: { rgb: 0xff0000, alpha: 1 } },
      }),
    );
    expect(rects(ex)[0]!.fill?.color).toBe('#112233');
    expect(rects(ex).some((r) => r.fill?.color === '#445566')).toBe(true);
    expect(rects(ex).some((r) => r.fill?.color === '#3f3f3f')).toBe(false);
    expect(texts(ex).every((t) => t.color.color === '#ff0000')).toBe(true);
    // 面が塗りなしなら backdrop（暗い背景でも塗りが無ければ今までどおり）
    expect(scene(mk({ style: 42, chartArea: { fill: { kind: 'none' } } })).backdrop).toBe(true);
    expect(scene(mk({ style: 42 })).backdrop).toBe(false);
  });

  it('TD-109-3: 暗い背景の棒の既定の塗りは系列の色の彩度 110%（グラデーションの 50% の位置に近い単色）。折れ線・円・明示の塗りは変えない', () => {
    const want = applyTransforms({ rgb: 0x4472c4, alpha: 1 }, '<a:satMod val="110000"/>');
    expect(hex6(want.rgb)).not.toBe('#4472c4');
    const s = scene(mk({ style: 42 }));
    expect(fillPaths(s).map((p) => p.fill!.color)).toEqual([hex6(want.rgb)]);
    // style 40 は変えない・暗くても折れ線の線は系列の色のまま・円は点の位置の色のまま
    expect(fillPaths(scene(mk({ style: 40 }))).map((p) => p.fill!.color)).toEqual(['#4472c4']);
    const ln = scene(mk({ style: 42, type: 'line', barDir: undefined, grouping: 'standard' }));
    expect(paths(ln).some((x) => x.fill === null && x.stroke?.color === '#4472c4')).toBe(true);
    const pc = fillPaths(scene(pie({ style: 42 }))).map((p) => p.fill!.color);
    expect(pc[0]).toBe('#4472c4');
    // 系列の明示の塗りが優先
    const ex = scene(
      mk({ style: 42, series: [ser(0, 'S1', [1, 2, 3], { fill: solid(0x00ff00) })] }),
    );
    expect(fillPaths(ex).map((p) => p.fill!.color)).toEqual(['#00ff00']);
  });

  it('TD-109-4: 面の角丸は半径 11px（8.25pt）一定。要素なし・真は角丸、偽は直角。プロットの面は角丸にしない', () => {
    expect(CHART_DEFAULTS.cornerRadiusPt * PX_PER_PT).toBeCloseTo(11, 9);
    for (const roundedCorners of [undefined, true]) {
      for (const [w, h] of [
        [288, 192],
        [480, 288],
        [960, 576],
      ] as const) {
        const s = scene(mk({ roundedCorners, plotArea: { fill: solid(0x123456) } }), w, h);
        expect(rects(s)[0]!.radius).toBeCloseTo(11, 9);
        expect(s.cornerRadius).toBeCloseTo(11, 9);
        const plot = rects(s).find((r) => r.fill?.color === '#123456')!;
        expect(plot.radius).toBeUndefined();
      }
    }
    const off = scene(mk({ roundedCorners: false }));
    expect(rects(off)[0]!.radius).toBeUndefined();
    expect(off.cornerRadius).toBe(0);
  });

  it('TD-109-4: 半径は 幅÷2・高さ÷2 を超えない。塗りなしの面は矩形を出さず cornerRadius だけ持つ（下地の半径用）', () => {
    expect(scene(mk(), 16, 100).cornerRadius).toBeCloseTo(8, 9);
    expect(rects(scene(mk(), 100, 10))[0]!.radius).toBeCloseTo(5, 9);
    const nf = scene(
      mk({ chartArea: { fill: { kind: 'none' }, line: { paint: { kind: 'none' } } } }),
    );
    expect(nf.backdrop).toBe(true);
    expect(rects(nf).some((r) => r.width === nf.width && r.height === nf.height)).toBe(false);
    expect(nf.cornerRadius).toBeCloseTo(11, 9);
    // 暗い背景の角丸の矩形（面の塗り・枠線ともに同じ半径）
    expect(rects(scene(mk({ style: 42, roundedCorners: true })))[0]!.radius).toBeCloseTo(11, 9);
  });
});
