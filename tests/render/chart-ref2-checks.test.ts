/// <reference types="node" />
/**
 * 2 回目の照合（ISSUE-105）の検査の部品（`xlsx-chart-ref2-checks.ts`・`xlsx-chart-ref-scene.ts`）の自己整合の試験。
 * Excel の基準値（COM の値）が無くても動くよう、場面から COM の値と同じ形の値を作り（`fakeCom`）、次を固定する：
 * 作った値と場面を比べると差が出ない（px ↔ pt・原点 4pt・凡例の記号の色 → 名前の対応が噛み合っている）・
 * 値を 1 つ変えると、その検査（軸・プロット・凡例の数と名前・補助の刻み・目盛の印・マーカー・PNG の ink）が差を出す。
 * 本物の基準値との照合は `xlsx-chart-ref2.test.ts`（ファイルが来るまで skip）。
 */
import { describe, expect, it } from 'vitest';
import type { PngChartResult } from '../../fixtures/excel/make-ref-charts-png.mjs';
import type { ChartModel } from '../../src/core/types';
import { barModel, lineModel, series } from './chart-fixtures';
import { layoutChart, type ChartScene, type ChartStroke } from '../../src/render/chart-layout';
import {
  axisCheck,
  GROUPS,
  legendCheck,
  markerCheck,
  minorCheck,
  plotCheck,
  pngCheck,
  runGroup,
  styleCheck,
  tickCheck,
  type Case,
  type CheckFn,
  type Group,
} from './chart-ref2-checks';
import {
  axisOf,
  axisStrokes,
  COM_ORIGIN_PT,
  expectedLegend,
  gridBox,
  judge,
  K,
  legendBox,
  legendItems,
  lim,
  listed,
  minorUnitOf,
  mixed,
  plotOf,
  same,
  segments,
  sortVisual,
  type ComChart,
  type Diff,
  type TextItem,
} from './chart-ref-scene';

const W_PT = 360;
const H_PT = 216;

const toPt = (px: number, origin = true): number => px / K - (origin ? COM_ORIGIN_PT : 0);
const rrggbb = (c: string): string => c.slice(1).toUpperCase();

/** 場面から、COM が返すはずの値（pt・原点 4pt）を作る。軸・プロット・凡例・面・文字は場面と一致する */
function fakeCom(model: ChartModel, scene: ChartScene): ComChart {
  const plot = plotOf(scene, model)!;
  const items = legendItems(scene);
  const lb = legendBox(scene);
  const ca = scene.items[0];
  const stroke = ca?.kind === 'rect' ? ca.stroke : null;
  const title = scene.items.find(
    (i): i is TextItem => i.kind === 'text' && i.baseline === 'middle' && i.anchor === 'middle',
  );
  const ax = model.type === 'pie' ? null : axisOf(scene, model);
  const lab = scene.items.find(
    (i): i is TextItem => i.kind === 'text' && i.anchor === 'end' && i.baseline === 'middle',
  );
  const legendText = scene.items.find(
    (i): i is TextItem => i.kind === 'text' && i.anchor === 'start' && i.baseline === 'middle',
  );
  const n = items.length;
  const exp: ComChart = {
    ChartObject: { Width: W_PT, Height: H_PT },
    ChartArea: {
      Fill: { RGB: 'FFFFFF', Visible: -1 },
      Line: { Weight: (stroke?.width ?? 1) / K, RGB: rrggbb(stroke?.color ?? '#000000') },
    },
    PlotArea: {
      Inside: { Left: toPt(plot.x), Top: toPt(plot.y), Width: plot.w / K, Height: plot.h / K },
    },
    HasTitle: title !== undefined,
    HasLegend: n > 0,
    Group: { VaryByCategories: false },
    Series: model.series.map((s) => {
      // 凡例に出ている系列は記号の色、削除された系列は適当な別の色（凡例に出ないので何でもよい）
      const color = rrggbb(items.find((it) => it.text === s.name)?.keyColor ?? '#010203');
      return {
        Name: s.name,
        Fill: { RGB: color },
        Line: { RGB: color, Weight: 1 },
        MarkerStyle: -4142,
        MarkerSize: 5,
      };
    }),
  };
  if (title !== undefined) {
    exp.Title = {
      Text: title.text,
      Pos: { Left: toPt(title.x), Top: toPt(title.y), Width: 0, Height: 0 },
      Font: { Size: title.size / K, Bold: title.bold, RGB: rrggbb(title.color.color) },
    };
  }
  if (n > 0 && lb !== null) {
    // 凡例を n 行に等分した箱を項目とする（外接＝凡例の外接矩形・上から系列の順）
    exp.Legend = {
      Pos: { Left: toPt(lb.x), Top: toPt(lb.y), Width: lb.w / K, Height: lb.h / K },
      Font: {
        Size: (legendText?.size ?? 0) / K,
        RGB: rrggbb(legendText?.color.color ?? '#000000'),
      },
      Entries: items.map((it, i) => ({
        Pos: {
          Left: toPt(lb.x),
          Top: toPt(lb.y + (i * lb.h) / n),
          Width: lb.w / K,
          Height: lb.h / n / K,
        },
        KeyFillRGB: rrggbb(it.keyColor ?? '#000000'),
        KeyLineRGB: rrggbb(it.keyColor ?? '#000000'),
      })),
    };
  }
  if (ax !== null) {
    exp.ValueAxis = {
      MinimumScale: ax.min,
      MaximumScale: ax.max,
      MajorUnit: ax.major,
      MinorUnit: ax.major / 5,
      TickLabels: {
        Font: { Size: (lab?.size ?? 0) / K, RGB: rrggbb(lab?.color.color ?? '#000000') },
      },
    };
  }
  return exp;
}

/** 3 系列の縦棒（凡例 右・タイトル あり）。over で上書き */
function threeSeries(over: Partial<ChartModel> = {}): ChartModel {
  return barModel({
    categories: ['A', 'B', 'C', 'D'],
    series: [
      series(0, 'S1', [12, 19, 7, 15]),
      series(1, 'S2', [8, 14, 11, 6]),
      series(2, 'S3', [5, 9, 13, 10]),
    ],
    legend: { pos: 'r', overlay: false, deleted: [], text: {} },
    title: { lines: ['Title'], text: {}, overlay: false },
    ...over,
  });
}

function makeCase(model: ChartModel, name = 'x', group: Case['group'] = 'charts'): Case {
  const scene = layoutChart(model, W_PT * K, H_PT * K, undefined);
  if (scene === null) throw new Error('場面が作れない');
  return { name, group, exp: fakeCom(model, scene), model, scene };
}

/** 検査を 1 つのグラフに当てた差のメッセージ */
function msgs(check: CheckFn, c: Case): string[] {
  const out: string[] = [];
  check(c, (m) => out.push(m));
  return out;
}

describe('場面から作った COM 値との自己整合（差が出ない）', () => {
  it('縦棒・凡例・タイトル：軸・プロット・凡例・既定の見た目', () => {
    const c = makeCase(threeSeries());
    for (const check of [axisCheck, plotCheck, legendCheck, styleCheck]) {
      expect(msgs(check, c)).toEqual([]);
    }
    expect(c.exp.Legend!.Entries).toHaveLength(3);
  });

  it('凡例の項目を削除した縦棒（C11）：残った項目の名前を記号の色から引いて一致する', () => {
    const model = threeSeries({
      legend: { pos: 'r', overlay: false, deleted: [1], text: {} },
    });
    const c = makeCase(model);
    expect(legendItems(c.scene).map((i) => i.text)).toEqual(['S1', 'S3']);
    expect(expectedLegend(c.exp, c.model).names).toEqual(['S1', 'S3']);
    expect(msgs(legendCheck, c)).toEqual([]);
  });

  it('折れ線（記号は線の色）：凡例の名前を線の色から引く', () => {
    const model = lineModel(3, 6);
    const c = makeCase(model);
    expect(expectedLegend(c.exp, c.model).names).toEqual(['系列1', '系列2', '系列3']);
    expect(msgs(legendCheck, c)).toEqual([]);
  });

  it('GROUPS：接頭辞の無いグラフは対象外・対象になれば検査が走る', () => {
    const c = makeCase(threeSeries(), 'c11_col3_del2');
    const g = GROUPS.find((x) => x.cat === 'c11')!;
    expect(runGroup(g, [c])).toEqual([]);
    const other = makeCase(threeSeries(), 'other');
    other.exp.ValueAxis!.MaximumScale += 100;
    expect(runGroup(g, [other])).toEqual([]);
    expect(
      runGroup(
        GROUPS.find((x) => x.cat === 'c1')!,
        [{ ...other, name: 'c1neg_x' }],
      ),
    ).toHaveLength(1);
  });
});

describe('値を変えると差が出る', () => {
  it('軸の最大・刻み（C1・C2）', () => {
    const c = makeCase(threeSeries());
    c.exp.ValueAxis!.MaximumScale += 5;
    const m = msgs(axisCheck, c);
    expect(m).toHaveLength(1);
    expect(m[0]).toMatch(/^got\[0\.\.\d+ \/\d+\] want\[/);
  });

  it('プロットの位置（C5・C7）は面の 3% を超えると差・以内なら差なし', () => {
    const c = makeCase(threeSeries());
    c.exp.PlotArea.Inside.Height -= 0.03 * H_PT * 0.5;
    expect(msgs(plotCheck, c)).toEqual([]);
    c.exp.PlotArea.Inside.Height -= 0.03 * H_PT * 1.5;
    expect(msgs(plotCheck, c)[0]).toMatch(/^内側 got\[/);
  });

  it('凡例の項目の数（C3・C11）と名前（C6）', () => {
    const c = makeCase(threeSeries());
    c.exp.Legend!.Entries.pop();
    expect(msgs(legendCheck, c)[0]).toMatch(/^項目の数 got\[3: S1\|S2\|S3\] want\[2: S1\|S2\]/);
    const d = makeCase(threeSeries());
    d.exp.Series[1]!.Name = 'X2';
    expect(msgs(legendCheck, d)[0]).toMatch(
      /^項目の文字（見た目の順） got\[S1\|S2\|S3\] want\[S1\|X2\|S3\]/,
    );
    const e = makeCase(threeSeries());
    e.exp.Legend!.Entries[0]!.KeyFillRGB = '123456';
    expect(msgs(legendCheck, e)[0]).toMatch(/^COM の記号の色から名前を引けない項目がある/);
  });

  it('凡例が無いはずなのに出ている／凡例の外接矩形のずれ', () => {
    const c = makeCase(threeSeries());
    c.exp.HasLegend = false;
    expect(msgs(legendCheck, c)).toEqual(['凡例が無いはずなのに出ている']);
    const d = makeCase(threeSeries());
    for (const e of d.exp.Legend!.Entries) e.Pos.Left -= 40;
    expect(msgs(legendCheck, d)[0]).toMatch(/^項目 got\[/);
  });

  it('既定の見た目（C4）：c:style・面の塗りと枠線・タイトルと軸と凡例の文字', () => {
    const c = makeCase(threeSeries({ style: 3 }));
    c.exp.ChartStyle = 3;
    expect(msgs(styleCheck, c)).toEqual([]);
    c.exp.ChartStyle = 10;
    c.exp.ChartArea.Fill = { RGB: '000000', Visible: -1 };
    c.exp.ChartArea.Line.RGB = '123456';
    c.exp.Title!.Font.Size += 4;
    c.exp.ValueAxis!.TickLabels.Font.RGB = '654321';
    c.exp.Legend!.Font.Size += 2;
    const m = msgs(styleCheck, c);
    const heads = [
      'c:style got',
      '面の塗り got[',
      '面の枠線 got[',
      'タイトル got[',
      '軸の文字 got[',
      '凡例の文字 got[',
    ];
    expect(m).toHaveLength(heads.length);
    heads.forEach((h, i) => expect(m[i]!.startsWith(h)).toBe(true));
  });

  it('ISSUE-109：面の枠線の COM の太さが負（自動）の時は太さを比べず、色だけ比べる', () => {
    const c = makeCase(threeSeries({ style: 3 }));
    c.exp.ChartStyle = 3;
    c.exp.ChartArea.Line.Weight = -2863311533;
    expect(msgs(styleCheck, c)).toEqual([]);
    // 色が違えば差（太さは負の値のまま表示される）
    c.exp.ChartArea.Line.RGB = '123456';
    const m = msgs(styleCheck, c);
    expect(m).toHaveLength(1);
    expect(m[0]!.startsWith('面の枠線 got[')).toBe(true);
    // 負でなければ、今までどおり太さも比べる
    c.exp.ChartArea.Line.Weight = 5;
    c.exp.ChartArea.Line.RGB = rrggbb(
      c.scene.items[0]!.kind === 'rect' ? c.scene.items[0]!.stroke!.color : '#000000',
    );
    expect(msgs(styleCheck, c)).toHaveLength(1);
  });
});

describe('補助の刻み（MinorUnit）', () => {
  const valAx = (extra: object) =>
    ({
      kind: 'val',
      deleted: false,
      orientation: 'minMax',
      text: {},
      majorGridlines: {},
      ...extra,
    }) as ChartModel['valAxis'];

  it('補助の目盛線：本数から補助の刻み（主の 1/5）を出す。COM と合えば差なし・違えば差', () => {
    const c = makeCase(threeSeries({ valAxis: valAx({ minorGridlines: {} }) }));
    const ax = axisOf(c.scene, c.model)!;
    expect(minorUnitOf(c.scene, c.model, ax)).toBeCloseTo(ax.major / 5, 9);
    expect(msgs(minorCheck, c)).toEqual([]);
    c.exp.ValueAxis!.MinorUnit = ax.major / 2;
    expect(msgs(minorCheck, c)[0]).toMatch(/^MinorUnit got\[/);
  });

  it('補助の目盛の印（目盛線なし）：短い線分の数から出す（cross は 2 本ずつ）', () => {
    for (const mark of ['out', 'cross'] as const) {
      const c = makeCase(threeSeries({ valAxis: valAx({ minorTickMark: mark }) }));
      const ax = axisOf(c.scene, c.model)!;
      expect(minorUnitOf(c.scene, c.model, ax)).toBeCloseTo(ax.major / 5, 9);
      expect(msgs(minorCheck, c)).toEqual([]);
    }
  });

  it('明示の補助の刻み（主 10・補助 2.5）：横棒でも出せる', () => {
    const model = threeSeries({
      barDir: 'bar',
      valAxis: valAx({ minorGridlines: {}, majorUnit: 10, minorUnit: 2.5 }),
    });
    const c = makeCase(model);
    const ax = axisOf(c.scene, c.model)!;
    expect(ax.major).toBe(10);
    expect(minorUnitOf(c.scene, c.model, ax)).toBeCloseTo(2.5, 9);
  });

  it('補助の線も印も無ければ null（差：出ていない）', () => {
    const c = makeCase(threeSeries());
    expect(minorUnitOf(c.scene, c.model, axisOf(c.scene, c.model)!)).toBeNull();
    expect(msgs(minorCheck, c)[0]).toMatch(/^補助の目盛線・目盛の印が出ていない/);
  });
});

describe('目盛の印（C12）・マーカー（C13）', () => {
  it('c:majorTickMark・c:minorTickMark が無い軸は cross とみなす（スキーマの既定）。COM が cross なら差なし・他なら差', () => {
    const c = makeCase(threeSeries());
    expect(c.model.catAxis?.majorTickMark).toBeUndefined();
    c.exp.Ticks = { CatMajor: 4, CatMinor: 4, ValMajor: 4, ValMinor: 4 };
    expect(msgs(tickCheck, c)).toEqual([]);
    c.exp.Ticks = { CatMajor: 3, CatMinor: 4, ValMajor: 2, ValMinor: -4142 };
    expect(msgs(tickCheck, c)).toEqual([
      '項目軸の主 got[cross(無し)] want[out]',
      '数値軸の主 got[cross(無し)] want[in]',
      '数値軸の補助 got[cross(無し)] want[none]',
    ]);
    expect(msgs(tickCheck, { ...c, exp: { ...c.exp, Ticks: undefined } })).toEqual([
      'COM の Ticks が読めていない',
    ]);
  });

  it('折れ線のマーカー：系列が symbol none なら場面にマーカーが無く、COM の MarkerStyle が none なら差なし・有れば差。c:marker が無い系列は自動の円（C13）', () => {
    const base = lineModel(2, 5);
    const none = makeCase({
      ...base,
      series: base.series.map((s) => ({ ...s, marker: { symbol: 'none' as const } })),
    });
    expect(msgs(markerCheck, none)).toEqual([]); // 場面もマーカーなし・COM も -4142
    none.exp.Series[0]!.MarkerStyle = -4105;
    expect(msgs(markerCheck, none)[0]).toMatch(/^マーカー got\[false\] want\[true\]/);
    // グループの c:marker も系列の c:marker も無い：場面は自動の円。COM が none なら差
    const auto = makeCase(base);
    expect(msgs(markerCheck, auto)[0]).toMatch(/^マーカー got\[true\] want\[false\]/);
  });

  it('マーカーがある折れ線：大きさ（MarkerSize pt）を比べる', () => {
    const c = makeCase(lineModel(2, 5));
    const m = layoutChart({ ...c.model, lineMarkers: true }, W_PT * K, H_PT * K, undefined)!;
    const d: Case = { ...c, scene: m };
    d.exp.Series.forEach((s) => {
      s.MarkerStyle = -4105;
      s.MarkerSize = 6;
    });
    expect(msgs(markerCheck, d)).toEqual([]);
    d.exp.Series[0]!.MarkerSize = 9;
    expect(msgs(markerCheck, d)[0]).toMatch(/^大きさ got\[8\.00px\] want\[12\.00px\]/);
  });
});

describe('PNG の ink（C4）', () => {
  /** 場面の線から、PNG（縮尺 scale・背景白）が返すはずの summary を作る */
  function fakePng(c: Case, scale: number): PngChartResult {
    const f = scale / K;
    const st = axisStrokes(c.scene, c.model);
    // 1 画素に満たない線は 1 画素に描かれる（pngCheck と同じ）
    const ink = (s: ChartStroke): number[] =>
      [0, 1, 2].map(
        (k) =>
          Math.max(1, s.width * f) *
          s.opacity *
          (255 - parseInt(s.color.slice(1 + 2 * k, 3 + 2 * k), 16)),
      );
    const summary: PngChartResult['summary'] = {};
    if (st.major) summary.majorGrid = { count: 4, peak: 'D9D9D9', ink: ink(st.major), rows: 1 };
    if (st.catLine ?? st.major)
      summary.catAxisLine = {
        count: 1,
        peak: 'BFBFBF',
        ink: ink((st.catLine ?? st.major)!),
        rows: 1,
      };
    if (st.valLine)
      summary.valAxisLine = { count: 1, peak: '000000', ink: ink(st.valLine), rows: 1 };
    return {
      width: 0,
      height: 0,
      scale,
      scaleY: scale,
      bg: 'FFFFFF',
      hasAxes: true,
      lines: [],
      summary,
      chartStyle: null,
    };
  }

  it('場面の線から作った ink は差なし。ink がずれたら（色・太さ）差・場面に無い線は差', () => {
    const c = makeCase(threeSeries());
    const p = fakePng(c, 2);
    expect(msgs(pngCheck({ x: p }), c)).toEqual([]);
    const bad = {
      ...p,
      summary: { ...p.summary, majorGrid: { ...p.summary.majorGrid!, ink: [90, 90, 90] } },
    };
    expect(msgs(pngCheck({ x: bad }), c)[0]).toMatch(/^主の目盛線 ink got\[90,90,90\]/);
    // 数値軸の線が無い場面（a:ln が noFill）に、PNG が線を返したら差
    const noVal = makeCase(
      threeSeries({ valAxis: { ...threeSeries().valAxis!, line: { paint: { kind: 'none' } } } }),
    );
    const q = fakePng(noVal, 2);
    expect(msgs(pngCheck({ x: q }), noVal)).toEqual([]);
    const extra = {
      ...q,
      summary: {
        ...q.summary,
        valAxisLine: { count: 1, peak: '000000', ink: [200, 200, 200], rows: 1 },
      },
    };
    expect(msgs(pngCheck({ x: extra }), noVal)[0]).toMatch(/^数値軸の線：場面に線が無い/);
    expect(msgs(pngCheck({}), c)).toEqual([]); // PNG が無いグラフは対象外
  });

  it('場面にあって PNG に線が無い時は差', () => {
    const c = makeCase(threeSeries());
    const p = fakePng(c, 2);
    delete p.summary.majorGrid;
    expect(msgs(pngCheck({ x: p }), c)[0]).toBe('主の目盛線：PNG に線が無い（場面に有る）');
  });
});

describe('取り出しの補助', () => {
  it('縦棒の目盛線・軸線の線を場面から取る（主の目盛線・項目軸の線）', () => {
    const c = makeCase(threeSeries());
    const st = axisStrokes(c.scene, c.model);
    expect(st.major).toBeDefined();
    expect(st.catLine).toBeDefined();
    expect(st.minor).toBeUndefined();
    expect(gridBox(c.scene)!.w).toBeGreaterThan(100);
  });

  it('segments：M/L の組だけを線分にし、他の命令が混じれば null', () => {
    expect(segments('M0 0L10 0M0 5L10 5')).toEqual([
      [0, 0, 10, 0],
      [0, 5, 10, 5],
    ]);
    expect(segments('M0 0L10 0Z')).toBeNull();
    expect(segments('M0 0A5 5 0 0 1 10 0')).toBeNull();
  });

  it('sortVisual：行（縦の差が許容内）→ 左から', () => {
    const pts = [
      { x: 50, y: 10.5 },
      { x: 10, y: 10 },
      { x: 30, y: 40 },
      { x: 5, y: 41 },
    ];
    expect(sortVisual(pts, (p) => p, 3)).toEqual([
      { x: 10, y: 10 },
      { x: 50, y: 10.5 },
      { x: 5, y: 41 },
      { x: 30, y: 40 },
    ]);
  });

  it('judge：記録に無い差・悪化・消えた組を仕分ける', () => {
    const diffs: Diff[] = [
      { name: 'a', cat: 'c', msg: 'm1', metric: 5 },
      { name: 'b', cat: 'c', msg: 'm2', metric: 9 },
      { name: 'c', cat: 'c', msg: 'new' },
      { name: 'd', cat: 'other', msg: 'x' },
    ];
    const known = {
      'c:a': lim(5, 'ok'),
      'c:b': lim(8, 'worse'),
      'c:gone': same('x', 'stale'),
      'other:d': lim(1, 'other cat'),
    };
    expect(judge(diffs, 'c', known)).toEqual({
      unexpected: ['c: new'],
      worse: ['b: m2'],
      stale: ['c:gone'],
    });
  });

  it('judge：差が複数ある組は mixed（数で測れる差は max・測れない差は msgs）・listed（msgs だけ）で仕分ける', () => {
    const diffs: Diff[] = [
      { name: 'a', cat: 'c', msg: 'box', metric: 4 },
      { name: 'a', cat: 'c', msg: 'text' },
      { name: 'b', cat: 'c', msg: 'box', metric: 6 },
      { name: 'b', cat: 'c', msg: 'text' },
      { name: 'c', cat: 'c', msg: 't1' },
      { name: 'c', cat: 'c', msg: 't3' },
    ];
    const known = {
      'c:a': mixed(5, ['text'], 'ok'),
      'c:b': mixed(5, ['other'], 'worse'),
      'c:c': listed(['t1', 't2'], 'worse'),
    };
    expect(judge(diffs, 'c', known)).toEqual({
      unexpected: [],
      worse: ['b: box', 'b: text', 'c: t3'],
      stale: [],
    });
  });

  it('expectedLegend：円は varyColors が偽でも項目ごと（色が同じなので位置から categories の名前）・面の外の項目は数えない', () => {
    const model = barModel({
      type: 'pie',
      barDir: undefined,
      grouping: 'standard',
      varyColors: false,
      categories: ['A', 'B', 'C'],
      series: [series(0, 'S1', [1, 2, 3])],
      catAxis: undefined,
      valAxis: undefined,
      legend: { pos: 'r', overlay: false, deleted: [], text: {} },
    });
    const c = makeCase(model);
    const e = c.exp.Legend!.Entries;
    expect(e).toHaveLength(3); // 円は varyColors が偽でも項目ごと（C3）
    c.exp.Legend!.Entries = [...e, { ...e[0]!, Pos: { ...e[0]!.Pos, Left: W_PT + 10 } }];
    expect(expectedLegend(c.exp, c.model).names).toEqual(['A', 'B', 'C']);
  });

  it('GROUPS：検査の分類は 11 個・名前の接頭辞は重ならない', () => {
    const groups: Group[] = GROUPS;
    expect(groups.map((g) => g.cat)).toEqual([
      'c1',
      'c2',
      'c3',
      'c4',
      'c5',
      'c6',
      'c7',
      'minor',
      'c11',
      'c12',
      'c13',
    ]);
    const prefixes = groups.flatMap((g) => [...(g.charts ?? []), ...(g.defaults ?? [])]);
    expect(new Set(prefixes).size).toBe(prefixes.length);
  });
});
