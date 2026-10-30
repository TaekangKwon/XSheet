import { describe, expect, it } from 'vitest';
import { parseChart } from '../../src/core/chart';
import { resolveColor } from '../../src/core/color';
import { DEFAULT_LIMITS, MAX_FORMAT_CODE_LENGTH, type XlsxLimits } from '../../src/core/limits';
import { openXlsx } from '../../src/core/index';
import {
  buildXlsx,
  chartTitleXml,
  chartTxPrXml,
  chartXml,
  graphicFrameXml,
  twoCellAnchor,
  wsDr,
  type ChartCatSpec,
  type ChartSerSpec,
  type ChartSpec,
} from '../../src/testing/make-xlsx';
import { DEFAULT_THEME, type Theme } from '../../src/core/theme';
import type { ChartModel, ChartResult } from '../../src/core/types';

/** 合成した `ChartSpec`（または XML の文字列）を読む */
function read(
  spec: ChartSpec | string = {},
  opts: { limits?: Partial<XlsxLimits>; theme?: Theme; date1904?: boolean } = {},
): ChartResult {
  const xml = typeof spec === 'string' ? spec : chartXml(spec);
  return parseChart(
    xml,
    opts.theme ?? DEFAULT_THEME,
    { ...DEFAULT_LIMITS, ...opts.limits },
    opts.date1904 ?? false,
  );
}

/** 読めたグラフを取り出す（読めなければ試験を落とす） */
function chartOf(...args: Parameters<typeof read>): ChartModel {
  const r = read(...args);
  if (!r.ok) throw new Error(`グラフを読めない: ${r.reason}`);
  return r.chart;
}

const SALES: ChartSerSpec = { name: '売上', values: [10, 20, 30] };
const COST: ChartSerSpec = { name: '費用', values: [5, 15, 25] };
const CATS: ChartCatSpec = { kind: 'str', values: ['東', '西', '南'] };
const fill = (hex: string) => `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;

describe('TD-101-2 parseChart：G1 の 10 種（種類・向き・積み上げ・系列名・値・項目）', () => {
  const series = [{ ...SALES, cat: CATS }, COST];
  const cases: [string, ChartSpec, Partial<ChartModel>][] = [
    [
      '縦棒 clustered',
      { barDir: 'col', grouping: 'clustered' },
      { type: 'bar', barDir: 'col', grouping: 'clustered' },
    ],
    [
      '縦棒 stacked',
      { barDir: 'col', grouping: 'stacked' },
      { type: 'bar', barDir: 'col', grouping: 'stacked' },
    ],
    [
      '縦棒 percentStacked',
      { barDir: 'col', grouping: 'percentStacked' },
      { type: 'bar', barDir: 'col', grouping: 'percentStacked' },
    ],
    [
      '横棒 clustered',
      { barDir: 'bar', grouping: 'clustered' },
      { type: 'bar', barDir: 'bar', grouping: 'clustered' },
    ],
    [
      '横棒 stacked',
      { barDir: 'bar', grouping: 'stacked' },
      { type: 'bar', barDir: 'bar', grouping: 'stacked' },
    ],
    [
      '横棒 percentStacked',
      { barDir: 'bar', grouping: 'percentStacked' },
      { type: 'bar', barDir: 'bar', grouping: 'percentStacked' },
    ],
    [
      '折れ線 standard',
      { group: 'line', grouping: 'standard' },
      { type: 'line', grouping: 'standard' },
    ],
    [
      '折れ線 stacked',
      { group: 'line', grouping: 'stacked' },
      { type: 'line', grouping: 'stacked' },
    ],
    [
      '折れ線 マーカー付き',
      { group: 'line', grouping: 'standard', groupMarker: true },
      { type: 'line', grouping: 'standard', lineMarkers: true },
    ],
    [
      '円',
      { group: 'pie', varyColors: true },
      { type: 'pie', grouping: 'standard', varyColors: true },
    ],
    [
      'ドーナツ',
      { group: 'doughnut', varyColors: true },
      { type: 'doughnut', grouping: 'standard', varyColors: true },
    ],
  ];

  it.each(cases)('%s は ok で、種類・系列名・値・項目を持つ', (_label, spec, expected) => {
    const c = chartOf({ ...spec, series });
    expect(c).toMatchObject(expected);
    expect(c.categories).toEqual(['東', '西', '南']);
    expect(c.series.map((s) => [s.name, s.values])).toEqual([
      ['売上', [10, 20, 30]],
      ['費用', [5, 15, 25]],
    ]);
  });

  it('円・ドーナツは、軸の要素があっても軸を持たない。棒・折れ線は項目軸と数値軸を持つ', () => {
    const pie = chartOf({ group: 'pie', series, catAx: {}, valAx: {} });
    expect(pie.catAxis).toBeUndefined();
    expect(pie.valAxis).toBeUndefined();
    const bar = chartOf({ series });
    expect(bar.catAxis).toMatchObject({ kind: 'cat', deleted: false, orientation: 'minMax' });
    expect(bar.valAxis).toMatchObject({ kind: 'val', deleted: false });
  });

  it('既定：棒の向きは col・並べ方は棒 clustered・折れ線 standard、varyColors は要素が無ければ true（Excel 16 の照合：ISSUE-104）、空の点は gap', () => {
    const bar = chartOf({ series });
    expect(bar).toMatchObject({
      type: 'bar',
      barDir: 'col',
      grouping: 'clustered',
      varyColors: true,
      dispBlanksAs: 'gap',
      autoTitleDeleted: false,
    });
    // 種類にない並べ方（棒の standard・折れ線の clustered）は種類の既定に倒す
    expect(chartOf({ grouping: 'standard', series }).grouping).toBe('clustered');
    expect(chartOf({ group: 'line', grouping: 'clustered', series }).grouping).toBe('standard');
    expect(chartOf({ group: 'line', series }).grouping).toBe('standard');
  });

  it('種類ごとの値：gapWidth・overlap（棒）・firstSliceAng・holeSize（円・ドーナツ）・dispBlanksAs・varyColors', () => {
    const bar = chartOf({
      series,
      gapWidth: 150,
      overlap: -27,
      varyColors: true,
      dispBlanksAs: 'span',
    });
    expect(bar).toMatchObject({
      gapWidth: 150,
      overlap: -27,
      varyColors: true,
      dispBlanksAs: 'span',
    });
    expect(bar.firstSliceAng).toBeUndefined();
    const donut = chartOf({ group: 'doughnut', series, firstSliceAng: 90, holeSize: 60 });
    expect(donut).toMatchObject({ firstSliceAng: 90, holeSize: 60 });
    expect(donut.gapWidth).toBeUndefined();
    const pie = chartOf({ group: 'pie', series, firstSliceAng: 45, holeSize: 60 });
    expect(pie.firstSliceAng).toBe(45);
    expect(pie.holeSize).toBeUndefined();
    // 範囲外は丸める
    expect(chartOf({ series, gapWidth: 9999, overlap: -500 })).toMatchObject({
      gapWidth: 500,
      overlap: -100,
    });
    expect(chartOf({ series, dispBlanksAs: 'zero' }).dispBlanksAs).toBe('zero');
    expect(chartOf({ series, dispBlanksAs: null }).dispBlanksAs).toBe('gap');
  });

  it('折れ線：系列のマーカー（symbol・size・塗り）と c:smooth', () => {
    const c = chartOf({
      group: 'line',
      groupMarker: true,
      series: [
        {
          ...SALES,
          marker: `<c:symbol val="diamond"/><c:size val="9"/><c:spPr>${fill('FF0000')}</c:spPr>`,
          smooth: true,
        },
        { ...COST, marker: '<c:symbol val="none"/>', smooth: false },
        { name: '素', values: [1] },
      ],
    });
    expect(c.lineMarkers).toBe(true);
    expect(c.series[0]).toMatchObject({
      marker: {
        symbol: 'diamond',
        size: 9,
        fill: { kind: 'solid', color: { rgb: 0xff0000, alpha: 1 } },
      },
      smooth: true,
    });
    expect(c.series[1]).toMatchObject({ marker: { symbol: 'none' }, smooth: false });
    expect(c.series[2]!.marker).toBeUndefined();
    expect('smooth' in c.series[2]!).toBe(false);
  });
});

describe('TD-101-3 parseChart：値（保存値・疎な点）と項目', () => {
  it('疎な c:pt は null で埋め、数値でない値・無い点は null。ptCount の長さに並べる', () => {
    const c = chartOf({
      series: [{ name: 'a', values: [1, null, '#N/A', 4.5, null], ptCount: 6 }],
    });
    expect(c.series[0]!.values).toEqual([1, null, null, 4.5, null, null]);
    expect(c.categories).toEqual(['1', '2', '3', '4', '5', '6']);
  });

  it('c:numCache が無ければ c:numLit を使う。数値の文字（指数・符号・小数）を数にする', () => {
    const c = chartOf({
      series: [{ name: 'lit', lit: true, values: [1, '2.5', '-3', '1e2', '.5'] }],
    });
    expect(c.series[0]!.values).toEqual([1, 2.5, -3, 100, 0.5]);
  });

  it('c:order の順に並べ替える（無ければ出現順）。系列名は c:strRef の保存値か c:v', () => {
    const c = chartOf({
      series: [
        { name: 'C', order: 2, values: [3] },
        { name: 'A', order: 0, values: [1], nameLit: true },
        { name: 'B', order: 1, values: [2] },
      ],
    });
    expect(c.series.map((s) => s.name)).toEqual(['A', 'B', 'C']);
    expect(chartOf({ series: [{ values: [1] }] }).series[0]!.name).toBe('');
  });

  it('項目：c:strLit・c:numCache の書式（日付・小数）・書式の無い数値・多段（最も内側の段）・項目なし', () => {
    const one = (cat: ChartSerSpec['cat']) =>
      chartOf({ series: [{ name: 's', values: [1, 2], cat }] }).categories;
    expect(one({ kind: 'str', values: ['x', 'y'], lit: true })).toEqual(['x', 'y']);
    expect(one({ kind: 'num', values: [45000, 45001], formatCode: 'yyyy/m/d' })).toEqual([
      '2023/3/15',
      '2023/3/16',
    ]);
    expect(one({ kind: 'num', values: [1.5, 2], formatCode: '0.0' })).toEqual(['1.5', '2.0']);
    expect(one({ kind: 'num', values: [1, 2.25] })).toEqual(['1', '2.25']);
    expect(one({ kind: 'num', values: [7, 8], lit: true, formatCode: '0%' })).toEqual([
      '700%',
      '800%',
    ]);
    expect(one({ kind: 'multi', levels: [['東1', '東2'], ['東']] })).toEqual(['東1', '東2']);
    expect(one(undefined)).toEqual(['1', '2']);
    // 項目の保存値に欠けた点があれば空の文字
    expect(one({ kind: 'str', values: ['x', null] })).toEqual(['x', '']);
  });

  it('1904 年起点では日付の項目をその起点で書式化する', () => {
    const spec: ChartSpec = {
      series: [
        { name: 's', values: [1], cat: { kind: 'num', values: [45000], formatCode: 'yyyy/m/d' } },
      ],
    };
    expect(chartOf(spec).categories).toEqual(['2023/3/15']);
    expect(chartOf(spec, { date1904: true }).categories).toEqual(['2027/3/16']);
  });

  it('系列の長さが違うときは項目の数（無ければ最長の系列）に揃える（足りない点は null・余りは切る）', () => {
    const a = chartOf({
      series: [
        { name: 'a', values: [1, 2, 3], cat: { kind: 'str', values: ['x', 'y'] } },
        { name: 'b', values: [9] },
      ],
    });
    expect(a.categories).toEqual(['x', 'y']);
    expect(a.series.map((s) => s.values)).toEqual([
      [1, 2],
      [9, null],
    ]);
    const b = chartOf({
      series: [
        { name: 'a', values: [1] },
        { name: 'b', values: [1, 2, 3] },
      ],
    });
    expect(b.categories).toEqual(['1', '2', '3']);
    expect(b.series.map((s) => s.values)).toEqual([
      [1, null, null],
      [1, 2, 3],
    ]);
  });

  it('項目は最初の系列の c:cat だけを使う', () => {
    const c = chartOf({
      series: [
        { name: 'a', values: [1, 2], cat: { kind: 'str', values: ['p', 'q'] } },
        { name: 'b', values: [1, 2], cat: { kind: 'str', values: ['x', 'y'] } },
      ],
    });
    expect(c.categories).toEqual(['p', 'q']);
  });

  it('系列の書式コード（c:numCache の c:formatCode）を持つ', () => {
    const c = chartOf({ series: [{ name: 'a', values: [1], formatCode: '#,##0' }] });
    expect(c.series[0]!.formatCode).toBe('#,##0');
  });

  it('データラベル・近似曲線・誤差範囲があっても ok（無視する。P2）', () => {
    const extra =
      '<c:dLbls><c:showVal val="1"/></c:dLbls><c:trendline><c:trendlineType val="linear"/></c:trendline><c:errBars><c:errBarType val="both"/></c:errBars>';
    const c = chartOf({
      series: [{ ...SALES, extra }],
      groupExtra: '<c:dLbls><c:showVal val="1"/></c:dLbls>',
    });
    expect(c.series[0]!.values).toEqual([10, 20, 30]);
  });
});

describe('TD-101-4 parseChart：G1 の外・保存値なし・壊れ・上限', () => {
  const series = [SALES];

  it.each([
    ['面グラフ', { groupName: 'areaChart' }],
    ['散布図', { groupName: 'scatterChart' }],
    ['3D 棒', { groupName: 'bar3DChart' }],
    ['レーダー', { groupName: 'radarChart' }],
    ['円の 3D', { groupName: 'pie3DChart' }],
    [
      '組み合わせ（棒＋折れ線）',
      { extraGroups: ['<c:lineChart><c:grouping val="standard"/></c:lineChart>'] },
    ],
    ['数値軸が 2 本（第 2 軸）', { extraValAx: true }],
    ['3D 表示（c:view3D）', { view3D: true }],
    ['対数軸（c:logBase）', { valAx: { logBase: 10 } }],
  ] as [string, ChartSpec][])('%s は unsupported', (_label, spec) => {
    expect(read({ series, ...spec })).toEqual({ ok: false, reason: 'unsupported' });
  });

  it('種類の要素が 1 つも無いグラフは unsupported', () => {
    const xml = chartXml({ series }).replace(/<c:barChart>[\s\S]*<\/c:barChart>/, '');
    expect(read(xml)).toEqual({ ok: false, reason: 'unsupported' });
  });

  it('保存値の無い系列（numCache 無し・ptCount 0・c:val 無し）は除き、残りがあれば ok。全系列に無ければ no-data', () => {
    const c = chartOf({
      series: [{ name: 'no', noCache: true }, { name: 'zero', values: [] }, SALES],
    });
    expect(c.series.map((s) => s.name)).toEqual(['売上']);
    expect(
      read({
        series: [
          { name: 'no', noCache: true },
          { name: 'zero', values: [] },
        ],
      }),
    ).toEqual({ ok: false, reason: 'no-data' });
    expect(read({ series: [] })).toEqual({ ok: false, reason: 'no-data' });
    const noVal = chartXml({ series: [SALES] }).replace(/<c:val>[\s\S]*<\/c:val>/, '');
    expect(read(noVal)).toEqual({ ok: false, reason: 'no-data' });
  });

  it('壊れた XML（空・グラフでない・途中で切れた・閉じが無い）は broken。例外にしない', () => {
    expect(read('')).toEqual({ ok: false, reason: 'broken' });
    expect(read('<root><a/></root>')).toEqual({ ok: false, reason: 'broken' });
    expect(read('<c:chartSpace xmlns:c="x"><c:chart><c:plotArea>')).toEqual({
      ok: false,
      reason: 'broken',
    });
    const full = chartXml({ series });
    expect(read(full.slice(0, full.length - 20))).toEqual({ ok: false, reason: 'broken' });
    expect(read('<c:chartSpace xmlns:c="x"><c:chart/></c:chartSpace>')).toEqual({
      ok: false,
      reason: 'broken',
    });
    expect(read('<c:chartSpace xmlns:c="x"><c:chart><c:title/></c:chart></c:chartSpace>')).toEqual({
      ok: false,
      reason: 'broken',
    });
  });

  it('maxChartSeries を超えると too-large（超えない数は ok）', () => {
    const many: ChartSerSpec[] = Array.from({ length: 4 }, (_, i) => ({
      name: `s${i}`,
      values: [i],
    }));
    expect(read({ series: many }, { limits: { maxChartSeries: 3 } })).toEqual({
      ok: false,
      reason: 'too-large',
    });
    expect(read({ series: many }, { limits: { maxChartSeries: 4 } }).ok).toBe(true);
  });

  it('maxChartPoints（系列数 × 項目の数・1 系列の ptCount）を超えると too-large', () => {
    const s = (n: number): ChartSerSpec => ({
      name: 's',
      values: Array.from({ length: n }, (_, i) => i),
    });
    expect(read({ series: [s(5), s(5)] }, { limits: { maxChartPoints: 9 } })).toEqual({
      ok: false,
      reason: 'too-large',
    });
    expect(read({ series: [s(5), s(5)] }, { limits: { maxChartPoints: 10 } }).ok).toBe(true);
    // 巨大な ptCount は、配列を作る前に断る（点が 1 つでも）
    expect(read({ series: [{ name: 'a', values: [1], ptCount: 2_000_000_000 }] })).toEqual({
      ok: false,
      reason: 'too-large',
    });
    const hugeCat: ChartSerSpec = { name: 'a', values: [1], cat: { kind: 'str', values: ['x'] } };
    expect(
      read(
        chartXml({ series: [hugeCat] }).replace(
          '<c:ptCount val="1"/><c:pt idx="0"><c:v>x',
          '<c:ptCount val="999999999"/><c:pt idx="0"><c:v>x',
        ),
      ),
    ).toEqual({ ok: false, reason: 'too-large' });
    // ptCount が無くて idx が巨大
    const noCount = chartXml({ series: [{ name: 'a', values: [1] }] })
      .replace('<c:ptCount val="1"/>', '')
      .replace('idx="0"', 'idx="99999999999"');
    expect(read(noCount)).toEqual({ ok: false, reason: 'too-large' });
  });

  it('maxChartTextChars（系列名・項目・タイトルの文字の合計）を超えると too-large', () => {
    const named: ChartSpec = {
      series: [{ name: 'abcde', values: [1, 2], cat: { kind: 'str', values: ['xxx', 'yyy'] } }],
      title: chartTitleXml(['tttt']),
    };
    // 5 + 3 + 3 + 4 = 15 文字
    expect(read(named, { limits: { maxChartTextChars: 14 } })).toEqual({
      ok: false,
      reason: 'too-large',
    });
    expect(read(named, { limits: { maxChartTextChars: 15 } }).ok).toBe(true);
  });

  it('G1 の外の判定は、系列数・点の上限より先（上限を超える散布図も unsupported）', () => {
    const many: ChartSerSpec[] = Array.from({ length: 4 }, (_, i) => ({
      name: `s${i}`,
      values: [i],
    }));
    expect(
      read({ groupName: 'scatterChart', series: many }, { limits: { maxChartSeries: 1 } }),
    ).toEqual({ ok: false, reason: 'unsupported' });
  });
});

describe('TD-101-5 parseChart：軸', () => {
  it('数値軸・項目軸の明示の値（範囲・向き・間隔・交差・位置・目盛・文字の間引き・表示形式）を持つ', () => {
    const c = chartOf({
      series: [SALES],
      catAx: {
        axPos: 'b',
        orientation: 'maxMin',
        delete: true,
        tickLblPos: 'low',
        tickLblSkip: 2,
        majorTickMark: 'none',
        minorTickMark: 'cross',
        crosses: 'max',
        crossBetween: 'midCat',
      },
      valAx: {
        axPos: 'l',
        min: 0,
        max: 100.5,
        majorUnit: 20,
        minorUnit: 5,
        crossesAt: -3,
        numFmt: { code: '#,##0', sourceLinked: false },
        majorTickMark: 'out',
        tickLblPos: 'nextTo',
        crosses: 'autoZero',
        crossBetween: 'between',
      },
    });
    expect(c.catAxis).toEqual({
      kind: 'cat',
      deleted: true,
      pos: 'b',
      orientation: 'maxMin',
      tickLblPos: 'low',
      tickLblSkip: 2,
      majorTickMark: 'none',
      minorTickMark: 'cross',
      crosses: 'max',
      crossBetween: 'midCat',
      text: {},
    });
    expect(c.valAxis).toEqual({
      kind: 'val',
      deleted: false,
      pos: 'l',
      min: 0,
      max: 100.5,
      orientation: 'minMax',
      majorUnit: 20,
      minorUnit: 5,
      crossesAt: -3,
      crosses: 'autoZero',
      crossBetween: 'between',
      majorTickMark: 'out',
      tickLblPos: 'nextTo',
      numFmt: { code: '#,##0', sourceLinked: false },
      text: {},
    });
  });

  it('明示の無い項目は持たない（範囲・間隔・目盛線・書式・位置）', () => {
    const v = chartOf({ series: [SALES] }).valAxis!;
    for (const k of [
      'min',
      'max',
      'majorUnit',
      'minorUnit',
      'majorGridlines',
      'minorGridlines',
      'numFmt',
      'pos',
      'crosses',
      'tickLblSkip',
      'title',
    ]) {
      expect(k in v).toBe(false);
    }
  });

  it('目盛線：要素があれば持つ（線の明示は中に）。無ければ持たない', () => {
    const c = chartOf({
      series: [SALES],
      valAx: {
        majorGridlines: `<a:ln w="9525">${fill('D9D9D9')}<a:prstDash val="dash"/></a:ln>`,
        minorGridlines: true,
      },
    });
    expect(c.valAxis!.majorGridlines).toEqual({
      widthEmu: 9525,
      paint: { kind: 'solid', color: { rgb: 0xd9d9d9, alpha: 1 } },
      dash: 'dash',
    });
    expect(c.valAxis!.minorGridlines).toEqual({});
    expect(c.catAxis!.majorGridlines).toBeUndefined();
  });

  it('c:sourceLinked と c:numFmt：sourceLinked が真・偽、書式コードが長すぎれば General', () => {
    const linked = chartOf({
      series: [SALES],
      valAx: { numFmt: { code: '0.00', sourceLinked: true } },
    }).valAxis!.numFmt;
    expect(linked).toEqual({ code: '0.00', sourceLinked: true });
    const long = chartOf({
      series: [SALES],
      valAx: { numFmt: { code: '0'.repeat(MAX_FORMAT_CODE_LENGTH + 1) } },
    }).valAxis!.numFmt;
    expect(long).toEqual({ code: 'General', sourceLinked: false });
  });

  it('c:dateAx は項目軸（kind が date）。軸のタイトル・軸の線・目盛の文字（回転・大きさ）', () => {
    const c = chartOf({
      series: [SALES],
      dateAx: true,
      catAx: {
        txPr: chartTxPrXml(' sz="900" b="1"', fill('595959'), ' rot="-2700000" vert="horz"'),
        spPr: `<a:ln w="12700">${fill('808080')}</a:ln>`,
      },
      valAx: { title: chartTitleXml(['金額'], ' sz="1000"', '', ' rot="-5400000" vert="horz"') },
    });
    expect(c.catAxis).toMatchObject({
      kind: 'date',
      line: { widthEmu: 12700, paint: { kind: 'solid', color: { rgb: 0x808080, alpha: 1 } } },
      text: { sizePt: 9, bold: true, color: { rgb: 0x595959, alpha: 1 }, rotDeg: -45 },
    });
    expect(c.valAxis!.title).toMatchObject({
      lines: ['金額'],
      overlay: false,
      text: { sizePt: 10, rotDeg: -90 },
    });
  });

  it('文字の回転：±90° を超える値（Excel の「自動」の特別値）は持たない', () => {
    const c = chartOf({
      series: [SALES],
      catAx: { txPr: chartTxPrXml('', '', ' rot="-60000000" vert="horz"') },
    });
    expect('rotDeg' in c.catAxis!.text).toBe(false);
  });
});

describe('TD-101-6 parseChart：タイトル・凡例・manualLayout', () => {
  it('グラフのタイトル：c:rich の段落とラン（大きさ・太字・色・overlay）。autoTitleDeleted', () => {
    const c = chartOf({
      series: [SALES],
      title: chartTitleXml(
        ['売上推移', '（単位：円）'],
        ' sz="1400" b="1"',
        '<c:overlay val="1"/>',
      ),
    });
    expect(c.title!.lines).toEqual(['売上推移', '（単位：円）']);
    expect(c.title!.overlay).toBe(true);
    expect(c.title!.text).toEqual({ sizePt: 14, bold: true });
  });

  it('タイトルの文字の見た目は先頭のランの a:rPr（sz・b・色）、無ければ a:defRPr', () => {
    const rich = (rPr: string, def: string) =>
      `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr${def}/></a:pPr><a:r>${rPr}<a:t>T</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
    const a = chartOf({
      series: [SALES],
      title: rich(
        `<a:rPr lang="ja-JP" sz="1800" b="0"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr>`,
        ' sz="1400" b="1"',
      ),
    }).title!;
    expect(a.text).toEqual({ sizePt: 18, bold: false, color: { rgb: 0x00ff00, alpha: 1 } });
    const b = chartOf({
      series: [SALES],
      title: rich('<a:rPr lang="ja-JP"/>', ' sz="1400" b="1"'),
    }).title!;
    expect(b.text).toEqual({ sizePt: 14, bold: true });
  });

  it('タイトルの c:strRef の保存値・文字の無いタイトル（自動）・a:br での改行・空の段落', () => {
    const ref = `<c:title><c:tx><c:strRef><c:f>S!A1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>参照</c:v></c:pt></c:strCache></c:strRef></c:tx><c:overlay val="0"/></c:title>`;
    expect(chartOf({ series: [SALES], title: ref }).title).toMatchObject({
      lines: ['参照'],
      overlay: false,
    });
    expect(
      chartOf({ series: [SALES], title: '<c:title><c:overlay val="0"/></c:title>' }).title,
    ).toMatchObject({ lines: [] });
    const br = `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>上</a:t></a:r><a:br/><a:r><a:t>下</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
    expect(chartOf({ series: [SALES], title: br }).title!.lines).toEqual(['上', '下']);
    const empty = `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:endParaRPr lang="ja-JP"/></a:p></c:rich></c:tx></c:title>`;
    expect(chartOf({ series: [SALES], title: empty }).title!.lines).toEqual([]);
    expect(chartOf({ series: [SALES] }).title).toBeUndefined();
  });

  it('c:autoTitleDeleted（val 無しは真）', () => {
    expect(chartOf({ series: [SALES], autoTitleDeleted: true }).autoTitleDeleted).toBe(true);
    const bare = chartXml({ series: [SALES] }).replace(
      '<c:plotArea>',
      '<c:autoTitleDeleted/><c:plotArea>',
    );
    expect(chartOf(bare).autoTitleDeleted).toBe(true);
  });

  it('凡例：位置・overlay・削除した項目・文字・manualLayout。無ければ持たない', () => {
    const legend =
      '<c:legend><c:legendPos val="b"/><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry><c:legendEntry><c:idx val="2"/><c:delete val="0"/></c:legendEntry><c:legendEntry><c:idx val="3"/><c:delete/></c:legendEntry>' +
      '<c:layout><c:manualLayout><c:xMode val="edge"/><c:x val="0.5"/></c:manualLayout></c:layout><c:overlay val="1"/>' +
      chartTxPrXml(' sz="800"') +
      '</c:legend>';
    const l = chartOf({ series: [SALES], legend }).legend!;
    expect(l).toEqual({
      pos: 'b',
      overlay: true,
      deleted: [1, 3],
      text: { sizePt: 8 },
      layout: { xMode: 'edge', x: 0.5 },
    });
    // 既定の位置は r・未知の位置も r
    expect(
      chartOf({ series: [SALES], legend: '<c:legend><c:overlay val="0"/></c:legend>' }).legend,
    ).toEqual({ pos: 'r', overlay: false, deleted: [], text: {} });
    expect(
      chartOf({ series: [SALES], legend: '<c:legend><c:legendPos val="zz"/></c:legend>' }).legend!
        .pos,
    ).toBe('r');
    expect(chartOf({ series: [SALES] }).legend).toBeUndefined();
    for (const pos of ['r', 'l', 't', 'b', 'tr']) {
      expect(
        chartOf({ series: [SALES], legend: `<c:legend><c:legendPos val="${pos}"/></c:legend>` })
          .legend!.pos,
      ).toBe(pos);
    }
  });

  it('manualLayout（プロットの c:layout）：対象・モード・x・y・w・h の明示だけ持つ', () => {
    const ml = (inner: string) => `<c:manualLayout>${inner}</c:manualLayout>`;
    expect(
      chartOf({
        series: [SALES],
        plotLayout: ml(
          '<c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="factor"/><c:hMode val="edge"/><c:x val="0.1"/><c:y val="0.2"/><c:w val="0.7"/><c:h val="0.6"/>',
        ),
      }).plotLayout,
    ).toEqual({
      target: 'inner',
      xMode: 'edge',
      yMode: 'edge',
      wMode: 'factor',
      hMode: 'edge',
      x: 0.1,
      y: 0.2,
      w: 0.7,
      h: 0.6,
    });
    expect(
      chartOf({ series: [SALES], plotLayout: ml('<c:x val="-0.05"/><c:w val="abc"/>') }).plotLayout,
    ).toEqual({ x: -0.05 });
    expect(
      chartOf({ series: [SALES], plotLayout: ml('<c:layoutTarget val="zz"/>') }).plotLayout,
    ).toBeUndefined();
    expect(chartOf({ series: [SALES] }).plotLayout).toBeUndefined();
    // タイトルの c:layout
    const t = `<c:title><c:layout>${ml('<c:xMode val="edge"/><c:x val="0.3"/>')}</c:layout><c:overlay val="0"/></c:title>`;
    expect(chartOf({ series: [SALES], title: t }).title!.layout).toEqual({ xMode: 'edge', x: 0.3 });
  });
});

describe('TD-101-7 parseChart：色・線・文字・点ごとの上書き・c:style', () => {
  it('面（グラフ・プロット）の塗りと枠線：srgbClr・noFill・線の太さ・線種', () => {
    const c = chartOf({
      series: [SALES],
      chartSpPr: `${fill('FFFFFF')}<a:ln w="9525">${fill('D9D9D9')}<a:prstDash val="sysDot"/></a:ln>`,
      plotSpPr: '<a:noFill/><a:ln><a:noFill/></a:ln>',
    });
    expect(c.chartArea).toEqual({
      fill: { kind: 'solid', color: { rgb: 0xffffff, alpha: 1 } },
      line: {
        widthEmu: 9525,
        paint: { kind: 'solid', color: { rgb: 0xd9d9d9, alpha: 1 } },
        dash: 'sysDot',
      },
    });
    expect(c.plotArea).toEqual({ fill: { kind: 'none' }, line: { paint: { kind: 'none' } } });
    expect(chartOf({ series: [SALES] }).chartArea).toEqual({});
  });

  it('色：a:schemeClr（テーマ）＋lumMod・a:srgbClr・a:gradFill の代表色。テーマはブックのものを使う', () => {
    const theme: Theme = new Map(DEFAULT_THEME).set('accent1', 0x112233);
    const scheme =
      '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>';
    const expected = resolveColor(
      '<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>',
      theme,
    )!;
    const c = chartOf(
      {
        series: [
          { ...SALES, spPr: scheme },
          { ...COST, spPr: '<a:solidFill><a:schemeClr val="accent2"/></a:solidFill>' },
          {
            name: 'g',
            values: [1],
            spPr: '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs></a:gsLst></a:gradFill>',
          },
        ],
      },
      { theme },
    );
    expect(c.series[0]!.fill).toEqual({ kind: 'solid', color: expected });
    expect(expected.rgb).not.toBe(0x112233);
    expect(c.series[1]!.fill).toEqual({
      kind: 'solid',
      color: { rgb: DEFAULT_THEME.get('accent2'), alpha: 1 },
    });
    // グラデーションは代表色（stop が 1 つならその色）
    expect(c.series[2]!.fill).toEqual({ kind: 'solid', color: { rgb: 0x00ff00, alpha: 1 } });
    // 色が解けない塗りは持たない
    const bad = chartOf({
      series: [{ ...SALES, spPr: '<a:solidFill><a:srgbClr val="zzzzzz"/></a:solidFill>' }],
    });
    expect(bad.series[0]!.fill).toBeUndefined();
  });

  it('系列の線（c:spPr の a:ln）と塗りなし。明示が無い系列は塗りも線も持たない', () => {
    const c = chartOf({
      group: 'line',
      series: [
        { ...SALES, spPr: `<a:ln w="28575" cap="rnd">${fill('ED7D31')}<a:round/></a:ln>` },
        { ...COST, spPr: '<a:noFill/>' },
        { name: 'p', values: [1] },
      ],
    });
    expect(c.series[0]!.line).toEqual({
      widthEmu: 28575,
      paint: { kind: 'solid', color: { rgb: 0xed7d31, alpha: 1 } },
    });
    expect(c.series[0]!.fill).toBeUndefined();
    expect(c.series[1]!.fill).toEqual({ kind: 'none' });
    expect('fill' in c.series[2]! || 'line' in c.series[2]!).toBe(false);
  });

  it('c:dPt（点ごとの塗りの上書き）。点の位置が負・塗りの無いものは持たない', () => {
    const c = chartOf({
      group: 'pie',
      series: [
        {
          ...SALES,
          dPts: [
            { idx: 1, spPr: fill('00FF00') },
            { idx: 2, spPr: '<a:noFill/>' },
            { idx: 0, spPr: '<a:ln w="9525"/>' },
            { idx: -1, spPr: fill('0000FF') },
          ],
        },
      ],
    });
    expect([...c.series[0]!.points!]).toEqual([
      [1, { kind: 'solid', color: { rgb: 0x00ff00, alpha: 1 } }],
      [2, { kind: 'none' }],
    ]);
    expect(chartOf({ series: [SALES] }).series[0]!.points).toBeUndefined();
  });

  it('文字（c:chartSpace の c:txPr）：大きさ・太字・色。明示が無ければ持たない', () => {
    const c = chartOf({ series: [SALES], txPr: chartTxPrXml(' sz="1200" b="1"', fill('404040')) });
    expect(c.text).toEqual({ sizePt: 12, bold: true, color: { rgb: 0x404040, alpha: 1 } });
    expect(chartOf({ series: [SALES] }).text).toEqual({});
    // 大きさは 1〜400pt に丸める。b="0" は偽
    expect(chartOf({ series: [SALES], txPr: chartTxPrXml(' sz="99999" b="0"') }).text).toEqual({
      sizePt: 400,
      bold: false,
    });
    // 文字の noFill は透明
    expect(chartOf({ series: [SALES], txPr: chartTxPrXml('', '<a:noFill/>') }).text.color).toEqual({
      rgb: 0,
      alpha: 0,
    });
  });

  it('c:style：mc:AlternateContent の mc:Fallback の中（Choice の c14:style は読まない）・直下。101〜148 は 100 を引く。範囲外は丸めず持たない', () => {
    expect(chartOf({ series: [SALES], style: 2 }).style).toBe(2);
    expect(chartOf({ series: [SALES], styleDirect: 26 }).style).toBe(26);
    expect(chartOf({ series: [SALES], styleDirect: 102 }).style).toBe(2);
    expect(chartOf({ series: [SALES], styleDirect: 148 }).style).toBe(48);
    for (const v of [0, 49, 99, 100, 149, 999]) {
      expect(chartOf({ series: [SALES], styleDirect: v }).style).toBeUndefined();
    }
    expect(chartOf({ series: [SALES] }).style).toBeUndefined();
  });
});

describe('TD-109-1 parseChart：c14:style・c:roundedCorners（ISSUE-109）', () => {
  it('c14Style：mc:Choice の c14:style の val を持つ（style は今までどおり Fallback の値）。要素が無い・整数でなければ持たない', () => {
    const m = chartOf({ series: [SALES], style: 42 });
    expect(m.c14Style).toBe(142);
    expect(m.style).toBe(42);
    // Choice だけで Fallback の無いファイル：c14Style は持ち、style は持たない
    const choiceOnly = chartXml({ series: [SALES], style: 2 }).replace(
      /<mc:Fallback>.*?<\/mc:Fallback>/,
      '',
    );
    const c = chartOf(choiceOnly);
    expect(c.c14Style).toBe(102);
    expect(c.style).toBeUndefined();
    // c14:style と c:style が食い違っても style は Fallback のまま
    const mixedXml = chartXml({ series: [SALES], style: 2 }).replace(
      'c14:style val="102"',
      'c14:style val="145"',
    );
    expect(chartOf(mixedXml)).toMatchObject({ c14Style: 145, style: 2 });
    // 無い・直下の c:style だけ（mc:AlternateContent の外）・整数でない値は持たない
    expect('c14Style' in chartOf({ series: [SALES] })).toBe(false);
    expect('c14Style' in chartOf({ series: [SALES], styleDirect: 26 })).toBe(false);
    const bad = chartXml({ series: [SALES], style: 2 }).replace(
      'c14:style val="102"',
      'c14:style val="x"',
    );
    expect('c14Style' in chartOf(bad)).toBe(false);
  });

  it('roundedCorners：val の真偽をそのまま持つ。val が無い・未知の値は真（CT_Boolean の既定）。要素が無ければ持たない', () => {
    const rc = (roundedCorners?: string) => chartOf({ series: [SALES], roundedCorners });
    expect(rc('1').roundedCorners).toBe(true);
    expect(rc('true').roundedCorners).toBe(true);
    expect(rc('0').roundedCorners).toBe(false);
    expect(rc('false').roundedCorners).toBe(false);
    expect(rc('').roundedCorners).toBe(true);
    expect(rc('maybe').roundedCorners).toBe(true);
    expect('roundedCorners' in rc()).toBe(false);
  });

  it('2 項目を足しても他の項目は変わらない（要素の無い XML は 2 項目を持たず、持つ XML との差はこの 2 項目だけ）', () => {
    const base = chartOf({ series: [{ ...SALES, cat: CATS }, COST] });
    const withBoth = chartOf({
      series: [{ ...SALES, cat: CATS }, COST],
      style: 2,
      roundedCorners: '0',
    });
    const { c14Style, roundedCorners, style, ...rest } = withBoth;
    expect([c14Style, roundedCorners, style]).toEqual([102, false, 2]);
    expect(rest).toEqual(base);
  });
});

describe('TD-101-8 parseChart：ファイルの文字列の扱い', () => {
  it('書式コードが MAX_FORMAT_CODE_LENGTH を超えると General（項目・系列の書式コード）', () => {
    const long = '0'.repeat(MAX_FORMAT_CODE_LENGTH + 1);
    const c = chartOf({
      series: [
        {
          name: 's',
          values: [1.5],
          formatCode: long,
          cat: { kind: 'num', values: [1.5], formatCode: long },
        },
      ],
    });
    expect(c.categories).toEqual(['1.5']);
    expect(c.series[0]!.formatCode).toBe('General');
    // 上限ちょうどは書式として使う
    const ok = '0'.repeat(MAX_FORMAT_CODE_LENGTH - 2) + '.0';
    expect(
      chartOf({
        series: [
          { name: 's', values: [1], cat: { kind: 'num', values: [1.5], formatCode: '0.000' } },
        ],
      }).categories,
    ).toEqual(['1.500']);
    expect(ok.length).toBe(MAX_FORMAT_CODE_LENGTH);
  });

  it('系列名・項目・タイトルの文字以外の欄に、ファイルの文字列が入らない（列挙値・書式コード以外）', () => {
    const evil = 'EVIL_ENUM_script';
    const spec: ChartSpec = {
      barDir: evil,
      grouping: evil,
      series: [
        {
          name: 'EVIL_NAME',
          values: [1, 2],
          cat: { kind: 'str', values: ['EVIL_CAT', 'EVIL_CAT2'] },
          marker: `<c:symbol val="${evil}"/>`,
          spPr: `<a:ln><a:prstDash val="${evil}"/></a:ln><a:solidFill><a:srgbClr val="${evil}"/></a:solidFill>`,
        },
      ],
      catAx: {
        axPos: evil,
        tickLblPos: evil,
        majorTickMark: evil,
        crosses: evil,
        orientation: evil,
      },
      valAx: { crossBetween: evil, minorTickMark: evil },
      title: chartTitleXml(['EVIL_TITLE']),
      legend: `<c:legend><c:legendPos val="${evil}"/></c:legend>`,
      dispBlanksAs: evil,
      plotLayout: `<c:manualLayout><c:layoutTarget val="${evil}"/><c:xMode val="${evil}"/><c:x val="${evil}"/></c:manualLayout>`,
    };
    const c = chartOf(spec);
    // 文字の欄（系列名・項目・タイトルの行）だけが、ファイルの文字をそのまま持つ
    expect(c.series[0]!.name).toBe('EVIL_NAME');
    expect(c.categories).toEqual(['EVIL_CAT', 'EVIL_CAT2']);
    expect(c.title!.lines).toEqual(['EVIL_TITLE']);
    const json = JSON.stringify(c, (_k, v: unknown) => (v instanceof Map ? [...v] : v));
    const stripped = json
      .replace('EVIL_NAME', '')
      .replace('EVIL_CAT2', '')
      .replace('EVIL_CAT', '')
      .replace('EVIL_TITLE', '');
    expect(stripped).not.toContain('EVIL');
    // 既定に倒れている
    expect(c).toMatchObject({ barDir: 'col', grouping: 'clustered', dispBlanksAs: 'gap' });
    expect(c.legend!.pos).toBe('r');
  });

  it('例外のメッセージを作らない：失敗はすべて固定の理由（壊れた XML でも文字列を返さない）', () => {
    const r = read('<c:chartSpace>EVIL_SECRET<c:chart><c:plotArea><c:barChart><c:ser>');
    expect(r).toEqual({ ok: false, reason: 'broken' });
    expect(JSON.stringify(r)).not.toContain('EVIL');
  });

  it('XML の実体参照を戻す（系列名・項目・タイトル）', () => {
    const c = chartOf({
      series: [{ name: 'A&B', values: [1], cat: { kind: 'str', values: ['x<y'] } }],
      title: chartTitleXml(['"q"']),
    });
    expect(c.series[0]!.name).toBe('A&B');
    expect(c.categories).toEqual(['x<y']);
    expect(c.title!.lines).toEqual(['"q"']);
  });
});

describe('TD-101-9 readChart：本のレベル（テーマ・上限・合計・閉じた後）', () => {
  const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const series = (n: number): ChartSerSpec => ({
    name: 's',
    values: Array.from({ length: n }, (_, i) => i),
  });

  /** チャートを複数持つ 1 シートのブック（chart1.xml… が rId1…） */
  async function bookOf(
    specs: (ChartSpec | string)[],
    limits: Partial<XlsxLimits> = {},
    theme?: string,
  ) {
    const anchors = specs.map((_, i) =>
      twoCellAnchor(
        { col: 0, row: 0 },
        { col: 1, row: 1 },
        graphicFrameXml(`c${i}`, CHART_URI, 5 + i, [0, 0, 0, 0], `rId${i + 1}`),
      ),
    );
    const charts: Record<string, string> = {};
    specs.forEach(
      (s, i) => (charts[`chart${i + 1}.xml`] = typeof s === 'string' ? s : chartXml(s)),
    );
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: {
            xml: wsDr(anchors),
            rels: specs.map((_, i) => ({
              id: `rId${i + 1}`,
              target: `../charts/chart${i + 1}.xml`,
              type: 'chart',
            })),
          },
        },
        { name: 'T', rows: [['b']] },
      ],
      charts,
      ...(theme === undefined ? {} : { theme: { xml: theme } }),
    });
    return openXlsx(bytes, { limits });
  }

  it('openDrawings が返した chartPath を readChart で読める（ok・種類と値）', async () => {
    const book = await bookOf([{ series: [SALES] }]);
    const { drawings } = await book.openDrawings(1);
    const r = await book.readChart(1, drawings[0]!.chartPath!);
    expect(r.ok && r.chart.series[0]!.values).toEqual([10, 20, 30]);
  });

  it('テーマの色を使う（ブックのテーマ）', async () => {
    const { buildThemeXml } = await import('../../src/testing/make-xlsx');
    const book = await bookOf(
      [{ series: [{ ...SALES, spPr: '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>' }] }],
      {},
      buildThemeXml({ accent1: '<a:srgbClr val="112233"/>' }),
    );
    const r = await book.readChart(1, 'xl/charts/chart1.xml');
    expect(r.ok && r.chart.series[0]!.fill).toEqual({
      kind: 'solid',
      color: { rgb: 0x112233, alpha: 1 },
    });
  });

  it('成功の結果にブックのテーマを付ける（描画側が既定の系列色に使う。同じグラフを読み直しても付く。ISSUE-103）', async () => {
    const { buildThemeXml } = await import('../../src/testing/make-xlsx');
    const book = await bookOf(
      [{ series: [SALES] }],
      {},
      buildThemeXml({ accent1: '<a:srgbClr val="112233"/>' }),
    );
    const r1 = await book.readChart(1, 'xl/charts/chart1.xml');
    const r2 = await book.readChart(1, 'xl/charts/chart1.xml'); // 数え済みのグラフ
    for (const r of [r1, r2]) {
      expect(r.ok && r.theme?.get('accent1')).toBe(0x112233);
    }
    const noTheme = await bookOf([{ series: [SALES] }]);
    const d = await noTheme.readChart(1, 'xl/charts/chart1.xml');
    expect(d.ok && d.theme?.get('accent1')).toBe(0x4472c4); // 読めなければ既定のテーマ
  });

  it('zip に無い部品は broken。壊れた部品も broken。小さな部品の上限（maxSmallPartBytes）超えは too-large', async () => {
    const book = await bookOf(['<not-a-chart/>', { series: [SALES] }]);
    expect(await book.readChart(1, 'xl/charts/none.xml')).toEqual({ ok: false, reason: 'broken' });
    expect(await book.readChart(1, 'xl/charts/chart1.xml')).toEqual({
      ok: false,
      reason: 'broken',
    });
    const big = { series: [series(300)] };
    const small = await bookOf([big], { maxSmallPartBytes: 5000 });
    expect(await small.readChart(1, 'xl/charts/chart1.xml')).toEqual({
      ok: false,
      reason: 'too-large',
    });
  });

  it('maxSheetChartPoints：呼んだ順に点（系列数 × 項目の数）を数え、超えたら以後そのシートの新しいグラフは too-large。数え済みの再読み込みは読める', async () => {
    const book = await bookOf(
      [{ series: [series(4)] }, { series: [series(4)] }, { series: [series(1)] }],
      { maxSheetChartPoints: 6 },
    );
    const p = (n: number) => `xl/charts/chart${n}.xml`;
    expect((await book.readChart(1, p(1))).ok).toBe(true); // 4
    expect(await book.readChart(1, p(2))).toEqual({ ok: false, reason: 'too-large' }); // 4+4>6
    expect(await book.readChart(1, p(3))).toEqual({ ok: false, reason: 'too-large' }); // 超えた後は小さなグラフも
    expect((await book.readChart(1, p(1))).ok).toBe(true); // 数え済み
    // 数えるのはシートごと：別のシートは別の合計
    expect((await book.readChart(2, p(3))).ok).toBe(true);
  });

  it('呼び出しを並べても（Promise.all）順に数える。G1 の外のグラフは点を数えない', async () => {
    const book = await bookOf(
      [
        { series: [series(5)] },
        { series: [series(5)] },
        { groupName: 'areaChart', series: [series(5)] },
      ],
      { maxSheetChartPoints: 5 },
    );
    const rs = await Promise.all(
      [1, 2, 3].map((n) => book.readChart(1, `xl/charts/chart${n}.xml`)),
    );
    expect(rs.map((r) => (r.ok ? 'ok' : r.reason))).toEqual(['ok', 'too-large', 'too-large']);
  });

  it('同じグラフを重ねて読んでも（A→B→A の素早い切り替えで、破棄された読み込みと新しい読み込みが並ぶ）点は 1 回だけ数える', async () => {
    // 合計の上限は 6。グラフ 1（4 点）を 2 回数えると 8 > 6 で、後のグラフ 2（2 点）が誤って too-large になる
    const book = await bookOf([{ series: [series(4)] }, { series: [series(2)] }], {
      maxSheetChartPoints: 6,
    });
    const p = (n: number) => `xl/charts/chart${n}.xml`;
    const rs = await Promise.all([
      book.readChart(1, p(1)),
      book.readChart(1, p(1)),
      book.readChart(1, p(2)),
    ]);
    expect(rs.map((r) => (r.ok ? 'ok' : r.reason))).toEqual(['ok', 'ok', 'ok']);
  });

  it('範囲外のシート番号は RangeError、close 後は broken', async () => {
    const book = await bookOf([{ series: [SALES] }]);
    await expect(book.readChart(0, 'x')).rejects.toBeInstanceOf(RangeError);
    await expect(book.readChart(9, 'x')).rejects.toBeInstanceOf(RangeError);
    book.close();
    await expect(book.readChart(1, 'xl/charts/chart1.xml')).rejects.toMatchObject({
      kind: 'broken',
    });
  });
});

describe('TD-101-10 parseChart：細工した XML でも時間が線形（> の無い開きタグ・閉じの無い要素・深い入れ子・大量の点）', () => {
  /** 緩い固定の時間（線形なら数十 ms。二乗なら秒単位）。linear-time.test.ts と同じ考え方 */
  const LIMIT_MS = 1500;
  const N = 40_000;
  const timed = (xml: string): { ms: number; r: ChartResult } => {
    const t0 = performance.now();
    const r = read(xml);
    return { ms: performance.now() - t0, r };
  };
  const wrap = (inner: string) =>
    `<c:chartSpace xmlns:c="x"><c:chart><c:plotArea>${inner}</c:plotArea></c:chart></c:chartSpace>`;

  it('> の無い開きタグ・閉じの無い要素の大量の繰り返しでも、例外にならず時間が線形', () => {
    for (const inner of [
      '<c:barChart>' + '<c:ser '.repeat(N),
      '<c:barChart>' + '<c:ser><c:val><c:numRef><c:numCache>'.repeat(N),
      '<c:barChart><c:ser><c:val><c:numLit>' + '<c:pt idx="0"><c:v>'.repeat(N),
      '<c:barChart><c:ser><c:val><c:numLit>' + '<c:pt '.repeat(N),
      '<c:barChart>' + '<c:ser><c:tx><c:strRef>'.repeat(N) + '</c:barChart>',
      '<c:barChart><c:ser>' + '<c:dPt><c:idx val="0"/>'.repeat(N),
    ]) {
      const { ms, r } = timed(wrap(inner));
      expect(r.ok).toBe(false);
      expect(ms).toBeLessThan(LIMIT_MS);
    }
  });

  it('同じ名前の深い入れ子（c:spPr・a:ln・c:title）でも時間が線形', () => {
    const deep = '<c:spPr>'.repeat(N) + '</c:spPr>'.repeat(N);
    const xml = chartXml({
      series: [SALES],
      title: '<c:title>' + '<c:tx>'.repeat(N) + '</c:title>',
      chartSpPr: deep,
      plotSpPr: '<a:ln>'.repeat(N),
    });
    const { ms } = timed(xml);
    expect(ms).toBeLessThan(LIMIT_MS);
  });

  it('点・dPt・凡例の項目が大量でも（上限の範囲で）時間が線形', () => {
    const many = chartXml({
      series: [
        {
          name: 'a',
          values: Array.from({ length: 20_000 }, (_, i) => i),
          dPts: Array.from({ length: 20_000 }, (_, i) => ({ idx: i, spPr: fill('FF0000') })),
        },
      ],
      legend:
        '<c:legend>' +
        '<c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry>'.repeat(20_000) +
        '</c:legend>',
    });
    const { ms, r } = timed(many);
    expect(r.ok).toBe(true);
    expect(ms).toBeLessThan(LIMIT_MS);
  });
});

describe('TD-101-11 parseChart：レビュー指摘の修正（idx・項目軸の書式・確保の上限・値の範囲）', () => {
  it('系列の idx（c:idx。無ければ出現の位置）は、保存値の無い系列を除いた後・c:order で並べた後も保たれる', () => {
    const c = chartOf({
      series: [
        { name: 'A', idx: 7, order: 2, values: [1] },
        { name: 'drop', idx: 1, order: 0, noCache: true },
        { name: 'B', idx: 3, order: 1, values: [2] },
        { name: 'C', idx: null, order: 3, values: [3] },
      ],
    });
    expect(c.series.map((s) => [s.name, s.idx])).toEqual([
      ['B', 3],
      ['A', 7],
      ['C', 3],
    ]);
  });

  it('項目軸の c:numFmt が sourceLinked="0" で項目が数値なら、その書式で項目を作る（保存値の formatCode より優先）。それ以外は保存値の書式', () => {
    const cat: ChartCatSpec = { kind: 'num', values: [1, 2.5], formatCode: '0.00' };
    const base = (catAx: ChartSpec['catAx']) =>
      chartOf({ series: [{ name: 's', values: [1, 2], cat }], catAx }).categories;
    expect(base({ numFmt: { code: '0.0', sourceLinked: false } })).toEqual(['1.0', '2.5']);
    expect(base({ numFmt: { code: '0.0', sourceLinked: true } })).toEqual(['1.00', '2.50']);
    expect(base({})).toEqual(['1.00', '2.50']);
    // sourceLinked の属性が無いときも今のまま
    const noAttr = chartXml({
      series: [{ name: 's', values: [1, 2], cat }],
      catAx: { numFmt: { code: '0.0' } },
    }).replace(' sourceLinked="0"', '');
    expect(chartOf(noAttr).categories).toEqual(['1.00', '2.50']);
    // 文字の項目は書式の影響を受けない。日付の軸（c:dateAx）も同じ
    const str = chartOf({
      series: [{ name: 's', values: [1], cat: { kind: 'str', values: ['x'] } }],
      catAx: { numFmt: { code: '0.0', sourceLinked: false } },
    });
    expect(str.categories).toEqual(['x']);
    const date = chartOf({
      series: [
        { name: 's', values: [1], cat: { kind: 'num', values: [45000], formatCode: 'General' } },
      ],
      dateAx: true,
      catAx: { numFmt: { code: 'yyyy/m/d', sourceLinked: false } },
    });
    expect(date.categories).toEqual(['2023/3/15']);
    // 軸の numFmt 自体は今までどおり軸に持つ
    expect(
      chartOf({
        series: [{ name: 's', values: [1, 2], cat }],
        catAx: { numFmt: { code: '0.0', sourceLinked: false } },
      }).catAxis!.numFmt,
    ).toEqual({ code: '0.0', sourceLinked: false });
  });

  it('項目 1 個・系列 255 本・各 ptCount 100000 でも、項目数の長さだけ確保する（結果は 1 点ずつ・短時間）', () => {
    const series: ChartSerSpec[] = Array.from({ length: 255 }, (_, i) => ({
      name: 's' + i,
      values: [i],
      ptCount: 100_000,
      ...(i === 0 ? { cat: { kind: 'str', values: ['x'] } as ChartCatSpec } : {}),
    }));
    const t0 = performance.now();
    const c = chartOf({ series });
    const ms = performance.now() - t0;
    expect(c.categories).toEqual(['x']);
    expect(c.series).toHaveLength(255);
    expect(c.series.every((s, i) => s.values.length === 1 && s.values[0] === i)).toBe(true);
    expect(ms).toBeLessThan(500);
  });

  it('項目より後ろの点（idx が項目数以上）は捨てる。シート合計に数えるのも項目数 × 系列数', async () => {
    const c = chartOf({
      series: [{ name: 'a', values: [1, 2, 3, 4], cat: { kind: 'str', values: ['x', 'y'] } }],
    });
    expect(c.series[0]!.values).toEqual([1, 2]);
    const series: ChartSerSpec[] = Array.from({ length: 10 }, (_, i) => ({
      name: 's' + i,
      values: [i],
      ptCount: 100_000,
      ...(i === 0 ? { cat: { kind: 'str', values: ['x'] } as ChartCatSpec } : {}),
    }));
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: {
            xml: wsDr([
              twoCellAnchor(
                { col: 0, row: 0 },
                { col: 1, row: 1 },
                graphicFrameXml(
                  'c',
                  'http://schemas.openxmlformats.org/drawingml/2006/chart',
                  5,
                  [0, 0, 0, 0],
                  'rId1',
                ),
              ),
            ]),
            rels: [{ id: 'rId1', target: '../charts/chart1.xml', type: 'chart' }],
          },
        },
      ],
      charts: { 'chart1.xml': chartXml({ series }) },
    });
    // 合計 10 点（項目 1 × 系列 10）。上限 10 ちょうどなら読める
    const ok = await openXlsx(bytes, { limits: { maxSheetChartPoints: 10 } });
    expect((await ok.readChart(1, 'xl/charts/chart1.xml')).ok).toBe(true);
    const ng = await openXlsx(bytes, { limits: { maxSheetChartPoints: 9 } });
    expect(await ng.readChart(1, 'xl/charts/chart1.xml')).toEqual({
      ok: false,
      reason: 'too-large',
    });
  });

  it('majorUnit・minorUnit は有限の正の値だけ持つ（0・負・数でない値は持たない）', () => {
    const v = (majorUnit?: number, minorUnit?: number) =>
      chartOf({ series: [SALES], valAx: { majorUnit, minorUnit } }).valAxis!;
    expect(v(0.5, 2)).toMatchObject({ majorUnit: 0.5, minorUnit: 2 });
    for (const bad of [0, -5, NaN, Infinity]) {
      const a = v(bad, bad);
      expect('majorUnit' in a || 'minorUnit' in a).toBe(false);
    }
  });

  it('dispBlanksAs：val の無い要素は ECMA の既定 zero、要素が無ければ gap。holeSize は 10〜90 に丸める。dPt の idx が項目数以上なら捨てる', () => {
    const bare = chartXml({ series: [SALES] }).replace(
      '<c:dispBlanksAs val="gap"/>',
      '<c:dispBlanksAs/>',
    );
    expect(chartOf(bare).dispBlanksAs).toBe('zero');
    expect(chartOf({ series: [SALES], dispBlanksAs: null }).dispBlanksAs).toBe('gap');
    expect(chartOf({ group: 'doughnut', series: [SALES], holeSize: 5 }).holeSize).toBe(10);
    expect(chartOf({ group: 'doughnut', series: [SALES], holeSize: 95 }).holeSize).toBe(90);
    expect(chartOf({ group: 'doughnut', series: [SALES], holeSize: 50 }).holeSize).toBe(50);
    const d = chartOf({
      group: 'pie',
      series: [
        {
          ...SALES,
          dPts: [
            { idx: 2, spPr: fill('00FF00') },
            { idx: 3, spPr: fill('FF0000') },
          ],
        },
      ],
    });
    expect([...d.series[0]!.points!.keys()]).toEqual([2]);
    const allOut = chartOf({
      group: 'pie',
      series: [{ ...SALES, dPts: [{ idx: 3, spPr: fill('FF0000') }] }],
    });
    expect(allOut.series[0]!.points).toBeUndefined();
  });
});
