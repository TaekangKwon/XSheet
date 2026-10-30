/**
 * グラフの描画（ISSUE-103）の試験用の `ChartModel`（`xlsx-chart-svg.test.ts`・`xlsx-image-layer.chart.test.ts`・`xlsx-chart.perf.test.ts` が使う）。
 * 製品のコードからは読み込まない（試験専用）。
 */
import type { ChartAxis, ChartModel, ChartSeries } from '../../src/core/types';

const catAx: ChartAxis = { kind: 'cat', deleted: false, orientation: 'minMax', text: {} };
const valAx: ChartAxis = {
  kind: 'val',
  deleted: false,
  orientation: 'minMax',
  text: {},
  majorGridlines: {},
};

/** 系列 1 つ */
export function series(
  idx: number,
  name: string,
  values: (number | null)[],
  over: Partial<ChartSeries> = {},
): ChartSeries {
  return { idx, name, values, ...over };
}

/** 縦棒（3 項目・1 系列）が既定の `ChartModel` */
export function barModel(over: Partial<ChartModel> = {}): ChartModel {
  return {
    type: 'bar',
    barDir: 'col',
    grouping: 'clustered',
    varyColors: false,
    categories: ['A', 'B', 'C'],
    series: [series(0, 'S1', [10, 20, 30])],
    catAxis: catAx,
    valAxis: valAx,
    autoTitleDeleted: false,
    chartArea: {},
    plotArea: {},
    text: {},
    dispBlanksAs: 'gap',
    ...over,
  };
}

/** 折れ線（系列数 × 点の数。値は系列と点で変わる）の `ChartModel` */
export function lineModel(seriesCount: number, points: number): ChartModel {
  return barModel({
    type: 'line',
    barDir: undefined,
    grouping: 'standard',
    categories: Array.from({ length: points }, (_, i) => `項目${i + 1}`),
    series: Array.from({ length: seriesCount }, (_, s) =>
      series(
        s,
        `系列${s + 1}`,
        Array.from({ length: points }, (_, i) => 50 + 30 * Math.sin((i + s * 7) / 40) + s * 3),
      ),
    ),
    legend: { pos: 'r', overlay: false, deleted: [], text: {} },
    title: { lines: ['売上'], text: {}, overlay: false },
  });
}
