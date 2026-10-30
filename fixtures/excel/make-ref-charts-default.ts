/// <reference types="node" />
/**
 * ISSUE-104（C4・既定）：`c:spPr`・`c:txPr` を持たないグラフだけを並べた xlsx（`ref-charts-default.xlsx`）を合成して書き出す。
 * Excel が「スタイルの指定が無いグラフ」を開いた時の既定（グラフの面・枠線・目盛線・軸線・文字の色と大きさ・系列の色）を、
 * `make-ref-charts.ps1 -DefaultXlsx <このファイル>` が COM で読むための入力になる。
 *
 * 実行（リポジトリのルートから。vitest 付属の vite-node で TypeScript のまま走る）：
 *   npx vite-node -c vitest.config.ts src_V1/src/lib/xlsx/__fixtures__/excel/make-ref-charts-default.ts [出力先の xlsx（省略時はこのフォルダの ref-charts-default.xlsx）]
 *
 * 作りの約束：
 * - シート `Sheet1` の A2:A9＝項目 Cat1〜Cat8、B2:B9＝値。すべてのグラフが B 列を参照する（`chartXml` の系列の参照は B 列固定）ため、
 *   系列が複数でも値は同じにしてある（Excel は開く時にセルの値で保存値を更新する）。系列名は `c:v` の直書き（S1…）で区別する。
 * - グラフの名前（`xdr:cNvPr name`）＝`def_…`。COM の `ChartObject.Name` がそのまま基準値のキーになる。
 * - 本物の Excel が書く要素（`c:varyColors`・積み上げの `c:overlap`・軸の `c:axPos` など）は入れ、見た目を決める `c:spPr`・`c:txPr` と `c:style` だけ入れない。
 * - ISSUE-105：末尾に `def2_…`（C12 目盛の印の無い軸・C13 グループの c:marker の無い折れ線・C4 の `c:style` だけがある縦棒）を足した（先頭の 19 個は不変）。
 *   COM で読むのは `make-ref-charts2.ps1 -DefaultXlsx <このファイル>`（`Ticks`＝MajorTickMark・MinorTickMark・MarkerStyle・MarkerSize の要約つき）。
 */
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  buildXlsx,
  chartTitleXml,
  chartXml,
  graphicFrameXml,
  twoCellAnchor,
  wsDr,
  type ChartAxisSpec,
  type ChartSerSpec,
  type ChartSpec,
  type DrawingRelSpec,
} from '../../src/testing/make-xlsx';

const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CATS = ['Cat1', 'Cat2', 'Cat3', 'Cat4', 'Cat5', 'Cat6', 'Cat7', 'Cat8'];
const VALUES = [12, 19, 7, 15, 9, 13, 6, 17];

/** 系列 n 本（値は全系列で同じ。B 列参照のため）・点 pts 個 */
function series(n: number, pts = 5): ChartSerSpec[] {
  return Array.from({ length: n }, (_v, i) => ({
    name: `S${i + 1}`,
    nameLit: true,
    values: VALUES.slice(0, pts),
    formatCode: 'General',
    cat: { kind: 'str' as const, values: CATS.slice(0, pts) },
  }));
}

/** 本物の Excel が書く形の軸（見た目の `c:spPr`・`c:txPr` は無し） */
function axes(
  horizontal: boolean,
  percent: boolean,
  minor = false,
): { catAx: ChartAxisSpec; valAx: ChartAxisSpec } {
  return {
    catAx: {
      axPos: horizontal ? 'l' : 'b',
      delete: false,
      majorTickMark: 'none',
      minorTickMark: 'none',
      tickLblPos: 'nextTo',
      crosses: 'autoZero',
    },
    valAx: {
      axPos: horizontal ? 'b' : 'l',
      delete: false,
      majorGridlines: true,
      ...(minor ? { minorGridlines: true as const } : {}),
      numFmt: { code: percent ? '0%' : 'General', sourceLinked: !percent },
      majorTickMark: 'none',
      minorTickMark: minor ? 'out' : 'none',
      tickLblPos: 'nextTo',
      crosses: 'autoZero',
      crossBetween: 'between',
    },
  };
}

const LEGEND_R = '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>';
const TITLE = chartTitleXml(['Title'], '', '<c:overlay val="0"/>');

/** 10 種（G1）の定義 */
const TYPES: { key: string; spec: ChartSpec }[] = [
  {
    key: 'colClustered',
    spec: { group: 'bar', barDir: 'col', grouping: 'clustered', ...axes(false, false) },
  },
  {
    key: 'colStacked',
    spec: { group: 'bar', barDir: 'col', grouping: 'stacked', overlap: 100, ...axes(false, false) },
  },
  {
    key: 'colStacked100',
    spec: {
      group: 'bar',
      barDir: 'col',
      grouping: 'percentStacked',
      overlap: 100,
      ...axes(false, true),
    },
  },
  {
    key: 'barClustered',
    spec: { group: 'bar', barDir: 'bar', grouping: 'clustered', ...axes(true, false) },
  },
  {
    key: 'barStacked',
    spec: { group: 'bar', barDir: 'bar', grouping: 'stacked', overlap: 100, ...axes(true, false) },
  },
  {
    key: 'barStacked100',
    spec: {
      group: 'bar',
      barDir: 'bar',
      grouping: 'percentStacked',
      overlap: 100,
      ...axes(true, true),
    },
  },
  {
    key: 'line',
    spec: { group: 'line', grouping: 'standard', groupMarker: true, ...axes(false, false) },
  },
  {
    key: 'lineMarkers',
    spec: { group: 'line', grouping: 'standard', groupMarker: true, ...axes(false, false) },
  },
  { key: 'pie', spec: { group: 'pie', varyColors: true, firstSliceAng: 0 } },
  {
    key: 'doughnut',
    spec: { group: 'doughnut', varyColors: true, firstSliceAng: 0, holeSize: 50 },
  },
];

/** 書き出すグラフ（名前 → 指定）を組む */
export function defaultCharts(): Record<string, ChartSpec> {
  const out: Record<string, ChartSpec> = {};
  for (const t of TYPES) {
    const isLine = t.key === 'line';
    const isPie = t.key === 'pie';
    const isBarLine = t.spec.group === 'bar' || t.spec.group === 'line';
    out[`def_${t.key}`] = {
      ...t.spec,
      // 折れ線（記号なし）は系列に記号なしを明示。lineMarkers は系列の記号を指定しない（自動の形と大きさ・C8）
      series: series(isPie ? 1 : 3).map((s) =>
        isLine ? { ...s, marker: '<c:symbol val="none"/>' } : s,
      ),
      ...(isBarLine ? { varyColors: false } : {}),
      title: TITLE,
      autoTitleDeleted: false,
      legend: LEGEND_R,
    };
  }
  // 系列 8 本・円の点 8 個（7 系列目以降の既定の色・C3）
  out['def_colClustered_ser8'] = { ...out['def_colClustered']!, series: series(8) };
  out['def_pie_pt8'] = { ...out['def_pie']!, series: series(1, 8) };
  // varyColors の当たり方（C3）：棒 1 系列で 0・1・要素なし
  const one: ChartSpec = { ...out['def_colClustered']!, series: series(1), legend: LEGEND_R };
  out['def_colClustered_1ser_vary0'] = { ...one, varyColors: false };
  out['def_colClustered_1ser_vary1'] = { ...one, varyColors: true };
  const { varyColors: _unused, ...noVary } = one;
  out['def_colClustered_1ser_varyAbsent'] = noVary;
  // 自動のタイトル（C10）：c:tx の無い c:title／c:title なし＋autoTitleDeleted 0・1
  out['def_title_auto_1ser'] = {
    ...one,
    title: '<c:title><c:overlay val="0"/></c:title>',
    autoTitleDeleted: false,
  };
  out['def_title_absent_1ser'] = { ...one, title: undefined, autoTitleDeleted: false };
  out['def_title_deleted_1ser'] = { ...one, title: undefined, autoTitleDeleted: true };
  // 補助の目盛線（C4）
  out['def_line_minor'] = { ...out['def_line']!, ...axes(false, false, true) };

  // ---- ここから ISSUE-105 の追加（既存の 19 個の順番・名前は変えない。末尾に足す）----
  // C12：`c:majorTickMark`・`c:minorTickMark` の無い軸（COM の MajorTickMark・MinorTickMark で、Excel が既定をどう解くかを読む）
  const colMinor: ChartSpec = { ...out['def_colClustered']!, ...axes(false, false, true) };
  out['def2_tick_absent_col'] = dropTickMarks(out['def_colClustered']!, [
    'majorTickMark',
    'minorTickMark',
  ]);
  out['def2_tick_absent_bar'] = dropTickMarks(out['def_barClustered']!, [
    'majorTickMark',
    'minorTickMark',
  ]);
  out['def2_tick_minorAbsent_col'] = dropTickMarks(
    withTicks(out['def_colClustered']!, 'out', 'out'),
    ['minorTickMark'],
  );
  out['def2_tick_majorAbsent_col'] = dropTickMarks(colMinor, ['majorTickMark']);
  // C13：グループの `c:marker` が無い折れ線（系列の `c:marker` 無し／記号なし／丸の明示／グループが 0）
  const lineNoGroup: ChartSpec = { ...out['def_lineMarkers']! };
  delete lineNoGroup.groupMarker;
  out['def2_line_groupMarkerAbsent'] = {
    ...lineNoGroup,
    series: series(3), // 系列の c:marker 無し
  };
  out['def2_line_groupMarkerAbsent_symNone'] = {
    ...lineNoGroup,
    series: series(3).map((s) => ({ ...s, marker: '<c:symbol val="none"/>' })),
  };
  out['def2_line_groupMarkerAbsent_symCircle7'] = {
    ...lineNoGroup,
    series: series(3).map((s) => ({ ...s, marker: '<c:symbol val="circle"/><c:size val="7"/>' })),
  };
  out['def2_line_groupMarker0'] = {
    ...out['def_lineMarkers']!,
    groupMarker: false,
    series: series(3),
  };
  // C4：`c:style` だけがあって `c:spPr`・`c:txPr` が無い縦棒（Excel が c:style を開く時に使うか。COM の ChartStyle と各部の色を読む）
  for (const n of [1, 3, 10, 42]) {
    out[`def2_style_${n}_col`] = { ...out['def_colClustered']!, styleDirect: n };
  }
  return out;
}

/** 軸の目盛の印の指定を省いた形にする（C12）。`which` に挙げた要素を項目軸・数値軸の両方から除く */
function dropTickMarks(spec: ChartSpec, which: ('majorTickMark' | 'minorTickMark')[]): ChartSpec {
  const strip = (a: ChartAxisSpec | null | undefined): ChartAxisSpec | null | undefined => {
    if (a === undefined || a === null) return a;
    const c = { ...a };
    for (const k of which) delete c[k];
    return c;
  };
  return { ...spec, catAx: strip(spec.catAx), valAx: strip(spec.valAx) };
}

/** 軸の目盛の印を主・補助とも明示した形にする（C12 の「補助だけ無い」の元） */
function withTicks(spec: ChartSpec, major: string, minor: string): ChartSpec {
  const put = (a: ChartAxisSpec | null | undefined): ChartAxisSpec | null | undefined =>
    a === undefined || a === null ? a : { ...a, majorTickMark: major, minorTickMark: minor };
  return { ...spec, catAx: put(spec.catAx), valAx: put(spec.valAx) };
}

/** xlsx を組む（シート Sheet1 のデータ・チャートごとの drawing・rels） */
export async function buildDefaultXlsx(): Promise<Uint8Array> {
  const specs = defaultCharts();
  const names = Object.keys(specs);
  const charts: Record<string, string> = {};
  const rels: DrawingRelSpec[] = [];
  const anchors: string[] = [];
  names.forEach((name, i) => {
    const rid = `rId${i + 1}`;
    charts[`chart${i + 1}.xml`] = chartXml(specs[name]);
    rels.push({ id: rid, target: `../charts/chart${i + 1}.xml`, type: 'chart' });
    // 横 4 つ（1 つ＝8 列×14 行）で並べる。データの A〜B 列を避けて D 列から
    const col = 3 + 9 * (i % 4);
    const row = 14 * Math.floor(i / 4);
    anchors.push(
      twoCellAnchor(
        { col, row },
        { col: col + 8, row: row + 14 },
        graphicFrameXml(name, CHART_URI, i + 2, [0, 0, 0, 0], rid),
      ),
    );
  });
  return buildXlsx({
    sheets: [
      {
        name: 'Sheet1',
        rows: [['', 'S1'], ...CATS.map((c, i) => [c, VALUES[i]!])],
        drawing: { xml: wsDr(anchors), rels },
      },
    ],
    charts,
    // Excel は styles.xml に fonts・fills・borders が無いと開けないことがあるため、Excel の既定と同じ最小構成を入れる
    fonts: ['<font><sz val="11"/><name val="Calibri"/></font>'],
    fills: [
      '<fill><patternFill patternType="none"/></fill>',
      '<fill><patternFill patternType="gray125"/></fill>',
    ],
    borders: ['<border><left/><right/><top/><bottom/><diagonal/></border>'],
  });
}

const out = process.argv[2] ?? fileURLToPath(new URL('./ref-charts-default.xlsx', import.meta.url));
const bytes = await buildDefaultXlsx();
writeFileSync(out, bytes);
console.log(`wrote ${out} (${bytes.length} bytes, ${Object.keys(defaultCharts()).length} charts)`);
