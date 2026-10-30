// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createGeometry, type ChartResult, type SheetDrawing } from '../../src/core/index';
import { parseChart } from '../../src/core/chart';
import { DEFAULT_LIMITS } from '../../src/core/limits';
import { chartXml } from '../../src/testing/make-xlsx';
import { DEFAULT_THEME } from '../../src/core/theme';
import { emptyLayout } from '../../src/core/types';
import { EMU_PER_PX, layoutDrawings, viewWindow } from '../../src/render/drawings';
import { layoutChart, type ChartScene } from '../../src/render/chart-layout';
import { barModel, lineModel, series } from './chart-fixtures';
import { buildChartSvg } from '../../src/render/chart-svg';
import { createImageLayer } from '../../src/render/image-layer';

/**
 * グラフの重さの計測（ISSUE-103・ADR-0031 §決定 11）：系列 10・点 1,000 のグラフ（`xlsx-format.perf.test.ts` と同じ形）。
 * - グラフ 1 つの読み取り（parseChart）・配置（layoutChart）＋SVG（buildChartSvg）の時間。読み取りは XML の大きさ（この大きさで約 880KB）に比例し、
 *   配置＋SVG とは別に測る（読み取りは ISSUE-101 の解析。20ms の目安は配置＋SVG に当てる）。
 * - 50 個を置いた合成シートの図の層で、初回の表示（配置＋SVG＋DOM。読み取り済みの結果を返す）と 2 回目（キャッシュ）の出し入れの時間（jsdom の SVG の出し入れは重いので、目安は jsdom 向けに緩めてある。下の TD-103-31 の注）。
 * jsdom は実ブラウザより遅いので余裕を見た値。ライト・ダークでの縦横スクロールの 200ms 超のフレームは実機で目視（TV-45-4）。
 */

const SERIES = 10;
const POINTS = 1000;
const CHARTS = 50;

/** 系列 10・点 1,000 の折れ線グラフの XML（タイトル・凡例つき） */
function bigChartXml(): string {
  const cats = Array.from({ length: POINTS }, (_, i) => `項目${i + 1}`);
  return chartXml({
    group: 'line',
    grouping: 'standard',
    title:
      '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>売上の推移</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>',
    legend: '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>',
    series: Array.from({ length: SERIES }, (_, s) => ({
      name: `系列${s + 1}`,
      cat: { kind: 'str' as const, values: cats },
      values: Array.from(
        { length: POINTS },
        (_, i) => 50 + 30 * Math.sin((i + s * 7) / 40) + s * 3,
      ),
    })),
  });
}

/** 関数を runs 回測り、最小の時間（ms）を返す（ウォームアップ 1 回を除く） */
function bestOf(runs: number, fn: () => void): { best: number; times: number[] } {
  fn();
  const times: number[] = [];
  for (let i = 0; i < runs; i++) {
    const t0 = performance.now();
    fn();
    times.push(performance.now() - t0);
  }
  return { best: Math.min(...times), times };
}

describe('TD-103-30 重さ：グラフ 1 つ（系列 10・点 1,000）', () => {
  it('配置＋SVG が 60ms 以内（目安 20ms・実測 7ms 前後）。読み取り（約 880KB の XML）は参考値を出し、大きく外れないことだけ確かめる', () => {
    const xml = bigChartXml();
    const parsed = parseChart(xml, DEFAULT_THEME, DEFAULT_LIMITS, false);
    if (!parsed.ok) throw new Error(`読めない: ${parsed.reason}`);
    let elements = 0;
    const draw = (): void => {
      const scene = layoutChart(parsed.chart, 600, 360, DEFAULT_THEME);
      if (scene === null) throw new Error('場面が null');
      elements = buildChartSvg(scene, document).querySelectorAll('*').length;
    };
    const read = bestOf(5, () => void parseChart(xml, DEFAULT_THEME, DEFAULT_LIMITS, false));
    const both = bestOf(7, draw);
    console.log(
      `[TD-103-30] 系列 ${SERIES}・点 ${POINTS}（xml ${(xml.length / 1024).toFixed(0)}KB・svg 要素 ${elements}）：配置＋SVG 最小 ${both.best.toFixed(1)}ms（${both.times.map((t) => t.toFixed(0)).join('/')}）／読み取り 最小 ${read.best.toFixed(1)}ms（${read.times.map((t) => t.toFixed(0)).join('/')}）`,
    );
    expect(elements).toBeGreaterThan(10);
    expect(both.best).toBeLessThan(60); // 目安は 20ms（実測 7ms 前後）。全件実行の負荷で揺れるため 3 倍の余裕（xlsx-borders.perf の worst < 100 と同じ扱い）
    expect(read.best).toBeLessThan(600); // 目安ではなく、桁違いの劣化を見つけるための上限
  });
});

/** 場面の経路 `d` の長さの合計・経路の数 */
function dStats(scene: ChartScene): { length: number; paths: number } {
  let length = 0;
  let paths = 0;
  for (const it of scene.items)
    if (it.kind === 'path') {
      length += it.d.length;
      paths++;
    }
  return { length, paths };
}

describe('TD-106-30 重さ：間引き（ISSUE-106・ADR-0031 §決定 11.1）の前後の経路の長さ（系列 10・点 1,000）', () => {
  /** 上限は実測（間引き後）の 2 倍。間引きが効かなくなる退行（桁違いに増える）を見つける */
  const LIMIT = { line: 30_500, marker: 62_000, bar: 530_000 }; // 実測 15,245・30,938・264,817 の 2 倍
  const sineBar = barModel({
    categories: Array.from({ length: POINTS }, (_, i) => `項目${i + 1}`),
    series: Array.from({ length: SERIES }, (_, s) =>
      series(
        s,
        `系列${s + 1}`,
        Array.from({ length: POINTS }, (_, i) => 50 + 30 * Math.sin((i + s * 7) / 40) + s * 3),
      ),
    ),
    legend: { pos: 'r', overlay: false, deleted: [], text: {} },
  });
  /** マーカーなしの折れ線（系列が symbol none。c:marker が無い系列は自動の円になる・ISSUE-105 C13） */
  const plainLine = (): ReturnType<typeof lineModel> => {
    const m = lineModel(SERIES, POINTS);
    return { ...m, series: m.series.map((s) => ({ ...s, marker: { symbol: 'none' as const } })) };
  };
  const withCircles = (): ReturnType<typeof lineModel> => {
    const m = lineModel(SERIES, POINTS);
    return { ...m, lineMarkers: true };
  };

  it('折れ線・マーカー付き折れ線・棒の d の長さが間引きで減る（出力と上限）', () => {
    const out: string[] = [];
    for (const [name, model, limit] of [
      ['折れ線', plainLine(), LIMIT.line],
      ['折れ線＋円マーカー', withCircles(), LIMIT.marker],
      ['縦棒', sineBar, LIMIT.bar],
    ] as const) {
      const off = layoutChart(model, 600, 360, DEFAULT_THEME, { thin: false });
      const on = layoutChart(model, 600, 360, DEFAULT_THEME);
      if (off === null || on === null) throw new Error('場面が null');
      const a = dStats(off);
      const b = dStats(on);
      out.push(
        `${name}：d ${a.length} → ${b.length}（${(b.length / a.length).toFixed(3)}）・経路 ${a.paths} → ${b.paths}`,
      );
      expect(b.length).toBeLessThan(a.length);
      expect(b.length).toBeLessThanOrEqual(limit);
    }
    console.log(`[TD-106-30] 系列 ${SERIES}・点 ${POINTS}（600×360）\n  ${out.join('\n  ')}`);
  });
});

describe('TD-103-31 重さ：50 個のグラフを置いた合成シートの図の層', () => {
  const px = (n: number): number => n * EMU_PER_PX;
  const geo = createGeometry(emptyLayout(), 1000, 100);
  const flush = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

  it('初回の表示（配置＋SVG＋DOM）は 1 つ当たり 100ms 以内（実測 6.5ms）。2 回目（キャッシュ）の出し入れは 50 個で 100ms 以内（実測 12ms）', async () => {
    const xml = bigChartXml();
    const drawings: SheetDrawing[] = Array.from({ length: CHARTS }, (_, i) => ({
      anchor: {
        type: 'oneCell',
        from: { col: 0, colOff: 0, row: i * 15, rowOff: 0 },
        cx: px(600),
        cy: px(300),
      },
      kind: 'chart',
      name: `c${i}`,
      descr: '',
      mediaPath: null,
      chartPath: `xl/charts/chart${i}.xml`,
    }));
    const layer = createImageLayer();
    document.body.append(layer.element);
    // 読み取り済みの結果を返す（読み取りの時間は TD-103-30 で別に測る）
    const parsed: ChartResult = parseChart(xml, DEFAULT_THEME, DEFAULT_LIMITS, false);
    const readChart = async (): Promise<ChartResult> => parsed;
    layer.setSheet({
      drawings,
      boxes: layoutDrawings(geo, drawings),
      readImage: async () => ({ ok: false, reason: 'unsupported', label: null }),
      readChart,
    });
    // 50 個すべてが範囲に入る大きな窓
    const ALL = viewWindow(0, 0, 100000, 100000, 0, 0);
    const FAR = viewWindow(0, 10_000_000, 100000, 100000, 0, 0);

    const t0 = performance.now();
    layer.update(ALL);
    for (
      let i = 0;
      i < CHARTS * 3 && layer.element.querySelectorAll('svg.xs-chart').length < CHARTS;
      i++
    ) {
      await flush();
    }
    const first = performance.now() - t0;
    expect(layer.element.querySelectorAll('svg.xs-chart')).toHaveLength(CHARTS);

    // 2 回目：全部を外して全部を戻す（DOM の出し入れだけ。読み取りも配置も SVG の構築も起きない）
    const times: number[] = [];
    for (let i = 0; i < 6; i++) {
      const s = performance.now();
      layer.update(FAR);
      layer.update(ALL);
      times.push(performance.now() - s);
    }
    const best = Math.min(...times.slice(1));
    console.log(
      `[TD-103-31] グラフ ${CHARTS} 個：初回 ${first.toFixed(0)}ms（1 つ当たり ${(first / CHARTS).toFixed(1)}ms）／2 回目の出し入れ（外す＋戻す）最小 ${best.toFixed(1)}ms（${times.map((t) => t.toFixed(1)).join('/')}）`,
    );
    expect(layer.element.querySelectorAll('svg.xs-chart')).toHaveLength(CHARTS);
    expect(first / CHARTS).toBeLessThan(100); // 実測 6.5ms 前後
    // 作業指示書の目安は 50 個で 5ms。jsdom は SVG の要素の挿入・取り外しが重く（svg 50 個×要素 50 を素の replaceChildren で出し入れするだけで 7〜14ms。
    // 破線枠の div 50 個なら 0.5ms）、層の処理を足しても 12ms 前後になる。作り直しが起きていないこと（要素の同一性・readChart の回数）は
    // xlsx-image-layer.chart.test.ts で確かめ、ここは桁違いの劣化を見つけるための上限にする。実ブラウザの体感は TV-45-4（目視）。
    expect(best).toBeLessThan(100); // 実測 12ms 前後。桁違いの劣化だけを見る
    layer.dispose();
  }, 60_000);
});
