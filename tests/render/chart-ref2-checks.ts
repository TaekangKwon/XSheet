/// <reference types="node" />
/**
 * Excel 16 の COM の値（2 回目の照合・ISSUE-105）と、グラフの場面を比べる検査の部品と、場面の分け方（{@link GROUPS}）。
 * 読み込みと実行は `xlsx-chart-ref2.test.ts`（基準値のファイルが無い間は skip）、部品そのものの自己整合は `xlsx-chart-ref2-checks.test.ts`。
 * 取り出しの補助は `xlsx-chart-ref-scene.ts`。許容は 1 回目（`xlsx-chart-ref.test.ts`）と同じ：位置・大きさはグラフの面の 3%（ADR-0031 §決定 12）・軸は 1e-9。
 */
import type { PngChartResult } from '../../fixtures/excel/make-ref-charts-png.mjs';
import type { ChartModel, ChartTickMark } from '../../src/core/types';
import { CHART_DEFAULTS, type ChartScene } from '../../src/render/chart-layout';
import {
  axisOf,
  axisStrokes,
  boxDiff,
  cmds,
  expectedLegend,
  hexOf,
  K,
  legendBox,
  legendItems,
  minorUnitOf,
  near,
  plotOf,
  px,
  unionBox,
  valueLabelTexts,
  type ComChart,
  type Diff,
  type PathItem,
  type TextItem,
} from './chart-ref-scene';

/** 1 つのグラフの読み込み結果 */
export interface Case {
  name: string;
  group: 'charts' | 'defaults';
  exp: ComChart;
  model: ChartModel;
  scene: ChartScene;
}

/** 差の記録（msg と、数で測れるなら metric＝%） */
export type Add = (msg: string, metric?: number) => void;
export type CheckFn = (c: Case, add: Add) => void;

/** 位置・大きさの許容（グラフの面に対する割合・ADR-0031 §決定 12） */
export const TOL = 0.03;

const isRound = (m: ChartModel): boolean => m.type === 'pie' || m.type === 'doughnut';

/** 数値軸の最小・最大・刻み（C1・C2）：場面の目盛の文字と COM の MinimumScale・MaximumScale・MajorUnit */
export const axisCheck: CheckFn = ({ exp, scene, model }, add) => {
  if (isRound(model) || exp.ValueAxis === undefined) return;
  const ax = axisOf(scene, model);
  const w = exp.ValueAxis;
  if (ax === null) {
    add('目盛が出ていない');
    return;
  }
  if (
    near(ax.min, w.MinimumScale, 1e-9) &&
    near(ax.max, w.MaximumScale, 1e-9) &&
    near(ax.major, w.MajorUnit, 1e-9 * Math.max(1, Math.abs(w.MajorUnit)))
  )
    return;
  // ずれの大きさ：端は範囲に対する割合・刻みは刻みに対する割合の大きい方（%）
  const span = Math.max(1e-12, w.MaximumScale - w.MinimumScale);
  const metric =
    Math.max(
      Math.abs(ax.min - w.MinimumScale) / span,
      Math.abs(ax.max - w.MaximumScale) / span,
      Math.abs(ax.major - w.MajorUnit) / Math.max(1e-12, Math.abs(w.MajorUnit)),
    ) * 100;
  add(
    `got[${ax.min}..${ax.max} /${ax.major}] want[${w.MinimumScale}..${w.MaximumScale} /${w.MajorUnit}]`,
    metric,
  );
};

/** プロットの内側の位置と大きさ（C5・C7）：円・ドーナツは外接の正方形 */
export const plotCheck: CheckFn = ({ exp, scene, model }, add) => {
  const got = plotOf(scene, model);
  if (got === null) {
    add('プロットの外接矩形が取れない');
    return;
  }
  const d = boxDiff('内側', got, px(exp.PlotArea.Inside), scene, TOL);
  if (d !== null) add(d.msg, d.metric);
};

/**
 * 凡例の項目（C3・C5・C6・C11）：数・文字（COM の記号の色から引いた名前）・外接矩形。
 * COM の `LegendEntry` に文字の値は無いので、名前は {@link expectedLegend} の対応（記号の色 → 系列／点）で出す。
 * 引けない項目があれば、名前の照合は諦めてその旨を差に書く（数と外接矩形は照合する）。
 */
export const legendCheck: CheckFn = ({ exp, scene, model }, add) => {
  const got = legendItems(scene);
  if (!exp.HasLegend || exp.Legend === undefined) {
    if (got.length > 0) add('凡例が無いはずなのに出ている');
    return;
  }
  const want = expectedLegend(exp, model);
  const gotNames = got.map((g) => g.text);
  if (got.length !== want.boxes.length) {
    add(
      `項目の数 got[${got.length}: ${gotNames.join('|')}] want[${want.boxes.length}: ${want.names.map((n) => n ?? '?').join('|')}]`,
    );
    return;
  }
  if (want.names.some((n) => n === null)) {
    add(
      `COM の記号の色から名前を引けない項目がある want[${want.names.map((n) => n ?? '?').join('|')}]`,
    );
  } else if (gotNames.join('|') !== want.names.join('|')) {
    add(`項目の文字（見た目の順） got[${gotNames.join('|')}] want[${want.names.join('|')}]`);
  }
  const lb = legendBox(scene);
  if (lb !== null && want.boxes.length > 0) {
    const d = boxDiff('項目', lb, unionBox(want.boxes), scene, TOL);
    if (d !== null) add(d.msg, d.metric);
  }
};

/**
 * 既定の見た目（C4）：`c:style`（1〜48 の時）・グラフの面の塗りと枠線・タイトルの文字・軸の文字・凡例の文字の大きさと色。
 * 目盛線・軸線は COM の `Format.Line` が当てにならない（`c:spPr` の無いグラフで黒 0.5pt を返す）ので、PNG の画素で比べる（{@link pngCheck}）。
 */
export const styleCheck: CheckFn = ({ exp, scene, model }, add) => {
  const id = exp.ChartStyle;
  if (id !== undefined && id >= 1 && id <= 48 && (model.style ?? 2) !== id) {
    add(`c:style got[${model.style ?? '(無し→2)'}] want[${id}]`);
  }
  const ca = scene.items[0];
  if (ca?.kind === 'rect' && ca.width === scene.width) {
    const fill = exp.ChartArea.Fill;
    const visible = fill !== undefined && (fill.Visible === true || fill.Visible === -1);
    if (fill !== undefined && fill.Visible !== undefined) {
      if (visible && (ca.fill === null || ca.fill.color !== hexOf(fill.RGB))) {
        add(`面の塗り got[${ca.fill?.color ?? 'なし'}] want[${hexOf(fill.RGB)}]`);
      } else if (!visible && ca.fill !== null) {
        add(`面の塗り got[${ca.fill.color}] want[なし]`);
      }
    }
    const wl = exp.ChartArea.Line;
    // COM の Weight が負（-2863311533 など）は太さ「自動」。黒い面の上で見えない枠線なので、太さは比べず色だけ比べる（ISSUE-109）
    const autoWidth = wl.Weight < 0;
    if (ca.stroke === null) add('グラフの面の枠線が無い');
    else if (
      (!autoWidth && !near(ca.stroke.width, wl.Weight * K, 0.05)) ||
      ca.stroke.color !== hexOf(wl.RGB)
    ) {
      add(
        `面の枠線 got[${ca.stroke.width.toFixed(2)}px ${ca.stroke.color}] want[${(wl.Weight * K).toFixed(2)}px ${hexOf(wl.RGB)}]`,
      );
    }
  }
  if (exp.HasTitle && exp.Title !== undefined) {
    const title = scene.items.find(
      (i): i is TextItem =>
        i.kind === 'text' &&
        i.baseline === 'middle' &&
        i.anchor === 'middle' &&
        i.text === exp.Title!.Text,
    );
    const f = exp.Title.Font;
    const bold = f.Bold === true || f.Bold === -1;
    if (
      title !== undefined &&
      (!near(title.size, f.Size * K, 0.05) ||
        title.bold !== bold ||
        title.color.color !== hexOf(f.RGB))
    ) {
      add(
        `タイトル got[${title.size.toFixed(2)}px bold=${title.bold} ${title.color.color}] want[${(f.Size * K).toFixed(2)}px bold=${bold} ${hexOf(f.RGB)}]`,
      );
    }
  }
  if (!isRound(model) && exp.ValueAxis !== undefined) {
    const lab = valueLabelTexts(scene, model)[0];
    const f = exp.ValueAxis.TickLabels.Font;
    if (
      lab !== undefined &&
      (!near(lab.size, f.Size * K, 0.05) || lab.color.color !== hexOf(f.RGB))
    ) {
      add(
        `軸の文字 got[${lab.size.toFixed(2)}px ${lab.color.color}] want[${(f.Size * K).toFixed(2)}px ${hexOf(f.RGB)}]`,
      );
    }
  }
  if (exp.HasLegend && exp.Legend !== undefined) {
    const lab = scene.items.find(
      (i): i is TextItem => i.kind === 'text' && i.anchor === 'start' && i.baseline === 'middle',
    );
    const f = exp.Legend.Font;
    if (
      lab !== undefined &&
      (!near(lab.size, f.Size * K, 0.05) || lab.color.color !== hexOf(f.RGB))
    ) {
      add(
        `凡例の文字 got[${lab.size.toFixed(2)}px ${lab.color.color}] want[${(f.Size * K).toFixed(2)}px ${hexOf(f.RGB)}]`,
      );
    }
  }
};

/**
 * 補助の刻み（MinorUnit）：数値軸の最小・最大・主の刻みに加え、場面の補助の線・印から出した補助の刻みを COM の MinorUnit と比べる。
 * 数値軸の線が none（Excel 2013 以降の既定）で補助の目盛線も無い時は、場面が補助の印を描かないのが正しい
 * （印は軸の線の書式で描く。3 回目の照合の PNG：線 none の軸には、主・補助とも印が出ない。ISSUE-110）ので、補助の線・印が無いことを差にしない。
 */
export const minorCheck: CheckFn = (c, add) => {
  axisCheck(c, add);
  const w = c.exp.ValueAxis;
  if (w === undefined || w.MinorUnit === undefined) {
    add('COM の MinorUnit が読めていない');
    return;
  }
  const ax = axisOf(c.scene, c.model);
  if (ax === null) return;
  const got = minorUnitOf(c.scene, c.model, ax);
  if (got === null) {
    const va = c.model.valAxis;
    if (va?.line?.paint?.kind === 'none' && va.minorGridlines === undefined) return;
    add(`補助の目盛線・目盛の印が出ていない want[MinorUnit=${w.MinorUnit}]`);
  } else if (!near(got, w.MinorUnit, 1e-9 * Math.max(1, Math.abs(w.MinorUnit)))) {
    add(
      `MinorUnit got[${got}] want[${w.MinorUnit}]`,
      (Math.abs(got - w.MinorUnit) / Math.abs(w.MinorUnit)) * 100,
    );
  }
};

/** COM の目盛の印の定数 → 語（xlTickMarkNone -4142・Inside 2・Outside 3・Cross 4） */
const TICK_WORD: Record<number, ChartTickMark> = { [-4142]: 'none', 2: 'in', 3: 'out', 4: 'cross' };

/**
 * 目盛の印（C12）：`c:majorTickMark`・`c:minorTickMark` が無い軸で、COM の MajorTickMark・MinorTickMark（Excel が既定をどう解いたか）と、
 * 読み取りの結果（無ければ描画側が {@link CHART_DEFAULTS.tickMark}＝交差とする）を、項目軸・数値軸の主・補助の 4 つで比べる。
 */
export const tickCheck: CheckFn = ({ exp, model }, add) => {
  const t = exp.Ticks;
  if (t === undefined) {
    add('COM の Ticks が読めていない');
    return;
  }
  const eff = (m: ChartTickMark | undefined): ChartTickMark => m ?? CHART_DEFAULTS.tickMark;
  const pairs: [string, number | undefined, ChartTickMark | undefined][] = [
    ['項目軸の主', t.CatMajor, model.catAxis?.majorTickMark],
    ['項目軸の補助', t.CatMinor, model.catAxis?.minorTickMark],
    ['数値軸の主', t.ValMajor, model.valAxis?.majorTickMark],
    ['数値軸の補助', t.ValMinor, model.valAxis?.minorTickMark],
  ];
  for (const [label, want, got] of pairs) {
    if (want === undefined) continue;
    const w = TICK_WORD[want] ?? `?${want}`;
    if (eff(got) !== w)
      add(`${label} got[${eff(got)}${got === undefined ? '(無し)' : ''}] want[${w}]`);
  }
};

/**
 * マーカー（C13）：グループの `c:marker` が無い折れ線で、COM の MarkerStyle（-4142＝なし）の有無と MarkerSize を場面のマーカー（弧だけの経路）と比べる。
 */
export const markerCheck: CheckFn = ({ exp, scene }, add) => {
  const markers = scene.items.filter(
    (it): it is PathItem => it.kind === 'path' && /A/.test(it.d) && !/L/.test(it.d),
  );
  const hasMarker = exp.Series.some((s) => s.MarkerStyle !== -4142);
  if (hasMarker !== markers.length > 0) {
    add(
      `マーカー got[${markers.length > 0}] want[${hasMarker}] (MarkerStyle ${exp.Series.map((s) => s.MarkerStyle).join()})`,
    );
    return;
  }
  if (!hasMarker) return;
  const diam = cmds(markers[0]!.d).find((x) => x.c === 'A')!.n[0]! * 2;
  const want = exp.Series[0]!.MarkerSize * K;
  if (!near(diam, want, 0.1)) add(`大きさ got[${diam.toFixed(2)}px] want[${want.toFixed(2)}px]`);
};

/** 線の ink の許容：チャンネルごと max(10, 期待の 15%)（画素数×0〜255 の単位） */
const inkTol = (want: number): number => Math.max(10, 0.15 * Math.abs(want));

/**
 * 目盛線・軸線の色と太さ（C4）：PNG の線の ink（背景−画素の行ごとの合計）と、場面の線の期待の ink（太さ × PNG の縮尺 × (背景−色) × 不透明度）を比べる。
 * 主の目盛線・項目軸の線・数値軸の線を見る。項目軸の線が場面に無い時は、0 の位置に重なる主の目盛線の色が出るはずなので、それと比べる。
 * アンチエイリアスで細い線が薄まっても ink は保たれる（面積の保存）ので、色と太さの組を 1 つの量で比べられる。
 * Excel は 1 画素に満たない線（0.5pt の目盛線は 120dpi で 0.83 画素）を 1 画素の線に描く（PNG で確認）ので、PNG での太さは max(1, 太さ × 縮尺) 画素とする。
 */
export function pngCheck(png: Record<string, PngChartResult>): CheckFn {
  return (c, add) => {
    const p = png[c.name];
    if (p === undefined || !p.hasAxes) return;
    const strokes = axisStrokes(c.scene, c.model);
    const bg = [0, 2, 4].map((i) => parseInt(p.bg.slice(i, i + 2), 16));
    const f = p.scale / K; // 場面の 1px あたりの PNG の画素数
    const expectedInk = (s: { width: number; color: string; opacity: number }): number[] =>
      [0, 1, 2].map(
        (k) =>
          Math.max(1, s.width * f) *
          s.opacity *
          (bg[k]! - parseInt(s.color.slice(1 + 2 * k, 3 + 2 * k), 16)),
      );
    const roles: [string, 'majorGrid' | 'catAxisLine' | 'valAxisLine', typeof strokes.major][] = [
      ['主の目盛線', 'majorGrid', strokes.major],
      ['項目軸の線', 'catAxisLine', strokes.catLine ?? strokes.major],
      ['数値軸の線', 'valAxisLine', strokes.valLine],
    ];
    for (const [label, role, stroke] of roles) {
      const got = p.summary[role];
      if (got === undefined && stroke === undefined) continue;
      if (got === undefined) {
        // 数値軸の線は、場面にあって PNG に無いなら（PNG の検出は縁の 1 本）差とする
        add(`${label}：PNG に線が無い（場面に有る）`);
        continue;
      }
      if (stroke === undefined) {
        add(`${label}：場面に線が無い（PNG の色 ${got.peak}・ink ${got.ink.join(',')}）`);
        continue;
      }
      const want = expectedInk(stroke);
      const bad = got.ink.some((v, k) => Math.abs(v - want[k]!) > inkTol(want[k]!));
      if (bad) {
        add(
          `${label} ink got[${got.ink.map((v) => v.toFixed(0)).join(',')}] (PNG の色 ${got.peak}) want[${want.map((v) => v.toFixed(0)).join(',')}] (場面 ${stroke.color} ${stroke.width.toFixed(2)}px)`,
        );
      }
    }
  };
}

// ---- 場面の分け方 ----

/** 検査の 1 つのまとまり：`cat`（TD-105-n の分類の名前）・名前の接頭辞（charts＝ref-charts2.xlsx／defaults＝ref-charts-default.xlsx）・検査 */
export interface Group {
  cat: string;
  label: string;
  charts?: string[];
  defaults?: string[];
  checks: CheckFn[];
}

/** 場面の分け方（ADR-0031 §決定 16「次の COM で確かめる事項」） */
export const GROUPS: Group[] = [
  {
    cat: 'c1',
    label: 'TD-105-1 負を含む横棒・7 桁以上の目盛（C1）',
    charts: ['c1neg_', 'c1mil_'],
    checks: [axisCheck, plotCheck],
  },
  {
    cat: 'c2',
    label: 'TD-105-2 全部が負の 100% 積み上げ（C2）',
    charts: ['c2allneg_'],
    checks: [axisCheck, plotCheck],
  },
  {
    cat: 'c3',
    label: 'TD-105-3 varyColors=0 の円・ドーナツの凡例（C3・C6）',
    charts: ['c3vary'],
    checks: [legendCheck, plotCheck],
  },
  {
    cat: 'c4',
    label: 'TD-105-4 c:style が 2 以外の面・文字（C4）',
    charts: ['c4style_'],
    defaults: ['def2_style_'],
    checks: [styleCheck, plotCheck],
  },
  {
    cat: 'c5',
    label: 'TD-105-5 右上の凡例×大きさ（C5）',
    charts: ['c5corner_'],
    checks: [axisCheck, plotCheck, legendCheck],
  },
  {
    cat: 'c6',
    label: 'TD-105-6 13 項目の横並びの凡例（C6）',
    charts: ['c6row13_'],
    checks: [legendCheck, plotCheck],
  },
  {
    cat: 'c7',
    label: 'TD-105-7 項目軸の回転・折り返しの条件（C7）',
    charts: ['c7sw_'],
    checks: [plotCheck],
  },
  {
    cat: 'minor',
    label: 'TD-105-8 補助の刻み（MinorUnit）',
    charts: ['minor_'],
    checks: [minorCheck],
  },
  {
    cat: 'c11',
    label: 'TD-105-9 凡例の項目の削除・系列の並べ替え（C11）',
    charts: ['c11_'],
    checks: [legendCheck],
  },
  {
    cat: 'c12',
    label: 'TD-105-10 目盛の印が無い軸の既定（C12）',
    defaults: ['def2_tick_'],
    checks: [tickCheck],
  },
  {
    cat: 'c13',
    label: 'TD-105-11 グループの c:marker が無い折れ線（C13）',
    defaults: ['def2_line_'],
    checks: [markerCheck],
  },
];

/** 1 つのまとまりの検査を、該当するグラフ全部に当てて差を返す */
export function runGroup(g: Group, cases: Case[], extra: CheckFn[] = []): Diff[] {
  const out: Diff[] = [];
  for (const c of cases) {
    const prefixes = (c.group === 'charts' ? g.charts : g.defaults) ?? [];
    if (!prefixes.some((p) => c.name.startsWith(p))) continue;
    const add: Add = (msg, metric) => out.push({ name: c.name, cat: g.cat, msg, metric });
    for (const check of [...g.checks, ...extra]) check(c, add);
  }
  return out;
}
