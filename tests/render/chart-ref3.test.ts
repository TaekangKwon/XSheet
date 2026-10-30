/// <reference types="node" />
/**
 * Excel 16 の COM が返した値との、グラフの 3 回目の照合（ISSUE-110・ADR-0031 §決定 16「次の COM で確かめる事項」）。
 * 基準値は `__fixtures__/excel/make-ref-charts3.ps1` の出力（ユーザー承認後に実行して取り込む）：
 * `ref-charts3.xlsx`（166 個）＋`ref-charts3.expected.json`（その COM の値）＋`ref-charts3-png.json`（PNG の線・角・目盛の印・項目軸の文字の塊）。
 * **基準値のファイルが無い間は describe ごと skip する**（スイートは緑のまま）。基準値が来たら、場面ごと（TD-110-1〜6）に
 * `openXlsx` → `readChart` → `layoutChart` の結果を COM の値と比べる（検査の部品は 2 回目の `xlsx-chart-ref2-checks.ts`・取り出しは `xlsx-chart-ref-scene.ts`）。
 * - 単位・許容・合わせられない組の扱い（{@link KNOWN}・`lim`・`same`・`listed`・`mixed`）は 2 回目（`xlsx-chart-ref2.test.ts`）と同じ。
 *   基準値が来て差が出たら、直せる差は規則を直し（ISSUE-110 の「照合」工程）、直せない差は理由・上限つきで KNOWN に足す。
 * - 場面：c1sw＝横棒×グラフの幅の掃引（C1 の KNOWN 3・item 7）／tick＝数値軸の線 none×補助の目盛の印（KNOWN 1・item 8。COM の MinorUnit と PNG の印）／
 *   c7w＝項目軸の文字×グラフの幅（ISSUE-107）／c4dark＝`c:style` 41〜48（ISSUE-109）／round＝`c:roundedCorners`（ISSUE-109）。
 * - 開発用：`REF110_DUMP=<ファイル> npx vitest run …` で差の一覧と、c7w（＋ PNG があれば c7sw）の項目軸の文字の表・tick の PNG の印の表・round の PNG の角を書き出す。
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import type { PngChartResult } from '../../fixtures/excel/make-ref-charts-png.mjs';
import { openXlsx } from '../../src/core/index';
import type { Theme } from '../../src/core/theme';
import type { ChartModel } from '../../src/core/types';
import { layoutChart } from '../../src/render/chart-layout';
import {
  axisCheck,
  minorCheck,
  plotCheck,
  runGroup,
  styleCheck,
  type Case,
  type CheckFn,
  type Group,
} from './chart-ref2-checks';
import {
  fixtures,
  judge,
  K,
  lim,
  listed,
  near,
  segments,
  type ComChart,
  type Diff,
  type Known,
  type TextItem,
} from './chart-ref-scene';

interface Expected {
  charts: Record<string, ComChart>;
}

const HAS_REF =
  existsSync(fixtures('ref-charts3.expected.json')) && existsSync(fixtures('ref-charts3.xlsx'));
const HAS_PNG = HAS_REF && existsSync(fixtures('ref-charts3-png.json'));

// 合わせられない組の理由（種類ごとに 1 つ。下の KNOWN が使う）
const WHY_c1four =
  'C1：Excel の刻みが 4×10ⁿ（{1,2,5}×10ⁿ の外）で、最大が刻みの倍数にならない（0〜1000000 を 400000 刻み・0〜10000000 を 4000000 刻み）。3 回目の幅の掃引では、目盛の数の上限が 3〜4 になる狭い幅だけで、幅が広がるにつれ 5×→4×→2× と進む。{1,2,4,5} に広げると縦軸・他の横棒の刻みが崩れる（[4,19] の 5 刻みなど）ので規則は未確定。2 回目の c1mil_barClustered_mil7_small と同じ型。上限は今の値';
const WHY_c7est =
  'C7（ISSUE-107）：長い項目名（12〜20 文字）で、形は Excel と同じ（45°）だが、推定の文字幅（ASCII 0.55em）が游ゴシック 9pt の実際より広く（20 文字で推定 99pt・Excel 84pt）、回した帯が最大 9pt 厚くプロットの左が右へ寄る。文字幅の推定（§決定 8）は別 ISSUE。上限は今の値';
const WHY_c7wrap =
  'C7（ISSUE-107）：空白入りの項目名で、折り返し ↔ 45° の境目（語が項目の間隔に入るか）や行数が Excel と違う（推定の語の幅が Excel より約 15% 広い）。文字幅の推定は別 ISSUE。上限は今の値';
const WHY_c7reg =
  'C7（ISSUE-107）：Excel は水平（1 行の幅 50.25pt ≤ 間隔 56.5pt）だが、推定の幅（12 文字で 59.4pt）が間隔を超え、こちらは 45°（空白入りは 2 行に折り返し）にする。ISSUE-107 で回転・折り返しを入れて新しく出た差。文字幅の推定は別 ISSUE。上限は今の値';
const WHY_c4dark =
  'C4（ISSUE-109）：c:style 42 の円の凡例の文字。COM は #595959 9pt を返し、同じ XML の 45 は白 10pt になって規則が読めない（Q3＝今回は入れない。こちらは表どおり白 10pt）。ISSUE-109 で他の c4dark 11 個（面・文字・プロット・タイトル 18pt・面の枠線）は解消した。メッセージは今の値';

/**
 * 合わせられない組（`<種類>:<グラフ名>` → 理由と今の値）。数で測れる差は `lim`、測れない差は `same`・`listed`、両方あれば `mixed`
 * （2 回目と同じ。直った組は一覧から消す＝残っていると stale で失敗する）。
 * C1 の刻みは規則を直して 94 組が合い、{1,2,5}×10ⁿ の外（4×10ⁿ）の 2 組だけ残る。補助の目盛の印（tick）は差が無い
 * （数値軸の線が none の時、Excel は主・補助とも印を描かない。場面も描かない）。c7w は ISSUE-107 の回転・折り返しで 23→13 件（残りは文字幅の推定が原因）、c4dark は ISSUE-109 で 12→1 件（42 の円の凡例だけ残る）、round は 3→0 件。
 */
const KNOWN: Record<string, Known> = {
  // ---- c1sw ----
  'c1sw:c1sw_barClustered_mil7_200': lim(50, WHY_c1four),
  'c1sw:c1sw_barClustered_mil8_240': lim(50, WHY_c1four),
  // ---- c7w ----
  'c7w:c7w_col_12_240': lim(3.14, WHY_c7est),
  'c7w:c7w_col_12_600': lim(17.67, WHY_c7reg),
  'c7w:c7w_col_12s_240': lim(7.61, WHY_c7wrap),
  'c7w:c7w_col_12s_360': lim(5.13, WHY_c7wrap),
  'c7w:c7w_col_12s_600': lim(5.12, WHY_c7reg),
  'c7w:c7w_col_20_240': lim(5.25, WHY_c7est),
  'c7w:c7w_col_20_300': lim(4.56, WHY_c7est),
  'c7w:c7w_col_20_360': lim(4.56, WHY_c7est),
  'c7w:c7w_col_20_480': lim(4.56, WHY_c7est),
  'c7w:c7w_col_20_600': lim(4.56, WHY_c7est),
  'c7w:c7w_col_20_720': lim(4.56, WHY_c7est),
  'c7w:c7w_col_20s_240': lim(5.32, WHY_c7est),
  'c7w:c7w_col_20s_300': lim(11.24, WHY_c7wrap),
  // ---- c4dark ----
  'c4dark:c4dark_42_pie_std': listed(
    ['凡例の文字 got[13.33px #ffffff] want[12.00px #595959]'],
    WHY_c4dark,
  ),
};

/** COM の ChartArea.RoundedCorners（ps1 が足した項目）を読む */
function roundedOf(c: Case): boolean | undefined {
  const v = (c.exp.ChartArea as { RoundedCorners?: boolean | number }).RoundedCorners;
  return v === undefined ? undefined : v === true || v === -1;
}

/**
 * 面の角丸（`c:roundedCorners`・ISSUE-109 の⑥）：COM の ChartArea.RoundedCorners が真なのに、場面の面が直角（場面の面の矩形に半径が無い）なら差とする。
 * 場面の面の矩形の半径 `radius`（ISSUE-109）で比べる。真なら 11px（8.25pt）、偽（`c:roundedCorners val="0"`）なら半径なし。
 */
const roundCheck: CheckFn = (c, add) => {
  const want = roundedOf(c);
  if (want === undefined) {
    add('COM の RoundedCorners が読めていない');
    return;
  }
  const ca = c.scene.items[0];
  const radius = ca?.kind === 'rect' ? (ca.radius ?? 0) : 0;
  const gotRound = radius > 0;
  if (want !== gotRound)
    add(`面の角 got[${gotRound ? '角丸' : '直角'}] want[RoundedCorners=${want}]`);
  // 半径は一定（PNG の角の 11px＝8.25pt。ref-charts3-png.json の corner）。最小の場面（216×144pt）でも 幅÷2・高さ÷2 を超えない
  else if (want && !near(radius, 11, 0.01))
    add(`面の角の半径 got[${radius.toFixed(2)}px] want[11px]`);
};

/** 場面が補助の目盛の印（補助の目盛線でなく、長さ minorTickLengthPt＝3pt の短い線分だけの経路）を描いているか。補助の目盛線の有無によらない */
function sceneDrawsMinorTicks(c: Case): boolean {
  return c.scene.items.some((it) => {
    if (it.kind !== 'path' || it.fill !== null || it.stroke === null) return false;
    const segs = segments(it.d);
    return (
      segs !== null &&
      segs.length > 0 &&
      segs.every((q) => near(Math.hypot(q[2] - q[0], q[3] - q[1]), 3 * K, 0.2))
    );
  });
}

/**
 * 補助の目盛の印（KNOWN 1・item 8）：PNG の数値軸の印（補助の目盛の位置に出た短い線）の有無を、場面が補助の印を描くかと比べる。
 * 数値軸の線が none（Excel 2013 以降の既定）の時、補助の印だけの場面（補助の目盛線なし）で Excel が印を描くかを PNG で決める。
 * 補助の目盛線がある場面は、内側の印が線に隠れるので外側の印だけを見る。
 */
function tickPngCheck(png: Record<string, PngChartResult>): CheckFn {
  return (c, add) => {
    const p = png[c.name]?.valTicks;
    if (p === undefined) return;
    const grid = c.model.valAxis?.minorGridlines !== undefined;
    const pngMinor = p.out.minor > 0 || (!grid && p.in.minor > 0);
    const sceneMinor = sceneDrawsMinorTicks(c);
    if (pngMinor !== sceneMinor) {
      add(
        `補助の目盛の印 got[場面 ${sceneMinor ? '有' : '無'}] want[PNG ${pngMinor ? '有' : '無'}：外 ${p.out.minor} 本（長さ ${p.out.minorLen}）・内 ${p.in.minor} 本（長さ ${p.in.minorLen}）]`,
      );
    }
  };
}

interface Loaded {
  model: ChartModel;
  theme: Theme | undefined;
}

/** ブックの全シートのグラフを、図の名前（`ChartObject.Name`）→ 読み取り結果にする */
async function readAll(file: string): Promise<Map<string, Loaded>> {
  const book = await openXlsx(readFileSync(fixtures(file)));
  const out = new Map<string, Loaded>();
  for (let i = 1; i <= book.sheets.length; i++) {
    const { drawings } = await book.openDrawings(i);
    for (const d of drawings) {
      if (d.chartPath === undefined) continue;
      const r = await book.readChart(i, d.chartPath);
      if (r.ok) out.set(d.name, { model: r.chart, theme: r.theme });
    }
  }
  return out;
}

const cases: Case[] = [];
let expected: Expected = { charts: {} };
let png: Record<string, PngChartResult> = {};
let all: Diff[] = [];

/** 項目軸の文字（`C01…` で始まる名前）の回転と表示数（c7w の表用）。無ければ空 */
function categoryRotations(c: Case): { rotations: number[]; shown: number } {
  const t = c.scene.items.filter((i): i is TextItem => i.kind === 'text' && /^C\d\d/.test(i.text));
  return { rotations: [...new Set(t.map((x) => x.rotate))], shown: t.length };
}

/** c7w の表：COM のプロットの内側の高さ・項目軸の帯の厚さ・TickLabels の向きと、場面の回転・表示数と、PNG の項目軸の文字の塊（あれば） */
function c7Table(): Record<string, unknown>[] {
  return cases
    .filter((c) => c.name.startsWith('c7w_'))
    .map((c) => {
      const w = c.exp.CategoryAxis;
      const { rotations, shown } = categoryRotations(c);
      const l = png[c.name]?.labels;
      return {
        name: c.name,
        comPlotH: c.exp.PlotArea.Inside.Height,
        comAxisBandH: w?.Pos?.Height,
        comOrientation: w?.TickLabels.Orientation,
        comSpacing: w?.TickLabelSpacing,
        sceneRotations: rotations,
        sceneLabelsShown: shown,
        pngClusters: l?.count,
        pngMedianAngle: l?.medianAngle,
        pngMedianW: l?.medianW,
        pngMedianH: l?.medianH,
      };
    });
}

/** PNG の項目軸の文字の塊の表（ref-charts2.xlsx の c7sw_* も、3 回目の PNG に入っていれば） */
function pngLabelTable(): Record<string, unknown>[] {
  return Object.entries(png)
    .filter(([n, p]) => /^c7(w|sw)_/.test(n) && p.labels !== undefined)
    .map(([n, p]) => ({
      name: n,
      clusters: p.labels!.count,
      medianAngle: p.labels!.medianAngle,
      medianW: p.labels!.medianW,
      medianH: p.labels!.medianH,
    }));
}

/** tick の PNG の印の表と、round の PNG の角の表 */
function pngTickRoundTable(): Record<string, unknown>[] {
  return Object.entries(png)
    .filter(([n]) => n.startsWith('tick_') || n.startsWith('round_'))
    .map(([n, p]) => ({ name: n, ticks: p.valTicks, corner: p.corner }));
}

const GROUPS110: Group[] = [
  {
    cat: 'c1sw',
    label: 'TD-110-1 横棒×グラフの幅の掃引：刻みの規則（C1 の KNOWN 3）',
    charts: ['c1sw_'],
    checks: [axisCheck, plotCheck],
  },
  {
    cat: 'tick',
    label:
      'TD-110-2 数値軸の線 none×補助の目盛の印：COM の MinorUnit と場面の補助の刻み（KNOWN 1）',
    charts: ['tick_'],
    checks: [minorCheck, plotCheck],
  },
  {
    cat: 'c7w',
    label: 'TD-110-4 項目軸の文字×グラフの幅（ISSUE-107・C7）',
    charts: ['c7w_'],
    checks: [plotCheck],
  },
  {
    cat: 'c4dark',
    label: 'TD-110-5 c:style 41〜48 の面・文字（ISSUE-109・C4）',
    charts: ['c4dark_'],
    checks: [styleCheck, plotCheck],
  },
  {
    cat: 'round',
    label: 'TD-110-6 面の角丸（ISSUE-109）',
    charts: ['round_'],
    checks: [roundCheck],
  },
];

describe.skipIf(!HAS_REF)('ISSUE-110 基準値の読み込み（ref-charts3）', () => {
  beforeAll(async () => {
    expected = JSON.parse(readFileSync(fixtures('ref-charts3.expected.json'), 'utf8')) as Expected;
    if (HAS_PNG) {
      png = (
        JSON.parse(readFileSync(fixtures('ref-charts3-png.json'), 'utf8')) as {
          charts: Record<string, PngChartResult>;
        }
      ).charts;
    }
    const read = await readAll('ref-charts3.xlsx');
    for (const [name, exp] of Object.entries(expected.charts)) {
      const r = read.get(name);
      if (r === undefined) continue;
      const scene = layoutChart(
        r.model,
        exp.ChartObject.Width * K,
        exp.ChartObject.Height * K,
        r.theme,
      );
      if (scene === null) continue;
      cases.push({ name, group: 'charts', exp, model: r.model, scene });
    }
    const tickPng: Group = {
      cat: 'tickpng',
      label: 'TD-110-3',
      charts: ['tick_'],
      checks: [tickPngCheck(png)],
    };
    all = [
      ...GROUPS110.flatMap((g) => runGroup(g, cases)),
      ...(HAS_PNG ? runGroup(tickPng, cases) : []),
    ];
    if (process.env.REF110_DUMP) {
      writeFileSync(
        process.env.REF110_DUMP,
        JSON.stringify(
          {
            diffs: all,
            c7w: c7Table(),
            pngLabels: pngLabelTable(),
            pngTickRound: pngTickRoundTable(),
          },
          null,
          1,
        ),
      );
    }
  });

  it('TD-110-0 ref-charts3.xlsx の 166 個が、すべて読めて場面になる', () => {
    expect(cases).toHaveLength(166);
    // 基準値の全部が場面になっている（読めなかったグラフがあれば名前を挙げる）
    const names = new Set(cases.map((c) => c.name));
    expect(Object.keys(expected.charts).filter((n) => !names.has(n))).toEqual([]);
  });

  for (const g of GROUPS110) {
    it(g.label, () => {
      const r = judge(all, g.cat, KNOWN);
      expect(r.unexpected).toEqual([]);
      // 外した組が今より悪化したら失敗（上限・メッセージを超える）
      expect(r.worse).toEqual([]);
      // 外した組が直っていたら、一覧から消す
      expect(r.stale).toEqual([]);
    });
  }

  // PNG（ref-charts3-png.json）が来ている時だけ：数値軸の補助の目盛の印（KNOWN 1）
  it.skipIf(!HAS_PNG)(
    'TD-110-3 PNG の数値軸の補助の目盛の印と、場面の補助の印が合う（KNOWN 1）',
    () => {
      const r = judge(all, 'tickpng', KNOWN);
      expect(r.unexpected).toEqual([]);
      expect(r.worse).toEqual([]);
      expect(r.stale).toEqual([]);
    },
  );

  it.skipIf(!HAS_PNG)('PNG の読み取りが、軸のあるグラフの全部で線を 1 本以上取っている', () => {
    const empty = Object.entries(png)
      .filter(([n, p]) => p.hasAxes && p.lines.length === 0 && !/^(c7sw|c4dark)_/.test(n))
      .map(([n]) => n);
    expect(empty).toEqual([]);
  });
});
