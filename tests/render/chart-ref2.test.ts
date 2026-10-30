/// <reference types="node" />
/**
 * Excel 16 の COM が返した値との、グラフの 2 回目の照合（ISSUE-105・ADR-0031 §決定 16「次の COM で確かめる事項」）。
 * 基準値は `__fixtures__/excel/make-ref-charts2.ps1` の出力（ユーザー承認後に実行して取り込む）：
 * `ref-charts2.xlsx`（127 個）＋`ref-charts2.expected.json`（`charts`＝その COM の値、`defaults`＝`ref-charts-default.xlsx` の 31 個）＋`ref-charts-png.json`（PNG の線）。
 * **基準値のファイルが無い間は describe ごと skip する**（スイートは緑のまま）。基準値が来たら、場面ごと（TD-105-1〜12）に
 * `openXlsx` → `readChart` → `layoutChart` の結果を COM の値と比べる（検査の部品は `xlsx-chart-ref2-checks.ts`・取り出しは `xlsx-chart-ref-scene.ts`）。
 * - 単位・許容は 1 回目（`xlsx-chart-ref.test.ts`）と同じ：COM は pt（× 96/72 で px）・位置と大きさはグラフの面の 3%・軸は 1e-9。
 * - 合わせられない組は {@link KNOWN} に理由付きで挙げて期待から外す（残っていることも検査する。1 回目と同じ作り）。
 *   基準値で直せた差（C2 の範囲・C3 の円の凡例・C11 の凡例の並び・C12 の目盛の印・C13 のマーカー・C4 の線の色、
 *   ISSUE-108 で C3 の縦の項目数の上限・C5 の右上の凡例・C6 の横並びの格子）は規則を直し、
 *   直せない差（C1 の一部・C4 の c14:style 付き／暗い背景・C7）は理由・上限つきで足した。
 * - 開発用：`REF105_DUMP=<ファイル> npx vitest run …` で差の一覧と、C7 の掃引の表（COM のプロットの高さ・項目軸の帯・回転と、場面の回転）を書き出す。
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import type { PngChartResult } from '../../fixtures/excel/make-ref-charts-png.mjs';
import { openXlsx } from '../../src/core/index';
import type { Theme } from '../../src/core/theme';
import type { ChartModel } from '../../src/core/types';
import { layoutChart } from '../../src/render/chart-layout';
import { GROUPS, pngCheck, runGroup, type Case, type Group } from './chart-ref2-checks';
import {
  fixtures,
  judge,
  K,
  lim,
  listed,
  type ComChart,
  type Diff,
  type Known,
  type TextItem,
} from './chart-ref-scene';

interface Expected {
  charts: Record<string, ComChart>;
  defaults?: Record<string, ComChart>;
}

const HAS_REF =
  existsSync(fixtures('ref-charts2.expected.json')) && existsSync(fixtures('ref-charts2.xlsx'));
const HAS_PNG = HAS_REF && existsSync(fixtures('ref-charts-png.json'));

// 合わせられない組の理由（種類ごとに 1 つ。下の KNOWN が使う）
const WHY_c1mil =
  'C1：7 桁の目盛の小さい横棒（プロットの幅 157pt）で、Excel の刻みは 400000（{1,2,5}×10ⁿ の外。こちらは 200000）。他の幅（標準・大）と 8 桁（小・標準・大）は合う。3 回目の幅の掃引（c1sw_*）で 4×10ⁿ は目盛の数の上限が 3〜4 の狭い幅だけに出ると分かったが、{1,2,4,5} に広げると他の刻みが崩れるので規則は未確定。上限は今の値';
const WHY_c2col =
  'C2：全部が負の 100% 積み上げ（縦棒・小さい面）で、プロットの下端がこちらは 7.7px 低い（軸の範囲・刻みは合う）。小さい面だけの差。原因は未特定で暫定のまま';
const WHY_c7est =
  'C7（ISSUE-107）：長い項目名（12〜20 文字）で、形は Excel と同じ（45°）だが、推定の文字幅（ASCII 0.55em）が游ゴシック 9pt の実際より広く（20 文字で推定 99pt・Excel 84pt）、回した帯が最大 9pt 厚くプロットの左が右へ寄る。文字幅の推定（§決定 8）は別 ISSUE。上限は今の値';
const WHY_c7wrap =
  'C7（ISSUE-107）：空白入りの項目名で、折り返し ↔ 45° の境目（語が項目の間隔に入るか）や行数が Excel と違う（推定の語の幅が Excel より約 15% 広い）。文字幅の推定は別 ISSUE。上限は今の値';
const WHY_c7n90sp =
  'C7（ISSUE-107）：30 項目×空白入りの 12 文字で Excel は 90° に回し、帯は 40pt・プロットの左の張り出しは無い（20pt のまま）。こちらは 90° の帯を 45° と同じ式（約 49pt）で見積もり、左も張り出す。90° の帯・張り出しは基準値が 6 場面だけで暫定。上限は今の値';
const WHY_c7n904 =
  'C7（ISSUE-107）：30 項目×4 文字で Excel は 90° に回すが、帯は 18pt（水平と同じ）。こちらは 90° の帯を 45° と同じ式（約 28pt）で見積もる。90° の帯は暫定。上限は今の値';
const WHY_c7bar =
  'C7：長い項目名の横棒（縦の軸）で、Excel も水平 1 行（折り返さない・回さない）。差は文字幅の推定（20 文字で推定 99pt・Excel 84pt）で項目名の帯が 約 15pt 広く、プロットの左が 148px（Excel は 132px）。ISSUE-107 で横棒は回さないと決めた（挙動は変えない）。文字幅の推定は別 ISSUE。上限は今の値';
const WHY_c4pie =
  'C4：円の凡例の文字。c4style_3_pie（c:style 3・c14:style 付き）の COM は #595959 9pt を返す（c:style 42 の円 c4dark_42_pie も同じで、同じ XML の 45 は白 10pt になり規則が読めない）。こちらは黒 10pt。ISSUE-109 の Q3＝今回は入れない。メッセージは今の値';
const WHY_png42 =
  'C4（ISSUE-109）：c:style 42（暗い背景）の項目軸の線・数値軸の線は、PNG の背景（プロットの #3F3F3F）に線が重なって拾えない（2 回目・3 回目とも）。COM の値は c14:style 無し＝黒・有り＝#898989 で、こちらはその色の線を描く（確かめられないだけ）。メッセージは今の値';

/**
 * 合わせられない組（`<種類>:<グラフ名>` → 理由と今の値）。数で測れる差は `lim`（上限は今の値を小数 2 桁へ切り上げ）、
 * 測れない差は `same`（1 件）・`listed`（複数件）、両方あれば `mixed`。直った組は一覧から消す（残っていると stale で失敗する）。
 */
const KNOWN: Record<string, Known> = {
  // ---- c1 ----
  'c1:c1mil_barClustered_mil7_small': lim(50, WHY_c1mil),
  // ---- c2 ----
  'c2:c2allneg_colStacked100_small': lim(4.03, WHY_c2col),
  // ---- c4 ----
  'c4:c4style_3_pie_std': listed(
    ['凡例の文字 got[13.33px #000000] want[12.00px #595959]'],
    WHY_c4pie,
  ),
  // ---- c7 ----
  'c7:c7sw_col_5_20': lim(4.56, WHY_c7est),
  'c7:c7sw_col_10_12s': lim(5.13, WHY_c7wrap),
  'c7:c7sw_col_10_20': lim(4.56, WHY_c7est),
  'c7:c7sw_col_15_12s': lim(7.5, WHY_c7wrap),
  'c7:c7sw_col_15_20': lim(4.56, WHY_c7est),
  'c7:c7sw_col_15_20s': lim(4.64, WHY_c7est),
  'c7:c7sw_col_20_12s': lim(3.66, WHY_c7est),
  'c7:c7sw_col_20_20': lim(4.56, WHY_c7est),
  'c7:c7sw_col_20_20s': lim(4.64, WHY_c7est),
  'c7:c7sw_col_30_4': lim(4.7, WHY_c7n904),
  'c7:c7sw_col_30_12s': lim(7.5, WHY_c7n90sp),
  'c7:c7sw_col_30_20': lim(4.56, WHY_c7est),
  'c7:c7sw_col_30_20s': lim(4.64, WHY_c7est),
  'c7:c7sw_bar_15_20': lim(3.51, WHY_c7bar),
  'c7:c7sw_bar_15_20s': lim(3.57, WHY_c7bar),
  'c7:c7sw_bar_30_20': lim(3.51, WHY_c7bar),
  'c7:c7sw_bar_30_20s': lim(3.57, WHY_c7bar),
  // ---- png ----
  'png:c4style_42_colClustered_std': listed(
    ['項目軸の線：PNG に線が無い（場面に有る）', '数値軸の線：PNG に線が無い（場面に有る）'],
    WHY_png42,
  ),
  'png:def2_style_42_col': listed(
    ['項目軸の線：PNG に線が無い（場面に有る）', '数値軸の線：PNG に線が無い（場面に有る）'],
    WHY_png42,
  ),
};

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

/** 項目軸の文字（`C01…` で始まる名前）の回転（C7 の表用）。無ければ空 */
function categoryRotations(c: Case): { rotations: number[]; shown: number } {
  const t = c.scene.items.filter((i): i is TextItem => i.kind === 'text' && /^C\d\d/.test(i.text));
  return { rotations: [...new Set(t.map((x) => x.rotate))], shown: t.length };
}

/** C7 の掃引の表：COM のプロットの内側の高さ・項目軸の帯の厚さ・TickLabels の値と、場面の回転・表示数 */
function c7Table(): Record<string, unknown>[] {
  return cases
    .filter((c) => c.group === 'charts' && c.name.startsWith('c7sw_'))
    .map((c) => {
      const w = c.exp.CategoryAxis;
      const { rotations, shown } = categoryRotations(c);
      return {
        name: c.name,
        comPlotH: c.exp.PlotArea.Inside.Height,
        comAxisBandH: w?.Pos?.Height,
        comOrientation: w?.TickLabels.Orientation,
        comSpacing: w?.TickLabelSpacing,
        sceneRotations: rotations,
        sceneLabelsShown: shown,
      };
    });
}

describe.skipIf(!HAS_REF)('ISSUE-105 基準値の読み込み（ref-charts2）', () => {
  beforeAll(async () => {
    expected = JSON.parse(readFileSync(fixtures('ref-charts2.expected.json'), 'utf8')) as Expected;
    if (HAS_PNG) {
      png = (
        JSON.parse(readFileSync(fixtures('ref-charts-png.json'), 'utf8')) as {
          charts: Record<string, PngChartResult>;
        }
      ).charts;
    }
    const sets: [Case['group'], string, Record<string, ComChart>][] = [
      ['charts', 'ref-charts2.xlsx', expected.charts],
      ['defaults', 'ref-charts-default.xlsx', expected.defaults ?? {}],
    ];
    for (const [group, file, exps] of sets) {
      if (Object.keys(exps).length === 0) continue;
      const read = await readAll(file);
      for (const [name, exp] of Object.entries(exps)) {
        const r = read.get(name);
        if (r === undefined) continue;
        const scene = layoutChart(
          r.model,
          exp.ChartObject.Width * K,
          exp.ChartObject.Height * K,
          r.theme,
        );
        if (scene === null) continue;
        cases.push({ name, group, exp, model: r.model, scene });
      }
    }
    const pngGroup: Group = {
      cat: 'png',
      label: 'TD-105-12',
      charts: [''],
      defaults: [''],
      checks: [pngCheck(png)],
    };
    all = [
      ...GROUPS.flatMap((g) => runGroup(g, cases)),
      ...(HAS_PNG ? runGroup(pngGroup, cases) : []),
    ];
    if (process.env.REF105_DUMP) {
      writeFileSync(
        process.env.REF105_DUMP,
        JSON.stringify({ diffs: all, c7: c7Table() }, null, 1),
      );
    }
  });

  it('TD-105-0 ref-charts2.xlsx の 127 個と ref-charts-default.xlsx の 31 個が、すべて読めて場面になる', () => {
    expect(cases.filter((c) => c.group === 'charts')).toHaveLength(127);
    expect(cases.filter((c) => c.group === 'defaults')).toHaveLength(31);
    // 基準値の全部が場面になっている（読めなかったグラフがあれば名前を挙げる）
    const names = new Set(cases.map((c) => `${c.group}:${c.name}`));
    const missing = [
      ...Object.keys(expected.charts).map((n) => `charts:${n}`),
      ...Object.keys(expected.defaults ?? {}).map((n) => `defaults:${n}`),
    ].filter((n) => !names.has(n));
    expect(missing).toEqual([]);
  });

  for (const g of GROUPS) {
    it(g.label, () => {
      const r = judge(all, g.cat, KNOWN);
      expect(r.unexpected).toEqual([]);
      // 外した組が今より悪化したら失敗（上限・メッセージを超える）
      expect(r.worse).toEqual([]);
      // 外した組が直っていたら、一覧から消す
      expect(r.stale).toEqual([]);
    });
  }

  // PNG（ref-charts-png.json）が来ている時だけ：目盛線・軸線の色（C4）
  it.skipIf(!HAS_PNG)(
    'TD-105-12 PNG の線の ink と、場面の線の期待の ink が合う（主の目盛線・項目軸の線・数値軸の線・C4）',
    () => {
      const r = judge(all, 'png', KNOWN);
      expect(r.unexpected).toEqual([]);
      expect(r.worse).toEqual([]);
      expect(r.stale).toEqual([]);
    },
  );

  it.skipIf(!HAS_PNG)('PNG の読み取りが、軸のあるグラフの全部で線を 1 本以上取っている', () => {
    const empty = Object.entries(png)
      .filter(([, p]) => p.hasAxes && p.lines.length === 0)
      .map(([n]) => n);
    expect(empty).toEqual([]);
  });
});
