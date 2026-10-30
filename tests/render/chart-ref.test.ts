/// <reference types="node" />
/**
 * Excel 16 の COM が返した値との、グラフの照合（ISSUE-104・ADR-0031 §決定 12・13・C1〜C10）。
 * `__fixtures__/excel/ref-charts.xlsx`（113 個：G1 の 10 種×データ×大きさの 80＋extra 33）と `ref-charts.expected.json`（`charts`＝COM の値・pt）、
 * `ref-charts-default.xlsx`（`c:spPr`・`c:txPr` の無い合成のグラフ 19 個）と `defaults`（同・C4）を、
 * `openXlsx` → `readChart` → `layoutChart` の結果（場面）と比べる。
 * - 単位：COM は pt。場面は 96dpi の px なので pt × 96/72 で比べる（画面の倍率は COM の pt に影響しない）。グラフの面の大きさは COM の ChartObject の幅・高さ。
 * - 場面は部品の名前を持たないので、数値の目盛の文字（軸の最小・最大・刻み）・主目盛線の外接矩形（プロットの内側）・
 *   タイトルの文字の中心・凡例の項目（記号と文字）の外接矩形・系列の塗り／線／マーカー／扇の色と太さから、基準値と同じ量を取り出す。
 * - 位置・大きさはグラフの面の 3%（manualLayout のグラフは 1%）以内で一致とする（ADR-0031 §決定 12）。
 * - 合わせられない組は {@link KNOWN} に理由付きで挙げて期待から外す（残っていることも検査する）。
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { openXlsx } from '../../src/core/index';
import type { Theme } from '../../src/core/theme';
import type { ChartModel } from '../../src/core/types';
import {
  estimateTextWidth,
  layoutChart,
  PX_PER_PT,
  type ChartScene,
} from '../../src/render/chart-layout';

// ---- 基準値の型（使う所だけ） ----
interface Box {
  Left: number;
  Top: number;
  Width: number;
  Height: number;
}
interface ComSeries {
  Name: string;
  Fill: { RGB: string | number };
  Line: { RGB: string | number; Weight: number };
  MarkerStyle: number;
  MarkerSize: number;
  Points?: { Fill: { RGB: string | number }; Line: { RGB: string | number } }[];
}
interface ComChart {
  ChartObject: { Width: number; Height: number };
  ChartArea: { Line: { Weight: number; RGB: string | number } };
  PlotArea: { Inside: Box };
  HasTitle: boolean;
  Title?: {
    Text: string;
    Pos: Box;
    Font: { Size: number; Bold: number | boolean; RGB: string | number };
  };
  HasLegend: boolean;
  Legend?: { Pos: Box; Font: { Size: number; RGB: string | number }; Entries: { Pos: Box }[] };
  ValueAxis?: {
    MinimumScale: number;
    MaximumScale: number;
    MajorUnit: number;
    TickLabels: { Font: { Size: number; RGB: string | number } };
  };
  Group: { VaryByCategories: boolean };
  Series: ComSeries[];
}

/** 1 つのグラフの読み込み結果 */
interface Case {
  name: string;
  group: 'charts' | 'defaults';
  exp: ComChart;
  model: ChartModel;
  theme: Theme | undefined;
  scene: ChartScene;
}

/** 照合の種類（TD-104-n の分け方） */
type Category = 'axis' | 'color' | 'marker' | 'width' | 'plot' | 'legend' | 'title' | 'style';

/** 照合の差 1 件 */
interface Diff {
  name: string;
  cat: Category;
  msg: string;
  /** ずれの大きさ（%。位置・大きさは面に対する割合・軸は範囲に対する割合）。数で測れない差（色・太さ・文字）は無し */
  metric?: number;
}

type TextItem = Extract<ChartScene['items'][number], { kind: 'text' }>;
type PathItem = Extract<ChartScene['items'][number], { kind: 'path' }>;

const K = PX_PER_PT;
/**
 * COM の座標の原点のずれ（pt）：COM の座標は、グラフの面の左上から 4pt 内側が原点（Left・Top に足して面の座標にする）。
 * 根拠（Excel 16 の COM の値。凡例・タイトルは面の縁から 6pt に置かれる、とみて 4pt のずれで揃うかを見た）：
 * - 凡例を左に置くと Left=2（2+4=6pt）。
 * - 右に置くと、凡例の右端から面の右端までが COM の値で 10（原点が 4pt 内側なので、実際の余白は 10−4=6pt）。
 * - タイトルの Top=2（2+4=6pt）。
 * - 下に置いた凡例の、下端から面の下端までが COM の値で 10（同じく 10−4=6pt）。
 * 4 つとも 6pt にそろうのは 4pt の時だけ。タイトルの中心が面の中心より 4pt 左になること（216pt 幅で 104・360pt 幅で 176）とも合う。
 */
const COM_ORIGIN_PT = 4;
const fixtures = (n: string): URL => new URL(`../../fixtures/excel/${n}`, import.meta.url);

// ---- 場面から基準値と同じ量を取り出す ----

interface Cmd {
  c: string;
  n: number[];
}

/** 経路の `d`（数値と `M L C A Z` だけ）を命令ごとに分ける */
function cmds(d: string): Cmd[] {
  const out: Cmd[] = [];
  for (const m of d.matchAll(/([MLCAZ])([^MLCAZ]*)/g)) {
    out.push({ c: m[1]!, n: (m[2]!.match(/-?\d+(?:\.\d+)?(?:e-?\d+)?/gi) ?? []).map(Number) });
  }
  return out;
}

interface PxBox {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** 経路の点（M・L の座標・A の終点）の外接矩形 */
function pathBox(d: string): PxBox | null {
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  for (const { c, n } of cmds(d)) {
    if (c === 'Z' || n.length < 2) continue;
    const x = n[n.length - 2]!;
    const y = n[n.length - 1]!;
    x0 = Math.min(x0, x);
    x1 = Math.max(x1, x);
    y0 = Math.min(y0, y);
    y1 = Math.max(y1, y);
  }
  return x0 <= x1 ? { x: x0, y: y0, w: x1 - x0, h: y1 - y0 } : null;
}

/**
 * 数値軸の目盛の文字（場面の文字のうち、数値軸の文字の置き方に当たるものだけ）。
 * 項目名・凡例・タイトルが年などの数字でも拾わないよう、向きごとの anchor・baseline・回転・位置で絞る：
 * - 縦の数値軸（縦棒・折れ線）：anchor end・baseline middle・回転なしで、プロットの左・縦はプロットの範囲内
 *   （凡例の文字は anchor start、項目名は anchor middle か回転付き）。
 * - 横の数値軸（横棒）：anchor middle・baseline hanging（下）で、プロットの下
 *   （項目名は anchor end、タイトルは baseline middle）。
 */
function valueLabelTexts(scene: ChartScene, model: ChartModel): TextItem[] {
  const plot = gridBox(scene);
  if (plot === null) return [];
  const horizontal = model.barDir === 'bar';
  const eps = 1;
  return scene.items.filter((it): it is TextItem => {
    if (it.kind !== 'text' || it.rotate !== 0) return false;
    if (!/^-?\d[\d.]*(?:E[+-]?\d+)?%?$/.test(it.text)) return false;
    if (horizontal) {
      return (
        it.anchor === 'middle' &&
        it.baseline === 'hanging' &&
        it.y >= plot.y + plot.h - eps &&
        it.x >= plot.x - plot.w &&
        it.x <= plot.x + plot.w * 2
      );
    }
    return (
      it.anchor === 'end' &&
      it.baseline === 'middle' &&
      it.x <= plot.x + eps &&
      it.y >= plot.y - eps &&
      it.y <= plot.y + plot.h + eps
    );
  });
}

/** 数値の目盛の文字（軸の最小・最大・刻みの元）。`%` は 100 で割る */
function numericLabels(scene: ChartScene, model: ChartModel): number[] {
  return valueLabelTexts(scene, model).map((it) =>
    it.text.endsWith('%') ? Number(it.text.slice(0, -1)) / 100 : Number(it.text),
  );
}

/** 軸の最小・最大・刻み（場面の目盛の文字から）。目盛が 2 つ未満なら null */
function axisOf(
  scene: ChartScene,
  model: ChartModel,
): { min: number; max: number; major: number } | null {
  const v = numericLabels(scene, model);
  if (v.length < 2) return null;
  return { min: v[0]!, max: v[v.length - 1]!, major: (v[v.length - 1]! - v[0]!) / (v.length - 1) };
}

/**
 * 目盛線の経路（塗り無し・butt・M/L だけ）の外接矩形＝プロットの内側。無ければ null。
 * 補助の目盛線は端（最小・最大）に線が無く小さい外接になる・目盛の印は細いので、面積が最大のものを取る。
 */
function gridBox(scene: ChartScene): PxBox | null {
  let best: PxBox | null = null;
  for (const it of scene.items) {
    if (it.kind !== 'path' || it.fill !== null || it.stroke === null) continue;
    if (it.stroke.cap !== 'butt' || /[ACZ]/.test(it.d)) continue;
    const b = pathBox(it.d);
    if (b !== null && (best === null || b.w * b.h > best.w * best.h)) best = b;
  }
  return best;
}

/** 円・ドーナツの外接の正方形（最大の外半径の扇の中心と半径。開始角 0 の前提：ドーナツは始点が 12 時の点） */
function pieBox(scene: ChartScene): PxBox | null {
  let best: { cx: number; cy: number; r: number } | null = null;
  for (const it of scene.items) {
    if (it.kind !== 'path' || it.fill === null) continue;
    const cs = cmds(it.d);
    const m0 = cs[0];
    const arcs = cs.filter((c) => c.c === 'A');
    if (m0 === undefined || m0.c !== 'M' || arcs.length === 0) continue;
    const r = Math.max(...arcs.map((c) => c.n[0]!));
    const isSector = cs[1]?.c === 'L';
    const cx = m0.n[0]!;
    const cy = isSector ? m0.n[1]! : m0.n[1]! + r;
    if (best === null || r > best.r) best = { cx, cy, r };
  }
  return best === null
    ? null
    : { x: best.cx - best.r, y: best.cy - best.r, w: 2 * best.r, h: 2 * best.r };
}

/** 凡例の文字（anchor start・baseline middle）と、その左の記号（四角・線）の外接矩形。無ければ null */
function legendBox(scene: ChartScene): PxBox | null {
  const texts = scene.items.filter(
    (i): i is TextItem => i.kind === 'text' && i.anchor === 'start' && i.baseline === 'middle',
  );
  if (texts.length === 0) return null;
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  for (const t of texts) {
    x0 = Math.min(x0, t.x);
    x1 = Math.max(x1, t.x + estimateTextWidth(t.text, t.size));
    y0 = Math.min(y0, t.y - t.size * 0.6);
    y1 = Math.max(y1, t.y + t.size * 0.6);
  }
  for (const it of scene.items) {
    let kx: number | undefined;
    let ky: number | undefined;
    if (it.kind === 'rect' && it.width < 20 && it.fill !== null) {
      kx = it.x;
      ky = it.y + it.height / 2;
    } else if (it.kind === 'line' && it.y1 === it.y2 && it.x2 - it.x1 < 40) {
      kx = it.x1;
      ky = it.y1;
    }
    if (kx !== undefined && ky !== undefined && texts.some((t) => Math.abs(t.y - ky!) < 0.5)) {
      x0 = Math.min(x0, kx);
    }
  }
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

/** COM の色（`RRGGBB`）を `#rrggbb` にする（数値＝色の定数はそのまま `?n`） */
const hexOf = (c: string | number | undefined): string =>
  typeof c === 'string' ? `#${c.toLowerCase()}` : `?${String(c)}`;

const near = (a: number, b: number, eps: number): boolean => Math.abs(a - b) <= eps;
/**
 * COM の要素の座標（pt）を、グラフの面の左上を原点とする px にする。COM の座標は面の左上から {@link COM_ORIGIN_PT} だけ内側が原点
 * （タイトルの中心が面の中心より 4pt 左になる：216pt 幅で 104・360pt 幅で 176）ため、足してから px にする。
 */
const px = (b: Box): PxBox => ({
  x: (b.Left + COM_ORIGIN_PT) * K,
  y: (b.Top + COM_ORIGIN_PT) * K,
  w: b.Width * K,
  h: b.Height * K,
});
const fmt = (b: PxBox): string =>
  `${b.x.toFixed(1)},${b.y.toFixed(1)} ${b.w.toFixed(1)}x${b.h.toFixed(1)}`;

/** 位置・大きさ（px）が許容（グラフの面に対する割合）内か。ずれを文字と大きさ（%）にして返す（合えば null） */
function boxDiff(
  label: string,
  got: PxBox,
  want: PxBox,
  scene: ChartScene,
  tol: number,
): { msg: string; metric: number } | null {
  const dx = Math.abs(got.x - want.x) / scene.width;
  const dy = Math.abs(got.y - want.y) / scene.height;
  const dw = Math.abs(got.w - want.w) / scene.width;
  const dh = Math.abs(got.h - want.h) / scene.height;
  const worst = Math.max(dx, dy, dw, dh);
  if (worst <= tol) return null;
  return {
    msg: `${label} got[${fmt(got)}] want[${fmt(want)}] (${(worst * 100).toFixed(1)}%)`,
    metric: worst * 100,
  };
}

/** 2 つの矩形の四辺のずれの最大（pt）。標準サイズの残差の報告用 */
const edgeResidualPt = (got: PxBox, want: PxBox): number =>
  Math.max(
    Math.abs(got.x - want.x),
    Math.abs(got.y - want.y),
    Math.abs(got.x + got.w - (want.x + want.w)),
    Math.abs(got.y + got.h - (want.y + want.h)),
  ) / K;

/** 標準サイズ（360×216pt）の残差（pt）：プロットの内側・凡例の項目の外接・タイトルの中心 */
interface Residual {
  plot?: number;
  legend?: number;
  title?: number;
}
/** グラフ名 → 残差（標準サイズのものだけ） */
const residuals = new Map<string, Residual>();

// ---- 照合 ----

/** 1 つのグラフを基準値と比べ、ずれを全部返す */
function compare(c: Case): Diff[] {
  const out: Diff[] = [];
  const add = (cat: Category, msg: string, metric?: number): void => {
    out.push({ name: c.name, cat, msg, metric });
  };
  const { exp, scene, model } = c;
  const std = exp.ChartObject.Width === 360 && exp.ChartObject.Height === 216;
  const res: Residual = {};
  if (std) residuals.set(c.name, res);
  const round = model.type === 'pie' || model.type === 'doughnut';
  const tol = c.name.startsWith('manual_') ? 0.01 : 0.03;

  // 数値軸（C1・C2）
  if (!round && exp.ValueAxis !== undefined) {
    const ax = axisOf(scene, model);
    const w = exp.ValueAxis;
    if (ax === null) add('axis', '目盛が出ていない');
    else if (
      !near(ax.min, w.MinimumScale, 1e-9) ||
      !near(ax.max, w.MaximumScale, 1e-9) ||
      !near(ax.major, w.MajorUnit, 1e-9 * Math.max(1, Math.abs(w.MajorUnit)))
    ) {
      // ずれの大きさ：端は範囲に対する割合・刻みは刻みに対する割合の大きい方（%）
      const span = Math.max(1e-12, w.MaximumScale - w.MinimumScale);
      const metric =
        Math.max(
          Math.abs(ax.min - w.MinimumScale) / span,
          Math.abs(ax.max - w.MaximumScale) / span,
          Math.abs(ax.major - w.MajorUnit) / Math.max(1e-12, Math.abs(w.MajorUnit)),
        ) * 100;
      add(
        'axis',
        `got[${ax.min}..${ax.max} /${ax.major}] want[${w.MinimumScale}..${w.MaximumScale} /${w.MajorUnit}]`,
        metric,
      );
    }
  }

  // 系列と点の色（C3）・線の太さ・マーカー（C8）・円の枠線（C9）
  if (model.type === 'bar') {
    const got: string[] = [];
    for (const it of scene.items) {
      if (it.kind !== 'path' || it.fill === null) continue;
      if (!/^M[\d.\- ]+L[\d.\- ]+L[\d.\- ]+L[\d.\- ]+Z/.test(it.d)) continue;
      if (!got.includes(it.fill.color)) got.push(it.fill.color);
    }
    const vary = exp.Group.VaryByCategories && exp.Series.length === 1;
    const want = vary
      ? exp.Series[0]!.Points!.map((p) => hexOf(p.Fill.RGB))
      : exp.Series.map((s) => hexOf(s.Fill.RGB));
    if (got.join() !== want.join()) add('color', `got[${got.join()}] want[${want.join()}]`);
  } else if (model.type === 'line') {
    const lines = scene.items.filter(
      (it): it is PathItem =>
        it.kind === 'path' && it.fill === null && it.stroke !== null && it.stroke.cap === 'round',
    );
    const markers = scene.items.filter(
      (it): it is PathItem => it.kind === 'path' && /A/.test(it.d) && !/L/.test(it.d),
    );
    const vary = exp.Group.VaryByCategories && exp.Series.length === 1;
    const gotC = lines.map((l) => l.stroke!.color);
    const wantC = exp.Series.map((s) => hexOf(s.Line.RGB));
    if (gotC.join() !== wantC.join()) add('color', `線 got[${gotC.join()}] want[${wantC.join()}]`);
    // varyColors の折れ線：マーカーは点ごとの色（COM は点の Line の色）
    const wantM = vary ? exp.Series[0]!.Points!.map((p) => hexOf(p.Line.RGB)) : wantC;
    const gotW = lines.map((l) => l.stroke!.width);
    const wantW = exp.Series.map((s) => s.Line.Weight * K);
    if (gotW.length !== wantW.length || gotW.some((x, i) => !near(x, wantW[i]!, 0.05))) {
      add(
        'width',
        `線の太さ got[${gotW.map((x) => x.toFixed(2)).join()}] want[${wantW.map((x) => x.toFixed(2)).join()}]`,
      );
    }
    const hasMarker = exp.Series.some((s) => s.MarkerStyle !== -4142);
    if (hasMarker !== markers.length > 0) {
      add('marker', `マーカー got[${markers.length > 0}] want[${hasMarker}]`);
    } else if (hasMarker) {
      const gotSet = [...new Set(markers.map((m) => m.fill?.color ?? '(なし)'))];
      if (gotSet.join() !== wantM.join())
        add('marker', `色 got[${gotSet.join()}] want[${wantM.join()}]`);
      const diam = cmds(markers[0]!.d).find((x) => x.c === 'A')!.n[0]! * 2;
      const wantD = exp.Series[0]!.MarkerSize * K;
      if (!near(diam, wantD, 0.1)) {
        add('marker', `大きさ got[${diam.toFixed(2)}px] want[${wantD.toFixed(2)}px]`);
      }
      // 形：基準が自動（-4105）か丸（8）なら丸（A 命令）。他の形は未対応
      const st = exp.Series[0]!.MarkerStyle;
      if (st !== -4105 && st !== 8) add('marker', `形 ${st} は未対応`);
    }
  } else {
    const got: string[] = [];
    const widths = new Set<number>();
    for (const it of scene.items) {
      if (it.kind !== 'path' || it.fill === null || !/A/.test(it.d)) continue;
      got.push(it.fill.color);
      if (it.stroke !== null) widths.add(Math.round(it.stroke.width * 100) / 100);
    }
    const want: string[] = [];
    const rings = model.type === 'pie' ? exp.Series.slice(0, 1) : exp.Series;
    for (const s of rings) for (const p of s.Points ?? []) want.push(hexOf(p.Fill.RGB));
    if (got.join() !== want.join()) add('color', `got[${got.join()}] want[${want.join()}]`);
    const w0 = exp.Series[0]?.Line.Weight;
    if (w0 !== undefined && w0 > 0 && (widths.size !== 1 || !near([...widths][0]!, w0 * K, 0.05))) {
      add('width', `扇の枠線 got[${[...widths].join()}px] want[${(w0 * K).toFixed(2)}px]`);
    }
  }

  // プロットの内側（C5）
  const gotPlot = round ? pieBox(scene) : gridBox(scene);
  if (gotPlot === null) add('plot', 'プロットの外接矩形が取れない');
  else {
    const want = px(exp.PlotArea.Inside);
    res.plot = edgeResidualPt(gotPlot, want);
    const d = boxDiff('内側', gotPlot, want, scene, tol);
    if (d !== null) add('plot', d.msg, d.metric);
  }

  // タイトル（C5・C10）：文字の中心。文字の無い自動のタイトル（系列 2 つ以上）は、Excel が言語ごとの代替の文字を出す＝文字は描かない
  // （帯の分のプロットの位置だけ「プロット」で見る）ので、文字の位置は照合しない
  const blankAutoTitle =
    model.title !== undefined && model.title.lines.length === 0 && model.series.length >= 2;
  const titleText =
    exp.HasTitle && exp.Title !== undefined ? exp.Title.Text : (model.series[0]?.name ?? '');
  const title = scene.items.find(
    (i): i is TextItem =>
      i.kind === 'text' && i.baseline === 'middle' && i.anchor === 'middle' && i.text === titleText,
  );
  if (blankAutoTitle) {
    if (
      scene.items.some((i) => i.kind === 'text' && i.baseline === 'middle' && i.anchor === 'middle')
    ) {
      add('title', '文字の無い自動のタイトルのはずなのに、文字が出ている');
    }
  } else if (exp.HasTitle && exp.Title !== undefined) {
    if (title === undefined) {
      add('title', `タイトル「${exp.Title.Text}」の文字が無い`);
    } else {
      const b = px(exp.Title.Pos);
      const wx = b.x + b.w / 2;
      const wy = b.y + b.h / 2;
      res.title = Math.max(Math.abs(title.x - wx), Math.abs(title.y - wy)) / K;
      const ex = Math.abs(title.x - wx) / scene.width;
      const ey = Math.abs(title.y - wy) / scene.height;
      if (ex > tol || ey > tol) {
        add(
          'title',
          `中心 got[${title.x.toFixed(1)},${title.y.toFixed(1)}] want[${wx.toFixed(1)},${wy.toFixed(1)}]`,
          Math.max(ex, ey) * 100,
        );
      }
      // 既定の文字（C4）：大きさ・太字・色
      const f = exp.Title.Font;
      const bold = f.Bold === true || f.Bold === -1;
      if (
        !near(title.size, f.Size * K, 0.05) ||
        title.bold !== bold ||
        title.color.color !== hexOf(f.RGB)
      ) {
        add(
          'style',
          `タイトル got[${title.size.toFixed(2)}px bold=${title.bold} ${title.color.color}] want[${(f.Size * K).toFixed(2)}px bold=${bold} ${hexOf(f.RGB)}]`,
        );
      }
    }
  } else if (title !== undefined) add('title', 'タイトルが無いはずなのに出ている');

  // 凡例（C5・C6）：項目の外接矩形（COM の LegendEntries の外接）
  const lb = legendBox(scene);
  if (exp.HasLegend && exp.Legend !== undefined) {
    if (lb === null) add('legend', '凡例が出ていない');
    else {
      const es = exp.Legend.Entries.map((e) => px(e.Pos));
      const x0 = Math.min(...es.map((e) => e.x));
      const y0 = Math.min(...es.map((e) => e.y));
      const x1 = Math.max(...es.map((e) => e.x + e.w));
      const y1 = Math.max(...es.map((e) => e.y + e.h));
      const want = { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
      res.legend = edgeResidualPt(lb, want);
      const d = boxDiff('項目', lb, want, scene, tol);
      if (d !== null) add('legend', d.msg, d.metric);
      const lab = scene.items.find(
        (i): i is TextItem => i.kind === 'text' && i.anchor === 'start' && i.baseline === 'middle',
      );
      const f = exp.Legend.Font;
      if (
        lab !== undefined &&
        (!near(lab.size, f.Size * K, 0.05) || lab.color.color !== hexOf(f.RGB))
      ) {
        add(
          'style',
          `凡例の文字 got[${lab.size.toFixed(2)}px ${lab.color.color}] want[${(f.Size * K).toFixed(2)}px ${hexOf(f.RGB)}]`,
        );
      }
    }
  } else if (lb !== null) add('legend', '凡例が無いはずなのに出ている');

  // 既定（C4）：グラフの面の枠線・軸の文字の大きさと色
  const ca = scene.items[0];
  if (ca?.kind === 'rect' && ca.width === scene.width) {
    const s = ca.stroke;
    const wl = exp.ChartArea.Line;
    if (s === null) add('style', 'グラフの面の枠線が無い');
    else if (!near(s.width, wl.Weight * K, 0.05) || s.color !== hexOf(wl.RGB)) {
      add(
        'style',
        `面の枠線 got[${s.width.toFixed(2)}px ${s.color}] want[${(wl.Weight * K).toFixed(2)}px ${hexOf(wl.RGB)}]`,
      );
    }
  }
  if (!round && exp.ValueAxis !== undefined) {
    const lab = valueLabelTexts(scene, model)[0];
    const f = exp.ValueAxis.TickLabels.Font;
    if (
      lab !== undefined &&
      (!near(lab.size, f.Size * K, 0.05) || lab.color.color !== hexOf(f.RGB))
    ) {
      add(
        'style',
        `軸の文字 got[${lab.size.toFixed(2)}px ${lab.color.color}] want[${(f.Size * K).toFixed(2)}px ${hexOf(f.RGB)}]`,
      );
    }
  }
  return out;
}

// ---- 読み込み ----

const cases: Case[] = [];
let all: Diff[] = [];

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

beforeAll(async () => {
  const expected = JSON.parse(readFileSync(fixtures('ref-charts.expected.json'), 'utf8')) as {
    charts: Record<string, ComChart>;
    defaults?: Record<string, ComChart>;
  };
  const sets: [Case['group'], string, Record<string, ComChart>][] = [
    ['charts', 'ref-charts.xlsx', expected.charts],
    ['defaults', 'ref-charts-default.xlsx', expected.defaults ?? {}],
  ];
  for (const [group, file, exps] of sets) {
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
      cases.push({ name, group, exp, model: r.model, theme: r.theme, scene });
    }
  }
  all = cases.flatMap(compare);
  // 差の一覧と、標準サイズ（360×216pt）の残差（pt）を書き出す（開発用：REF104_DUMP=<ファイル> npx vitest run …）
  if (process.env.REF104_DUMP) {
    const dump = {
      diffs: all,
      stdResidualPt: residualSummary(),
      stdResiduals: Object.fromEntries(residuals),
    };
    writeFileSync(process.env.REF104_DUMP, JSON.stringify(dump, null, 1));
  }
});

/** 標準サイズの残差の最大（pt）：KNOWN の組を除く最大・含めた最大と、そのグラフ名 */
interface ResidualMax {
  exceptKnown: number;
  exceptKnownAt: string;
  all: number;
  allAt: string;
}

/** 標準サイズ（360×216pt）のグラフの残差（pt）の最大を、プロット・凡例・タイトルごとに出す */
function residualSummary(): Record<keyof Residual, ResidualMax> {
  const out = {} as Record<keyof Residual, ResidualMax>;
  for (const k of ['plot', 'legend', 'title'] as const) {
    const r: ResidualMax = { exceptKnown: 0, exceptKnownAt: '', all: 0, allAt: '' };
    for (const [name, v] of residuals) {
      const x = v[k];
      if (x === undefined) continue;
      if (x > r.all) {
        r.all = x;
        r.allAt = name;
      }
      if (KNOWN[`${k}:${name}`] === undefined && x > r.exceptKnown) {
        r.exceptKnown = x;
        r.exceptKnownAt = name;
      }
    }
    out[k] = r;
  }
  return out;
}

describe('TD-104-0 基準値の読み込み', () => {
  it('ref-charts.xlsx の 113 個と ref-charts-default.xlsx の 19 個が、すべて読めて場面になる', () => {
    expect(cases.filter((c) => c.group === 'charts')).toHaveLength(113);
    expect(cases.filter((c) => c.group === 'defaults')).toHaveLength(19);
  });

  it('TD-106-9 参照のグラフは全部が密でない：間引きの有無（thin）で場面が完全に同じ（ADR-0031 §決定 11.1）', () => {
    expect(cases.length).toBeGreaterThan(0);
    const differing = cases
      .filter((c) => {
        const w = c.exp.ChartObject.Width * K;
        const h = c.exp.ChartObject.Height * K;
        const off = layoutChart(c.model, w, h, c.theme, { thin: false });
        const on = layoutChart(c.model, w, h, c.theme, { thin: true });
        return JSON.stringify(on) !== JSON.stringify(off);
      })
      .map((c) => c.name);
    expect(differing).toEqual([]);
  });
});

/**
 * 合わせられない組の記録。外した組が今より悪化したら失敗させる（直ったら失敗する＝一覧から消す）。
 * - `max`：ずれ（%）の許容の上限。今の値を小数 2 桁へ切り上げた値（位置・大きさは面に対する割合・軸は範囲／刻みに対する割合）。
 * - `msg`：数で測れない差（色・太さ）の今のメッセージ。完全に同じでなければ（＝差の中身が変われば）失敗にして見直す。
 */
type Known = { why: string } & ({ max: number } | { msg: string });
const lim = (max: number, why: string): Known => ({ why, max });
const same = (msg: string, why: string): Known => ({ why, msg });

/** 合わせられない組（`<種類>:<グラフ名>` → 理由と今の値） */
const KNOWN: Record<string, Known> = {
  // C1：軸の自動範囲（目盛の数の上限の規則だけでは合わない組）
  'axis:barStacked100_neg_std': lim(
    150,
    'C1：横棒の 100% 積み上げ（-100%〜100%）だけ、Excel の刻みは 0.2（こちらは 0.5）。標準の幅で目盛 10 個を許すのはこの組だけで、同じ幅の他の組（0〜1・-1〜0 の 100% 積み上げ・負を含む積み上げ）は 10 個を許さない。3 回目の幅の掃引でも説明できず暫定のまま',
  ),
  // C7（ISSUE-107）：cat30_* は項目名を 90° に回して 2 個おきに間引く規則で合った（KNOWN から外した）
  'plot:catlong_colClustered_small': lim(
    4.49,
    'C7（ISSUE-107）：長い項目名（21 文字）の小さい面（216×144pt）で Excel は 45° に回して帯を 25pt に厚くする。こちらは推定の幅が広く、回した帯が面の高さの 0.4 を超えるので安全弁で回さず間引く（今までどおり）。暫定のまま',
  ),
  'plot:catlong_colClustered_std': lim(
    26.39,
    'C7（ISSUE-107）：長い項目名（21 文字）で Excel は 45° に回して左に張り出す。こちらは推定の幅が広く（帯 117px・面の高さの 0.4＝115px）安全弁で回さない。安全弁を 0.45 にすると回るが、プロットの高さが Excel より 17px 低く数値軸の刻みが変わる（1 回目の KNOWN が増える）ので 0.4 のまま。文字幅の推定は別 ISSUE',
  ),
  'plot:catlong_barClustered_small': lim(
    7.65,
    'C7：長い項目名の横棒で Excel も水平 1 行。差は文字幅の推定（項目名の帯が広く、プロットの左が 155px・Excel は 133px）で、横棒は回さない（ISSUE-107）。文字幅の推定は別 ISSUE',
  ),
  // C5・C6：凡例・タイトル
  'legend:legend_b_lineMarkers_std': lim(
    7.25,
    'C6：折れ線の凡例は記号の右に記号と同じ長さの空きがあり、横並びで 7%（右・左の縦並びは 3% 以内）。暫定のまま',
  ),
  'legend:manual_colClustered_std': lim(
    15.54,
    'C5：手動の凡例の枠（70pt 角）で Excel は 3 項目を横 1 行に並べる。枠の形で並べ方を変えない暫定のまま',
  ),
  'title:manual_colClustered_std': lim(
    1.06,
    'C5：手動のタイトルの中心が 1.05%（許容 1%）。文字の幅の推定（Calibri 14pt の Title で 51px・実際は 41px）の差。暫定のまま',
  ),
  // C3・C9
  'color:def_colClustered_ser8': same(
    'got[#4472c4,#ed7d31,#a5a5a5,#ffc000,#5b9bd5,#70ad47,#264478,#9e480e] want[#3b64ad,#d26e2a,#929292,#e2aa00,#5089bc,#62993e,#8fa2d4,#f1a78a]',
    'C3：c:spPr・c:style の無い 8 系列の最初の 6 色は、Excel が系列の数で濃淡を変える（暫定のまま。Excel の保存したグラフは明示の色を持つ）',
  ),
  'color:def_pie_pt8': same(
    'got[#4472c4,#ed7d31,#a5a5a5,#ffc000,#5b9bd5,#70ad47,#264478,#9e480e] want[#3b64ad,#d26e2a,#929292,#e2aa00,#5089bc,#62993e,#8fa2d4,#f1a78a]',
    'C3：同上（円の 8 点）',
  ),
  'width:vary_lineMarkers_1ser_on_std': same(
    '線の太さ got[2.00] want[3.00]',
    'C9：varyColors の折れ線の太さ 2.25pt は点（c:dPt）に書かれ、点の線は読まない。系列に明示が無ければ既定 1.5pt（c:spPr の無いグラフの値）',
  ),
};
// 文字の無い自動のタイトル（title_auto_2ser）は、帯だけ取り文字は描かない（C10）ので KNOWN から外した（プロットは合う・文字の位置は照合しない）。
// 標準サイズ（360×216pt）の grid（種類×データ）の「プロットの内側の各辺が 1.5pt 以内」という補助の検査は置かない：
// 40 組のうち 16 組が超える（最大 6.0pt・負を含む組と 100% 積み上げ）。残差は REF104_DUMP の stdResidualPt で見る。

const label: Record<Category, string> = {
  axis: 'TD-104-1 数値軸の最小・最大・刻み（C1・C2）',
  color: 'TD-104-2 系列と点の色（C3）',
  marker: 'TD-104-3 マーカーの形・大きさ・色（C8）',
  width: 'TD-104-4 線・扇の枠線の太さ（C8・C9）',
  plot: 'TD-104-5 プロットの内側の位置と大きさ（C5・C9）',
  legend: 'TD-104-6 凡例の位置と大きさ（C5・C6）',
  title: 'TD-104-7 タイトルの位置（C5・C10）',
  style: 'TD-104-8 既定の枠線・文字の大きさと色（C4）',
};

/** 外した組の差が、記録（上限・メッセージ）より悪化しているか */
function worseThanKnown(d: Diff, k: Known): boolean {
  if ('max' in k) return d.metric === undefined || d.metric > k.max + 1e-9;
  return d.msg !== k.msg;
}

describe('Excel の COM の値との照合', () => {
  for (const cat of Object.keys(label) as Category[]) {
    it(label[cat], () => {
      const diffs = all.filter((d) => d.cat === cat);
      const unexpected = diffs.filter((d) => KNOWN[`${cat}:${d.name}`] === undefined);
      expect(unexpected.map((d) => `${d.name}: ${d.msg}`)).toEqual([]);
      // 外した組が今より悪化したら失敗（上限・メッセージを超える）
      const worse = diffs.filter((d) => {
        const k = KNOWN[`${cat}:${d.name}`];
        return k !== undefined && worseThanKnown(d, k);
      });
      expect(worse.map((d) => `${d.name}: ${d.msg}`)).toEqual([]);
      // 外した組が直っていたら、一覧から消す
      const stale = Object.keys(KNOWN).filter(
        (k) => k.startsWith(`${cat}:`) && !diffs.some((d) => `${cat}:${d.name}` === k),
      );
      expect(stale).toEqual([]);
    });
  }
});
