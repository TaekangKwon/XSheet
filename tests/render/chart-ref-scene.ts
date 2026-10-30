/// <reference types="node" />
/**
 * Excel 16 の COM の値とグラフの場面を比べるための、場面から量を取り出す補助（ISSUE-105・`xlsx-chart-ref2.test.ts` が使う）。
 * 1 回目の照合（`xlsx-chart-ref.test.ts`・ISSUE-104）の同名の補助を写したもの：1 回目のテストは動かさない（約 900 検査を保つ）ため、重ねて持つ。
 * 後で 1 回目のテストをこちらへ寄せて一本化してよい。新しい分は凡例の項目（記号の色つき）・補助の目盛・目盛の印・マーカーの取り出しと、基準値の型・外した組の記録の器。
 * 単位：COM は pt。場面は 96dpi の px なので pt × 96/72。COM の座標の原点はグラフの面の左上から {@link COM_ORIGIN_PT} だけ内側。
 */
import type { ChartModel } from '../../src/core/types';
import {
  estimateTextWidth,
  PX_PER_PT,
  type ChartScene,
  type ChartStroke,
} from '../../src/render/chart-layout';

// ---- 基準値（COM）の型（使う所だけ） ----
export interface Box {
  Left: number;
  Top: number;
  Width: number;
  Height: number;
}
export interface ComPoint {
  Fill: { RGB: string | number };
  Line: { RGB: string | number };
}
export interface ComSeries {
  Name: string;
  Fill: { RGB: string | number };
  Line: { RGB: string | number; Weight: number };
  MarkerStyle: number;
  MarkerSize: number;
  PlotOrder?: number;
  Points?: ComPoint[];
}
export interface ComEntry {
  Pos: Box;
  KeyFillRGB?: string | number;
  KeyLineRGB?: string | number;
}
export interface ComAxis {
  Pos?: Box;
  MinimumScale: number;
  MaximumScale: number;
  MajorUnit: number;
  MinorUnit?: number;
  MajorTickMark?: number;
  MinorTickMark?: number;
  TickLabelSpacing?: number;
  TickLabels: { Orientation?: number; Font: { Size: number; RGB: string | number } };
}
export interface ComChart {
  Kind?: string;
  Check?: string;
  ChartStyle?: number;
  ChartObject: { Width: number; Height: number };
  ChartArea: {
    Fill?: { RGB: string | number; Visible?: boolean | number };
    Line: { Weight: number; RGB: string | number };
  };
  PlotArea: { Inside: Box };
  HasTitle: boolean;
  Title?: {
    Text: string;
    Pos: Box;
    Font: { Size: number; Bold: number | boolean; RGB: string | number };
  };
  HasLegend: boolean;
  Legend?: { Pos: Box; Font: { Size: number; RGB: string | number }; Entries: ComEntry[] };
  CategoryAxis?: ComAxis;
  ValueAxis?: ComAxis;
  Group: { VaryByCategories: boolean };
  Series: ComSeries[];
  Ticks?: {
    CatMajor?: number;
    CatMinor?: number;
    ValMajor?: number;
    ValMinor?: number;
    MarkerStyles?: number[];
    MarkerSizes?: number[];
  };
}

export type TextItem = Extract<ChartScene['items'][number], { kind: 'text' }>;
export type PathItem = Extract<ChartScene['items'][number], { kind: 'path' }>;
export type LineItem = Extract<ChartScene['items'][number], { kind: 'line' }>;

export const K = PX_PER_PT;
/** COM の座標の原点のずれ（pt）。根拠は `xlsx-chart-ref.test.ts` の COM_ORIGIN_PT */
export const COM_ORIGIN_PT = 4;
export const fixtures = (n: string): URL => new URL(`../../fixtures/excel/${n}`, import.meta.url);

/** COM の色（`RRGGBB`）を `#rrggbb` にする（数値＝色の定数はそのまま `?n`） */
export const hexOf = (c: string | number | undefined): string =>
  typeof c === 'string' ? `#${c.toLowerCase()}` : `?${String(c)}`;
export const near = (a: number, b: number, eps: number): boolean => Math.abs(a - b) <= eps;

// ---- 経路・矩形 ----
export interface PxBox {
  x: number;
  y: number;
  w: number;
  h: number;
}
interface Cmd {
  c: string;
  n: number[];
}

/** 経路の `d`（数値と `M L C A Z` だけ）を命令ごとに分ける */
export function cmds(d: string): Cmd[] {
  const out: Cmd[] = [];
  for (const m of d.matchAll(/([MLCAZ])([^MLCAZ]*)/g)) {
    out.push({ c: m[1]!, n: (m[2]!.match(/-?\d+(?:\.\d+)?(?:e-?\d+)?/gi) ?? []).map(Number) });
  }
  return out;
}

/** 経路の点（M・L の座標・A の終点）の外接矩形 */
export function pathBox(d: string): PxBox | null {
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

/** `M x y L x y` の組（折れ線でない目盛線・目盛の印）を線分にする。他の命令が混じる経路は null */
export function segments(d: string): [number, number, number, number][] | null {
  const cs = cmds(d);
  const out: [number, number, number, number][] = [];
  for (let i = 0; i < cs.length; i += 2) {
    const m = cs[i];
    const l = cs[i + 1];
    if (m?.c !== 'M' || l?.c !== 'L' || m.n.length < 2 || l.n.length < 2) return null;
    out.push([m.n[0]!, m.n[1]!, l.n[0]!, l.n[1]!]);
  }
  return out;
}

/**
 * 目盛線の経路（塗り無し・butt・M/L だけ）の外接矩形＝プロットの内側。無ければ null。
 * 補助の目盛線は端に線が無く小さい外接になる・目盛の印は細いので、面積が最大のものを取る。
 */
export function gridBox(scene: ChartScene): PxBox | null {
  let best: PxBox | null = null;
  for (const it of scene.items) {
    if (it.kind !== 'path' || it.fill !== null || it.stroke === null) continue;
    if (it.stroke.cap !== 'butt' || /[ACZ]/.test(it.d)) continue;
    const b = pathBox(it.d);
    if (b !== null && (best === null || b.w * b.h > best.w * best.h)) best = b;
  }
  return best;
}

/** 円・ドーナツの外接の正方形（最大の外半径の扇の中心と半径。開始角 0 の前提） */
export function pieBox(scene: ChartScene): PxBox | null {
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

/** プロットの内側（円・ドーナツは外接の正方形） */
export function plotOf(scene: ChartScene, model: ChartModel): PxBox | null {
  return model.type === 'pie' || model.type === 'doughnut' ? pieBox(scene) : gridBox(scene);
}

// ---- 数値軸の目盛の文字 ----

/**
 * 数値軸の目盛の文字（向きごとの anchor・baseline・回転・位置で絞る。項目名・凡例・タイトルが数字でも拾わない）。
 * 縦の数値軸（縦棒・折れ線）：anchor end・baseline middle・回転なしで、プロットの左。横の数値軸（横棒）：anchor middle・baseline hanging で、プロットの下。
 */
export function valueLabelTexts(scene: ChartScene, model: ChartModel): TextItem[] {
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

/** 軸の最小・最大・刻み（場面の目盛の文字から。`%` は 100 で割る）。目盛が 2 つ未満なら null */
export function axisOf(
  scene: ChartScene,
  model: ChartModel,
): { min: number; max: number; major: number } | null {
  const v = valueLabelTexts(scene, model).map((it) =>
    it.text.endsWith('%') ? Number(it.text.slice(0, -1)) / 100 : Number(it.text),
  );
  if (v.length < 2) return null;
  return { min: v[0]!, max: v[v.length - 1]!, major: (v[v.length - 1]! - v[0]!) / (v.length - 1) };
}

// ---- 目盛線・軸線の線（PNG との比較用） ----

/** 場面の目盛線・軸線の見た目（無ければ undefined）。値の方向の目盛線は主・補助、軸線は項目軸・数値軸 */
export interface SceneStrokes {
  major?: ChartStroke;
  minor?: ChartStroke;
  catLine?: ChartStroke;
  valLine?: ChartStroke;
}

/**
 * 場面から、値の方向の目盛線（補助 → 主の順に出る）と、プロットの縁・0 の位置に引かれた軸線の見た目を取り出す。
 * 軸線は `line` の要素のうち、プロットの幅（高さ）と同じ長さのもの：縦棒・折れ線は横線が項目軸・縦線が数値軸、横棒は逆。
 */
export function axisStrokes(scene: ChartScene, model: ChartModel): SceneStrokes {
  const plot = gridBox(scene);
  if (plot === null) return {};
  const vert = model.barDir !== 'bar';
  const grids = fullGridPaths(scene, plot, vert);
  const hasMinor = model.valAxis?.minorGridlines !== undefined;
  const out: SceneStrokes = {};
  const major = grids[hasMinor ? 1 : 0];
  if (major?.stroke) out.major = major.stroke;
  const minor = hasMinor ? grids[0] : undefined;
  if (minor?.stroke) out.minor = minor.stroke;
  for (const it of scene.items) {
    if (it.kind !== 'line') continue;
    const horizontal = near(it.y1, it.y2, 1e-6) && near(Math.abs(it.x2 - it.x1), plot.w, 1);
    const vertical = near(it.x1, it.x2, 1e-6) && near(Math.abs(it.y2 - it.y1), plot.h, 1);
    if (horizontal === vertical) continue;
    const isCat = vert ? horizontal : vertical;
    if (isCat) out.catLine ??= it.stroke;
    else out.valLine ??= it.stroke;
  }
  return out;
}

// ---- 補助の目盛（MinorUnit） ----

/** 値の方向の目盛線（プロットの全長の線だけでできた経路）を、場面に出てきた順（補助 → 主）に返す */
function fullGridPaths(scene: ChartScene, plot: PxBox, horizontalLines: boolean): PathItem[] {
  return scene.items.filter((it): it is PathItem => {
    if (it.kind !== 'path' || it.fill !== null || it.stroke === null || it.stroke.cap !== 'butt')
      return false;
    const segs = segments(it.d);
    if (segs === null || segs.length === 0) return false;
    return segs.every(([x1, y1, x2, y2]) =>
      horizontalLines
        ? near(y1, y2, 1e-6) && near(Math.abs(x2 - x1), plot.w, 1)
        : near(x1, x2, 1e-6) && near(Math.abs(y2 - y1), plot.h, 1),
    );
  });
}

/**
 * 場面の補助の刻み（値の単位）。補助の目盛線（主と重ならない位置）か補助の目盛の印（短い線分）の数から、主の刻みを割って出す。
 * 補助の線・印が場面に無ければ null。`axis` は {@link axisOf} の結果。
 */
export function minorUnitOf(
  scene: ChartScene,
  model: ChartModel,
  axis: { min: number; max: number; major: number },
): number | null {
  const plot = gridBox(scene);
  if (plot === null) return null;
  const vert = model.barDir !== 'bar';
  const intervals = (axis.max - axis.min) / axis.major;
  const perMajor = (count: number): number => axis.major / (count / intervals + 1);
  const va = model.valAxis;
  if (va?.minorGridlines !== undefined) {
    const p = fullGridPaths(scene, plot, vert)[0];
    if (p === undefined) return null;
    return perMajor(segments(p.d)!.length);
  }
  const mark = va?.minorTickMark;
  if (mark === undefined || mark === 'none') return null;
  const signs = mark === 'cross' ? 2 : 1;
  const tick = scene.items.find((it): it is PathItem => {
    if (it.kind !== 'path' || it.fill !== null || it.stroke === null) return false;
    const segs = segments(it.d);
    if (segs === null || segs.length === 0) return false;
    // 短い線分だけ（長さは minorTickLengthPt の px。主の目盛の印・目盛線は除く）
    const len = (s: [number, number, number, number]): number =>
      Math.hypot(s[2] - s[0], s[3] - s[1]);
    return segs.every((s) => near(len(s), 3 * K, 0.2));
  });
  if (tick === undefined) return null;
  return perMajor(segments(tick.d)!.length / signs);
}

// ---- 凡例 ----

/** 凡例の項目 1 つ（場面から）：文字・文字の位置・記号の色（四角は塗り、線は線の色） */
export interface SceneLegendItem {
  text: string;
  x: number;
  y: number;
  keyColor: string | null;
}

/** 凡例の文字（anchor start・baseline middle）と、同じ高さで左にある記号の色。見た目の順（行 → 左から）に並べる */
export function legendItems(scene: ChartScene): SceneLegendItem[] {
  const texts = scene.items.filter(
    (i): i is TextItem => i.kind === 'text' && i.anchor === 'start' && i.baseline === 'middle',
  );
  const out: SceneLegendItem[] = texts.map((t) => {
    let key: { x: number; color: string } | null = null;
    for (const it of scene.items) {
      let kx: number | undefined;
      let ky: number | undefined;
      let color: string | undefined;
      if (it.kind === 'rect' && it.width < 20 && it.fill !== null) {
        kx = it.x;
        ky = it.y + it.height / 2;
        color = it.fill.color;
      } else if (it.kind === 'line' && it.y1 === it.y2 && it.x2 - it.x1 < 40) {
        kx = it.x1;
        ky = it.y1;
        color = it.stroke.color;
      }
      if (kx !== undefined && ky !== undefined && color !== undefined) {
        if (Math.abs(ky - t.y) < 0.5 && kx < t.x && (key === null || kx > key.x))
          key = { x: kx, color };
      }
    }
    return { text: t.text, x: t.x, y: t.y, keyColor: key?.color ?? null };
  });
  return sortVisual(out, (e) => ({ x: e.x, y: e.y }), 3 * K);
}

/** 見た目の順（行 → 左から）。行の高さ（rowTol）以内の縦の差は同じ行とみなす */
export function sortVisual<T>(
  list: T[],
  pos: (e: T) => { x: number; y: number },
  rowTol: number,
): T[] {
  const rows: { y: number; items: T[] }[] = [];
  for (const e of [...list].sort((a, b) => pos(a).y - pos(b).y)) {
    const row = rows.find((r) => Math.abs(r.y - pos(e).y) <= rowTol);
    if (row) row.items.push(e);
    else rows.push({ y: pos(e).y, items: [e] });
  }
  return rows.flatMap((r) => r.items.sort((a, b) => pos(a).x - pos(b).x));
}

/** 凡例の項目の外接矩形（文字の幅は推定・高さは文字の大きさの 1.2 倍）。記号は含める。無ければ null */
export function legendBox(scene: ChartScene): PxBox | null {
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

/**
 * COM の凡例の項目を、見た目の順に（Top・Left）並べ、項目ごとに「記号の色 → 名前」を引く。
 * 名前の引き方：点ごとの凡例（円・ドーナツ、または 1 系列で `VaryByCategories`）は、点の塗りの色から `model.categories` の名前。
 * 円は `VaryByCategories` が偽でも点ごと（全部が同じ色になるので、見た目の順の位置から `model.categories` の名前。C3）。
 * 系列ごとの凡例は、棒・円は系列の塗り、折れ線は線の色から系列名。引けない項目は null。
 * COM の `LegendEntry` に文字の値は無いため、この対応で「どの項目が残ったか」を言う（C11・C3・C6）。
 */
export function expectedLegend(
  exp: ComChart,
  model: ChartModel,
): { names: (string | null)[]; boxes: PxBox[] } {
  // 面の外（右端を越える位置）や凡例の枠の下に置かれた項目は、凡例に入りきらず描かれない（COM は項目数だけ返す。
  // c3vary0_pie_pt13_r：13 項目のうち 2 つは Left 585・612pt、1 つは凡例（高さ 168.75pt）の下端を越える Top 204pt で、PNG に出るのは 10 項目）
  const legBottom =
    exp.Legend === undefined ? Infinity : exp.Legend.Pos.Top + exp.Legend.Pos.Height;
  const entries = (exp.Legend?.Entries ?? []).filter(
    (e) =>
      e.Pos.Left + e.Pos.Width <= exp.ChartObject.Width + 1 &&
      e.Pos.Top + e.Pos.Height <= legBottom + 1,
  );
  const sorted = sortVisual(
    entries.map((e) => ({ e, p: px(e.Pos) })),
    (o) => ({ x: o.p.x, y: o.p.y }),
    3 * K,
  );
  const round = model.type === 'pie' || model.type === 'doughnut';
  const byPoint = exp.Group.VaryByCategories && (round || exp.Series.length === 1);
  // 円は varyColors が偽でも項目（点）ごとに出る（Excel 16 の COM・PNG で確認）。色が同じなので位置で名前を引く
  const pieFlat = model.type === 'pie' && !exp.Group.VaryByCategories;
  const names = sorted.map(({ e }, k) => {
    if (pieFlat) return model.categories[k] ?? null;
    if (byPoint) {
      const key = hexOf(e.KeyFillRGB);
      const i = (exp.Series[0]?.Points ?? []).findIndex((p) => hexOf(p.Fill.RGB) === key);
      return i < 0 ? null : (model.categories[i] ?? null);
    }
    const line = model.type === 'line';
    const key = hexOf(line ? e.KeyLineRGB : e.KeyFillRGB);
    const s = exp.Series.find((x) => hexOf(line ? x.Line.RGB : x.Fill.RGB) === key);
    return s === undefined ? null : s.Name;
  });
  return { names, boxes: sorted.map((o) => o.p) };
}

// ---- 座標・差 ----

/** COM の要素の座標（pt）を、グラフの面の左上を原点とする px にする（原点のずれ {@link COM_ORIGIN_PT} を足す） */
export const px = (b: Box): PxBox => ({
  x: (b.Left + COM_ORIGIN_PT) * K,
  y: (b.Top + COM_ORIGIN_PT) * K,
  w: b.Width * K,
  h: b.Height * K,
});
export const fmt = (b: PxBox): string =>
  `${b.x.toFixed(1)},${b.y.toFixed(1)} ${b.w.toFixed(1)}x${b.h.toFixed(1)}`;

/** 位置・大きさ（px）が許容（グラフの面に対する割合）内か。ずれを文字と大きさ（%）にして返す（合えば null） */
export function boxDiff(
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

/** 複数の矩形の外接矩形 */
export function unionBox(bs: PxBox[]): PxBox {
  const x0 = Math.min(...bs.map((e) => e.x));
  const y0 = Math.min(...bs.map((e) => e.y));
  const x1 = Math.max(...bs.map((e) => e.x + e.w));
  const y1 = Math.max(...bs.map((e) => e.y + e.h));
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

// ---- 差の記録・外した組 ----

/** 照合の差 1 件。`cat` は照合の種類（TD-105-n の分け方） */
export interface Diff {
  name: string;
  cat: string;
  msg: string;
  /** ずれの大きさ（%）。数で測れない差（色・文字・数）は無し */
  metric?: number;
}

/**
 * 合わせられない組の記録（xlsx-chart-ref.test.ts と同じ）。外した組が今より悪化したら失敗させる（直ったら失敗する＝一覧から消す）。
 * `max`：ずれ（%）の許容の上限（今の値を小数 2 桁へ切り上げ）。`msg`：数で測れない差の今のメッセージ（完全に同じでなければ見直す）。
 * 1 つのグラフに差が複数ある時は、数で測れる差は `max`、測れない差は `msgs`（今のメッセージの一覧）で記録する（どれにも当たらない差は悪化とみなす）。
 */
export interface Known {
  why: string;
  max?: number;
  msg?: string;
  msgs?: string[];
}
export const lim = (max: number, why: string): Known => ({ why, max });
export const same = (msg: string, why: string): Known => ({ why, msg });
/** 数で測れる差（上限 `max`）と、測れない差（`msgs`）が混ざる組 */
export const mixed = (max: number, msgs: string[], why: string): Known => ({ why, max, msgs });
/** 数で測れない差が複数ある組（今のメッセージの一覧） */
export const listed = (msgs: string[], why: string): Known => ({ why, msgs });

/** 外した組の差が、記録（上限・メッセージ）より悪化しているか */
export function worseThanKnown(d: Diff, k: Known): boolean {
  const withinMax = d.metric !== undefined && k.max !== undefined && d.metric <= k.max + 1e-9;
  const sameMsg = k.msg === d.msg || (k.msgs?.includes(d.msg) ?? false);
  return !(withinMax || sameMsg);
}

/**
 * 1 つの種類の差を、外した組の記録に照らして仕分ける。
 * - `unexpected`：記録に無い差。`worse`：記録より悪化した差。`stale`：記録にあるのに差が無くなった組（一覧から消す）
 */
export function judge(
  diffs: Diff[],
  cat: string,
  known: Record<string, Known>,
): { unexpected: string[]; worse: string[]; stale: string[] } {
  const mine = diffs.filter((d) => d.cat === cat);
  const key = (d: Diff): string => `${cat}:${d.name}`;
  return {
    unexpected: mine.filter((d) => known[key(d)] === undefined).map((d) => `${d.name}: ${d.msg}`),
    worse: mine
      .filter((d) => {
        const k = known[key(d)];
        return k !== undefined && worseThanKnown(d, k);
      })
      .map((d) => `${d.name}: ${d.msg}`),
    stale: Object.keys(known).filter(
      (k) => k.startsWith(`${cat}:`) && !mine.some((d) => key(d) === k),
    ),
  };
}
