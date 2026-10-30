// 正: InterHub docs/plan/adr/0031-xlsx-charts.md（§決定 5・P6・P7・C1・C2・C7）
/**
 * `.xlsx` のグラフの軸の計算（ISSUE-102）：数値軸の範囲・目盛間隔・目盛の値と文字、項目軸の帯・間引き。
 * - 純関数だけ（DOM・`chrome.*` に触らない）。Excel の自動範囲の近似で、係数は {@link AXIS_CONSTANTS} の 1 か所にまとめる。
 *   ISSUE-104 の Excel との照合では、この定数だけを直す（C1・C2・C7）。
 * - 目盛の数には上限を置く（明示の刻みが小さすぎても、自動でも、無限・巨大なループにならない）。
 * - ファイルの文字列は書式コード（ssf に渡す）にだけ使い、数値の文字として返す。
 */
import type { ChartAxis, ChartModel, ChartSeries } from '../core/index';
import { formatNumberByCode } from '../core/internal';

/** 軸の暫定の係数（C1・C2・C7。ISSUE-104 の照合で直す） */
export const AXIS_CONSTANTS = {
  /** C1：dmin がこの割合 × dmax 未満なら自動の最小を 0 にする（5/6） */
  zeroBaselineRatio: 5 / 6,
  /** C1：自動の範囲の端に足す余り（範囲に対する割合・5%） */
  marginRatio: 0.05,
  /**
   * C1：目盛の数（主の刻みの数）の上限 N ＝ プロットの長さ ÷ 目盛 1 つぶんの最小の間隔。
   * 間隔は縦軸＝文字の高さ × この値（Excel 16 の照合：10.5〜11.0pt の範囲＝9pt の文字で 1.0（手動の内側 110pt で 10 目盛が出る））。
   */
  labelPitchFactor: 1.0,
  /**
   * C1：横向きの軸（横棒の数値軸）の最小の間隔＝文字の高さ × この値。最長の目盛の文字の幅＋horizontalLabelGapPx がこれより大きければそちら。
   * 3 回目の照合（ISSUE-110・幅の掃引 96 個）：目盛の数の上限が 10 に届くのは、面の幅 360pt（自動配置の外側 459px）と 380pt（485px）の間
   * ＝ 9pt の文字（文字の高さ 14.4px）で最小の間隔 46.8px（3.25）。目盛の数を決める長さは、プロットの内側でなく自動配置の外側の幅
   * （layout の solveAxes の scaleW。Excel は項目軸の文字・目盛の文字の余白を引く前の幅で刻みを決める）。
   * 旧：内側の幅で 2.92（31.6pt）。外側の幅で 3.25 にすると、旧 KNOWN の C1（積み上げの負・7 桁・100% 積み上げ・長い項目名）が合う
   */
  horizontalPitchFactor: 3.25,
  /**
   * C1：横向きの軸で、最長の目盛の文字の幅に足す隙間（px）。隣の文字との隙間を 0 にしないための値。
   * 文字の幅の推定（数字 0.55em）は游ゴシック 9pt の実際より約 10% 広いので、項目軸の {@link labelPaddingPx}（4）は足さず 2 にした
   * （照合：barClustered_big_small は 6 桁の目盛で Excel の間隔が 42.6px 以下。4 を足すと 43.6px で目盛が 1 つ減って外れる）。
   */
  horizontalLabelGapPx: 2,
  /** C1：目盛の数の上限 N の下限と上限（2〜10。Excel 16 の照合：端を丸める前の刻みの数が N 以下の最小の刻み） */
  minTickLimit: 2,
  maxTickLimit: 10,
  /** C1：補助の刻み ＝ 主の刻み ÷ この値 */
  minorDivisions: 5,
  /** C2：100% 積み上げの書式（刻みは他の軸と同じく自動で決める。範囲は 0〜1、負があれば -1〜1） */
  percentFormat: '0%',
  /** 目盛の数の絶対の上限（主・補助。明示の刻みが細かすぎる時の安全弁） */
  maxMajorTicks: 100,
  maxMinorTicks: 500,
  /** C7：項目の文字の間に空ける余白（px） */
  labelPaddingPx: 4,
} as const;

/** データの範囲（最小・最大） */
export interface ValueExtent {
  min: number;
  max: number;
}

/** 数値軸の結果 */
export interface ValueScale {
  min: number;
  max: number;
  /** 主の刻み（正） */
  majorUnit: number;
  /** 補助の刻み。補助の目盛が上限を超える・明示の刻みが主の刻みより粗い時は null */
  minorUnit: number | null;
  /** 反転（`c:orientation` が `maxMin`） */
  reversed: boolean;
  /** 主の目盛の値（min から max まで。上限あり） */
  majorTicks: readonly number[];
  /** 補助の目盛の値（主の目盛と重なるものを除く。上限あり） */
  minorTicks: readonly number[];
  /** `percentStacked`（0〜1・目盛の文字は `0%`） */
  percent: boolean;
}

/** 浮動小数の誤差を丸める（12 桁） */
function round12(x: number): number {
  const r = Number(x.toPrecision(12));
  return r === 0 ? 0 : r;
}

/** 有限の数だけ通す */
function finiteOr(x: number | undefined | null): number | undefined {
  return typeof x === 'number' && Number.isFinite(x) ? x : undefined;
}

/**
 * 系列からデータの範囲を求める。点が 1 つも無ければ null。
 * - `clustered`・`standard`：全点の最小・最大。
 * - `stacked`：項目ごとの正の和の最大と負の和の最小（0 を含む）。
 * - `percentStacked`：最小は負の値が 1 つでもあれば -1・無ければ 0。最大は 1（全部が負なら 0。C2。Excel 16 の照合）。点があれば。
 */
export function dataExtent(
  grouping: ChartModel['grouping'],
  series: readonly Pick<ChartSeries, 'values'>[],
): ValueExtent | null {
  let any = false;
  const count = series.reduce((m, s) => Math.max(m, s.values.length), 0);
  if (grouping === 'stacked' || grouping === 'percentStacked') {
    let hi = 0;
    let lo = 0;
    let hasPos = false;
    const percent = grouping === 'percentStacked';
    for (let i = 0; i < count; i++) {
      let pos = 0;
      let neg = 0;
      for (const s of series) {
        const v = s.values[i];
        if (typeof v !== 'number' || !Number.isFinite(v)) continue;
        any = true;
        if (v >= 0) pos += v;
        else neg += v;
      }
      if (percent) {
        // 負の値が 1 つでもあれば最小は -1（Excel は割合を実際の値で測らない）
        if (neg < 0) lo = -1;
        if (pos > 0) hasPos = true;
      } else {
        if (pos > hi) hi = pos;
        if (neg < lo) lo = neg;
      }
    }
    if (!any) return null;
    // 全部が負の 100% 積み上げは -1〜0（C2。Excel 16 の照合）。正が 1 つでもあれば -1〜1・負が無ければ 0〜1
    return percent ? { min: lo, max: lo < 0 && !hasPos ? 0 : 1 } : { min: lo, max: hi };
  }
  let min = Infinity;
  let max = -Infinity;
  for (const s of series) {
    for (const v of s.values) {
      if (typeof v !== 'number' || !Number.isFinite(v)) continue;
      any = true;
      if (v < min) min = v;
      if (v > max) max = v;
    }
  }
  return any ? { min, max } : null;
}

/**
 * {1, 2, 5}×10ⁿ のうち、`span ÷ u ≤ maxTicks` を満たす最小の刻みを返す（P6）。
 * `span` が 0 以下・非有限なら 1。`maxTicks` は 1 以上に丸める。ループは高々数回。
 */
export function niceUnit(span: number, maxTicks: number): number {
  const s = Number.isFinite(span) && span > 0 ? span : 1;
  const n = Math.max(1, Math.floor(maxTicks));
  let exp = Math.floor(Math.log10(s / n)) - 1;
  for (let guard = 0; guard < 12; guard++) {
    for (const m of [1, 2, 5]) {
      const u = m * 10 ** exp;
      if (s / u <= n + 1e-9) return round12(u);
    }
    exp++;
  }
  return round12(10 ** exp);
}

/**
 * 目盛の数（主の刻みの数）の上限 N（C1）：プロットの長さ ÷ 目盛 1 つぶんの最小の間隔 を 2〜10 に丸める。
 * 縦軸の間隔は 文字の高さ × {@link AXIS_CONSTANTS.labelPitchFactor}。
 * 横向きの軸（横棒の数値軸）は `labelWidthPx`（最長の目盛の文字の幅）を渡すと、（幅 + {@link AXIS_CONSTANTS.horizontalLabelGapPx}）と
 * 文字の高さ × {@link AXIS_CONSTANTS.horizontalPitchFactor} の大きい方（隣の文字との隙間が 0 にならないようにする）。
 */
export function tickCountLimit(
  lengthPx: number,
  labelHeightPx: number,
  labelWidthPx?: number,
): number {
  const c = AXIS_CONSTANTS;
  const h = Math.max(1e-6, labelHeightPx);
  const pitch =
    labelWidthPx !== undefined && Number.isFinite(labelWidthPx)
      ? Math.max(h * c.horizontalPitchFactor, labelWidthPx + c.horizontalLabelGapPx)
      : h * c.labelPitchFactor;
  const raw = Number.isFinite(lengthPx) ? Math.floor(lengthPx / pitch) : c.minTickLimit;
  return Math.min(c.maxTickLimit, Math.max(c.minTickLimit, raw));
}

/** `min` から `max` まで `unit` おきの値（個数に上限。誤差は丸める） */
function ticksBetween(min: number, max: number, unit: number, cap: number): number[] {
  const out: number[] = [];
  const n = Math.min(cap, Math.floor((max - min) / unit + 1e-9));
  for (let i = 0; i <= n; i++) out.push(round12(min + i * unit));
  return out;
}

/**
 * 数値軸の範囲・刻み・目盛を決める（ADR-0031 §決定 5）。
 * 明示（`c:min`・`c:max`・`c:majorUnit`・`c:minorUnit`・`c:orientation`）があれば従い、無い値だけ自動で決める。
 * - 自動の最小：dmin ≥ 0 で dmin < dmax × 5/6 なら 0。dmax ≤ 0 側も対称に 0。それ以外は範囲の 5% の余りを取って刻みに切り下げる。
 * - 自動の最大：dmax に 5% の余りを足して刻みの倍数に切り上げる。
 * - 刻み：{@link niceUnit}（目盛の数の上限 {@link tickCountLimit}）。補助は 1/5。
 * - 点が無ければ 0〜1。`percentStacked` は 0〜1（負を含めば -1〜1）で、刻みは他の軸と同じく自動（{@link tickCountLimit}。明示があれば優先）。
 * - 目盛の数は {@link AXIS_CONSTANTS} の `maxMajorTicks`・`maxMinorTicks` で抑える。
 */
export function computeValueScale(input: {
  axis?: ChartAxis;
  extent: ValueExtent | null;
  percentStacked?: boolean;
  lengthPx: number;
  labelHeightPx: number;
  /** 横向きの軸：最長の目盛の文字の幅（px）。あれば目盛の数の上限をこの幅で決める */
  labelWidthPx?: number;
}): ValueScale {
  const c = AXIS_CONSTANTS;
  const axis = input.axis;
  const percent = input.percentStacked === true;
  const reversed = axis?.orientation === 'maxMin';
  /** 点が無い・全部 0：0〜1 に固定（自動の余りを足さない） */
  const fixed =
    percent || input.extent === null || (input.extent.min === 0 && input.extent.max === 0);
  const limitN = tickCountLimit(input.lengthPx, input.labelHeightPx, input.labelWidthPx);
  const exMin = finiteOr(axis?.min);
  const exMax = finiteOr(axis?.max);
  let dmin = input.extent?.min ?? 0;
  let dmax = input.extent?.max ?? (input.extent === null ? 1 : 0);
  if (dmin === dmax) {
    // 全部が同じ値：0 から値まで（0 なら 0〜1）
    if (dmax > 0) dmin = 0;
    else if (dmax < 0) dmax = 0;
    else dmax = 1;
  }
  const span0 = dmax - dmin;
  const margin = span0 * c.marginRatio;

  // 範囲の候補（刻みを決める前）
  let lo: number;
  let hi: number;
  let loAuto = false;
  let hiAuto = false;
  if (fixed) {
    lo = 0;
    hi = 1;
  } else {
    lo = dmin;
    hi = dmax;
  }
  if (exMin !== undefined) lo = exMin;
  else if (percent && input.extent !== null && input.extent.min < 0) {
    // 負を含む 100% 積み上げ：最小は -1（Excel 16 の照合。刻みの倍数に切り下げる）
    lo = input.extent.min;
    loAuto = true;
  } else if (!fixed) {
    if (dmin >= 0 && dmin < dmax * c.zeroBaselineRatio) lo = 0;
    else {
      lo = dmin - margin;
      loAuto = true;
    }
  }
  if (exMax !== undefined) hi = exMax;
  else if (percent && input.extent !== null) {
    // 100% 積み上げ：最大は 1（全部が負なら 0。C2）
    hi = input.extent.max;
  } else if (!fixed) {
    if (dmax <= 0 && dmax > dmin * c.zeroBaselineRatio) hi = 0;
    else {
      hi = dmax + margin;
      hiAuto = true;
    }
  }
  if (!(hi > lo)) {
    // 明示どうしが逆・同じ：解ける形に倒す
    if (exMin !== undefined && exMax === undefined) hi = lo + 1;
    else if (exMax !== undefined && exMin === undefined) lo = hi - 1;
    else {
      lo = exMin ?? 0;
      hi = lo + 1;
    }
    loAuto = false;
    hiAuto = false;
  }

  // 主の刻み
  const span = hi - lo;
  let major: number;
  const exMajor = finiteOr(axis?.majorUnit);
  if (exMajor !== undefined && exMajor > 0) major = exMajor;
  else major = niceUnit(span, limitN);
  if (span / major > c.maxMajorTicks) major = niceUnit(span, c.maxMajorTicks);

  // 自動の端は刻みの倍数に丸める
  if (loAuto) lo = round12(Math.floor(lo / major + 1e-9) * major);
  if (hiAuto) hi = round12(Math.ceil(hi / major - 1e-9) * major);
  if (!(hi > lo)) hi = round12(lo + major);

  // 補助の刻み
  let minor: number | null;
  const exMinor = finiteOr(axis?.minorUnit);
  if (exMinor !== undefined && exMinor > 0) minor = exMinor;
  else minor = round12(major / c.minorDivisions);
  if (minor !== null && (hi - lo) / minor > c.maxMinorTicks) minor = null;

  const majorTicks = ticksBetween(lo, hi, major, c.maxMajorTicks);
  let minorTicks: number[] = [];
  if (minor !== null) {
    const majorSet = new Set(majorTicks);
    minorTicks = ticksBetween(lo, hi, minor, c.maxMinorTicks).filter((t) => !majorSet.has(t));
  }
  return {
    min: round12(lo),
    max: round12(hi),
    majorUnit: major,
    minorUnit: minor,
    reversed,
    majorTicks,
    minorTicks,
    percent,
  };
}

/**
 * 値を軸上の位置（0〜1・既定は範囲外を端に収める）にする。反転（`maxMin`）なら 1 − 位置。
 * 0 は軸の始点側（縦軸なら下・横軸なら左）、反転すると逆の端。
 * `clamp` を false にすると範囲外を収めず、0 未満・1 超の位置をそのまま返す（折れ線を幾何的に切るため。非有限の値は 0）。
 */
export function valueFraction(
  scale: Pick<ValueScale, 'min' | 'max' | 'reversed'>,
  v: number,
  clamp = true,
): number {
  const span = scale.max - scale.min;
  const p = span > 0 && Number.isFinite(v) ? (v - scale.min) / span : 0;
  const q = clamp ? (p < 0 ? 0 : p > 1 ? 1 : p) : p;
  return scale.reversed ? 1 - q : q;
}

/**
 * 項目軸が数値軸と交わる値（`c:crosses`・`c:crossesAt` は項目軸の要素に書かれる＝数値軸上の位置。暫定）。
 * `crossesAt` があればそれ（範囲に収める）、`min`・`max` は端、`autoZero`（既定）は 0 が範囲内なら 0・外なら近い端。
 */
export function crossValue(scale: Pick<ValueScale, 'min' | 'max'>, catAxis?: ChartAxis): number {
  const at = finiteOr(catAxis?.crossesAt);
  const clamp = (v: number): number => Math.min(scale.max, Math.max(scale.min, v));
  if (at !== undefined) return clamp(at);
  if (catAxis?.crosses === 'min') return scale.min;
  if (catAxis?.crosses === 'max') return scale.max;
  return clamp(0);
}

/**
 * 数値軸の目盛の文字の書式コード（C2）：`percentStacked` で明示が無ければ `0%`。
 * `c:numFmt` が明示で `sourceLinked` でなければその書式。`sourceLinked` か無指定なら最初の系列の書式（無ければ General）。
 */
export function tickFormatCode(
  axis: ChartAxis | undefined,
  firstSeriesFormat: string | undefined,
  percent: boolean,
): string {
  if (axis?.numFmt !== undefined && !axis.numFmt.sourceLinked) return axis.numFmt.code;
  if (percent) return AXIS_CONSTANTS.percentFormat;
  return firstSeriesFormat ?? 'General';
}

/** 目盛の値を書式コードで文字にする（ssf。書式が壊れていても General で出す）。-0 は 0 */
export function tickLabel(code: string, value: number): string {
  return formatNumberByCode(code, value === 0 ? 0 : value, false);
}

// ---- 項目軸 ----

/**
 * 項目軸の位置（軸の始点からの距離 px。項目ごとの中心）。
 * - 既定は帯の等分：中心は (i + 0.5) × 帯の幅。
 * - `midCat`（折れ線の点が目盛の上）：項目が 2 つ以上なら 0 〜 長さを (n−1) 等分、1 つなら中央。
 * - `reversed` なら始点と終点を入れ替える（`c:orientation maxMin`・横棒の既定の向きの補正）。
 */
export function categoryPositions(
  count: number,
  lengthPx: number,
  opts: { midCat?: boolean; reversed?: boolean } = {},
): number[] {
  const n = Math.max(0, Math.floor(count));
  const out: number[] = [];
  const mid = opts.midCat === true && n >= 2;
  for (let i = 0; i < n; i++) {
    const p = mid ? (i * lengthPx) / (n - 1) : ((i + 0.5) * lengthPx) / n;
    out.push(opts.reversed === true ? lengthPx - p : p);
  }
  return out;
}

/** 項目の間隔（px）：`midCat` なら点どうしの間隔、そうでなければ帯の幅 */
export function categorySpacing(count: number, lengthPx: number, midCat = false): number {
  const n = Math.max(1, Math.floor(count));
  return midCat && n >= 2 ? lengthPx / (n - 1) : lengthPx / n;
}

/**
 * 項目の文字を何個おきに出すか（C7・P7）。
 * - `c:tickLblSkip` の明示（1 以上の整数）があればそれ。
 * - 無ければ、最長の文字の大きさ（軸の向きに沿った幅 px。横軸なら文字の幅・縦軸なら文字の高さ）＋余白が間隔に入る最小の個数。
 *   余白は既定で {@link AXIS_CONSTANTS.labelPaddingPx}。縦に回した文字の間引き（ISSUE-107）は `paddingPx` で小さい余白を渡す。
 */
export function categoryLabelSkip(
  count: number,
  spacingPx: number,
  maxLabelExtentPx: number,
  explicit?: number,
  paddingPx: number = AXIS_CONSTANTS.labelPaddingPx,
): number {
  const n = Math.max(1, Math.floor(count));
  const ex = finiteOr(explicit);
  if (ex !== undefined && ex >= 1) return Math.min(n, Math.floor(ex));
  if (!(spacingPx > 0)) return n;
  const need = maxLabelExtentPx + paddingPx;
  const k = Math.ceil(need / spacingPx - 1e-9);
  return Math.min(n, Math.max(1, k));
}

/** 文字を出す項目の位置（0, skip, 2·skip, …） */
export function categoryLabelIndices(count: number, skip: number): number[] {
  const n = Math.max(0, Math.floor(count));
  const k = Math.max(1, Math.floor(skip));
  const out: number[] = [];
  for (let i = 0; i < n; i += k) out.push(i);
  return out;
}
