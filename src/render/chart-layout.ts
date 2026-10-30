// 正: InterHub docs/plan/adr/0031-xlsx-charts.md（§決定 3〜10・P4・C3〜C10）
/**
 * `.xlsx` のグラフの配置の計算（ISSUE-102）：{@link ChartModel} とグラフの面の px の大きさから、
 * 軸の範囲・目盛・文字、既定の色、タイトル・凡例・プロットの矩形、棒・線・マーカー・扇の図形を計算し、
 * 「場面」（{@link ChartScene}）を返す純関数。
 * - DOM・canvas・`chrome.*` に触らない。文字の幅は文字の種類ごとの係数で推定する（P4）。
 * - 場面の値は数値・`#rrggbb`・不透明度・固定の語と、`text` の文字だけ。ファイル由来の文字列（系列名・項目・タイトル）は
 *   `text` 項目の `text` にしか入らない（`path` の `d` は数値と固定の命令だけ）。ISSUE-103 が属性へ文字列を入れずに済む。
 * - Excel の既定値（系列色の変化・文字の大きさ・余白・凡例の寸法・目盛の長さなど）は {@link CHART_DEFAULTS}・{@link CHART_LAYOUT}・
 *   {@link CHART_COLOR_VARIATIONS} の表にまとめる。ISSUE-104 の Excel との照合ではこの表だけを直す（C3〜C10）。
 */
import { applyTransforms, type Rgba } from '../core/internal';
import { DEFAULT_THEME, type Theme, type ThemeSlot } from '../core/internal';
import type {
  ChartAxis,
  ChartLegend,
  ChartLine,
  ChartMarkerSymbol,
  ChartModel,
  ChartPaint,
  ChartSeries,
  ChartTextStyle,
  ChartTickMark,
  LineDash,
  ManualLayout,
} from '../core/index';
import {
  categoryLabelIndices,
  categoryLabelSkip,
  categoryPositions,
  categorySpacing,
  computeValueScale,
  crossValue,
  dataExtent,
  tickFormatCode,
  tickLabel,
  valueFraction,
  type ValueScale,
} from './chart-axis';
import {
  bandGap,
  CHART_THIN,
  clusteredBarColumnPx,
  mergeBarColumns,
  splitMarkerBands,
  thinMarkerIndices,
  thinPolyline,
  thinSectorRuns,
  type BarBox,
  type BarBoxIn,
  type Pt,
} from './chart-thin';

// ---- 場面の型 ----

/** 場面の要素数の上限（円の扇・点ごとの上書き・文字の数。超えたら null＝破線枠。ADR-0031 §決定 9） */
export const MAX_CHART_SCENE_ITEMS = 5_000;

/** 色と不透明度（`#rrggbb`＋0〜1） */
export interface ChartColor {
  color: string;
  opacity: number;
}

/** 線の見た目（太さは px。端・結びは固定の語） */
export interface ChartStroke extends ChartColor {
  width: number;
  dash: LineDash;
  cap: 'butt' | 'round';
  join: 'miter' | 'round';
}

export type ChartAnchor = 'start' | 'middle' | 'end';
export type ChartBaseline = 'alphabetic' | 'middle' | 'hanging';

/** 矩形（`fill`・`stroke` が null なら塗らない・描かない） */
export interface ChartRectItem {
  kind: 'rect';
  x: number;
  y: number;
  width: number;
  height: number;
  fill: ChartColor | null;
  stroke: ChartStroke | null;
  /** 角丸の半径（px。`rx`・`ry` の数値）。グラフの面の矩形だけが持つ（`c:roundedCorners` が真の時。ISSUE-109）。無ければ直角 */
  radius?: number;
}

/** 経路（`d` は数値と固定の命令 `M L C A Z` だけ） */
export interface ChartPathItem {
  kind: 'path';
  d: string;
  fill: ChartColor | null;
  stroke: ChartStroke | null;
}

/** 線分 */
export interface ChartLineItem {
  kind: 'line';
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  stroke: ChartStroke;
}

/** 文字（`text` だけがファイル由来の文字列。回転は (x, y) を中心に度） */
export interface ChartTextItem {
  kind: 'text';
  x: number;
  y: number;
  text: string;
  size: number;
  bold: boolean;
  color: ChartColor;
  anchor: ChartAnchor;
  baseline: ChartBaseline;
  rotate: number;
}

export type ChartItem = ChartRectItem | ChartPathItem | ChartLineItem | ChartTextItem;

/** 場面：グラフの面の大きさ（px）と、描く順の要素の並び */
export interface ChartScene {
  width: number;
  height: number;
  /** グラフの面が塗りなしの時に真（ダーク表示で白い下地を敷く。ADR-0031 §決定 10） */
  backdrop: boolean;
  /** グラフの面の角丸の半径（px）。0 または無ければ直角。ダーク表示の下地（`backdrop`）の矩形にも同じ値を付ける（ISSUE-109） */
  cornerRadius?: number;
  items: ChartItem[];
}

// ---- 既定値の表（ISSUE-104 の照合で直す所） ----

/** pt → px（96/72） */
export const PX_PER_PT = 96 / 72;

/** 色の既定の指定（テーマの色の位置＋lumMod・lumOff。1/1000 %） */
interface ColorSpec {
  slot: ThemeSlot;
  lumMod?: number;
  lumOff?: number;
}

/**
 * 部品の既定（`c:spPr`・`c:txPr`・`c:style` が無い時。C4・C8）。
 * ISSUE-104：`c:spPr` 等の無い合成のグラフを Excel 16 で開いた値に合わせた（面の枠線 0.5pt・#898989、文字 10pt・黒、タイトルは太字、折れ線 1.5pt、マーカー 6pt）。
 * Excel が保存したグラフは `c:spPr`・`c:txPr` を持つのでこの表は使われない。
 */
export const CHART_DEFAULTS = {
  /** グラフの面：塗り `bg1`・枠線 #898989（黒＋輝度 53.725%）・0.5pt */
  chartAreaFill: { slot: 'lt1' } as ColorSpec,
  chartAreaLine: { slot: 'dk1', lumOff: 53725 } as ColorSpec,
  chartAreaLinePt: 0.5,
  /**
   * 主の目盛線・補助の目盛線・項目軸の線・数値軸の線（`c:spPr` の無い軸。ISSUE-105 の PNG の画素の照合）。
   * Excel 16 は主の目盛線と両方の軸の線を黒（`dk1`）・補助の目盛線を #BCBCBC（黒＋輝度 73.725%）の 1 画素の線で描く（COM の `Format.Line` の黒 0.5pt と合う）。
   * 旧値（目盛線 #d9d9d9・項目軸 #bfbfbf・数値軸なし）は、Excel が COM で作って保存した `c14:style` 付きのグラフ（#898989）とも合わなかった。
   */
  gridMajor: { slot: 'dk1' } as ColorSpec,
  gridMinor: { slot: 'dk1', lumOff: 73725 } as ColorSpec,
  catAxisLine: { slot: 'dk1' } as ColorSpec,
  valAxisLine: { slot: 'dk1' } as ColorSpec,
  /** 上の線の太さ（pt。COM の Weight。Excel は 1 画素未満の線を 1 画素に描く） */
  axisLinePt: 0.5,
  /**
   * `c14:style` を持つグラフ（Excel が COM の旧 API などで作った `c:spPr` 無しのグラフ。ISSUE-109・ISSUE-105 の COM）：
   * 主の目盛線・項目軸・数値軸の線は #898989（黒＋輝度 53.725%）0.5pt（太さは {@link axisLinePt}）、タイトルは 18pt 太字。
   * 補助の目盛線・軸と凡例の文字・面・折れ線は `c14:style` 無しと同じ。暗い背景（`c:style` 41〜48）とは独立に効く
   */
  c14Line: { slot: 'dk1', lumOff: 53725 } as ColorSpec,
  c14TitlePt: 18,
  /**
   * `c:style` 41〜48（暗い背景。ISSUE-109・ISSUE-110 の COM と PNG の値）：面の塗り 黒・面の枠線 黒 0.5pt（Excel の COM は太さ「自動」で、黒い面の上では見えない）・
   * 文字（軸・凡例・タイトル）白・プロット（棒・折れ線）#3F3F3F・円とドーナツのプロットは塗りなし（黒い面が見える）。
   * 軸線・補助の目盛線・折れ線の太さ・系列の色の組は未確定で変えない（41・43〜48 の系列の色の組・円の凡例の `#595959` 9pt はこの ISSUE の範囲外）。
   * ダーク表示でも補正しない（P5）
   */
  darkStyleMin: 41,
  darkStyleMax: 48,
  darkChartAreaFill: { rgb: 0x000000, alpha: 1 } as Rgba,
  darkChartAreaLine: { rgb: 0x000000, alpha: 1 } as Rgba,
  darkText: { rgb: 0xffffff, alpha: 1 } as Rgba,
  darkPlotFill: { rgb: 0x3f3f3f, alpha: 1 } as Rgba,
  /**
   * 暗い背景（41〜48）の棒の既定の塗り（単色近似。ISSUE-109・ADR-0031 §決定 4）。Excel の棒は上から下へのグラデーション（明→系列の色に近い色→暗）で、
   * 50% の位置の色は系列の色の彩度を 110% にした色（`a:satMod` 110%。ADR-0028 §決定 6.4 の `a:gradFill`＝50% の位置の色と同じ規則）。`linearGradient` は使わない。1/1000 %
   */
  darkBarSatMod: 110000,
  /** 文字（黒） */
  text: { slot: 'dk1' } as ColorSpec,
  /**
   * グラフの面の角丸（`c:roundedCorners` が真・要素が無い時も真。ISSUE-109）。半径は 8.25pt（11px）で、グラフの大きさによらず一定
   * （Excel 16 の PNG：216×144・360×216・720×432pt のどれも 11px。小さい面では 幅÷2・高さ÷2 を超えない）
   */
  roundedCorners: true,
  cornerRadiusPt: 8.25,
  /** 円・ドーナツの枠線 */
  pieLine: { slot: 'lt1' } as ColorSpec,
  /** 円・ドーナツの枠線の太さ（pt）。Excel は点（`c:dPt`）ごとに 1.5pt を書く（点の線は読まないので、系列に明示が無ければこの値。C9） */
  pieLinePt: 1.5,
  /** 線の太さ（pt）：マーカーの枠・線だけの記号 0.75、折れ線 1.5（c:spPr の無いグラフの値。目盛線・軸線は {@link axisLinePt}） */
  thinLinePt: 0.75,
  seriesLinePt: 1.5,
  /** マーカーの自動の形と大きさ（pt・C8） */
  autoMarker: 'circle' as ChartMarkerSymbol,
  markerSizePt: 6,
  /** 文字の大きさ（pt・C4）：タイトル 10（太字）・軸と凡例 10・軸タイトル 10 */
  titlePt: 10,
  titleBold: true,
  axisPt: 10,
  legendPt: 10,
  axisTitlePt: 10,
  /** ドーナツの穴（%・C9） */
  holeSizePercent: 50,
  /** 棒の隙間（%）・重なり（集合 0・積み上げ 100） */
  gapWidth: 150,
  overlapClustered: 0,
  overlapStacked: 100,
  /** 円の開始角（度・12 時から時計回り） */
  firstSliceAng: 0,
  /** 項目軸の文字の回転の既定（`a:bodyPr rot` が無く、自動の形の判定を使わない時＝横棒の項目軸・小さい面。P7・ISSUE-107：回さない） */
  autoRotation: 0,
  /** `c:majorTickMark`・`c:minorTickMark` の要素が無い軸の目盛の印（スキーマの既定。Excel 16 の COM は 4 つとも Cross を返す・C12） */
  tickMark: 'cross' as ChartTickMark,
} as const;

/**
 * 配置の寸法（C5・C6・C10。pt は × 96/72、`Ratio` は文字の大きさ（px）に対する割合）。
 * ISSUE-104 で Excel 16 の COM の値（座標は面の左上から 4pt 内側が原点）に合わせた：
 * タイトルの上端は面の上から 6pt・プロット（軸の文字を除く内側）の上と右はタイトル／凡例の下・面の右（凡例の左）から 11pt、
 * 軸の文字の左と下は面の左・下（凡例の右・上）から 5pt、凡例は面の縁から 6pt。円・ドーナツは四方 11pt。
 */
export const CHART_LAYOUT = {
  /** 軸の文字の左・下の余白（pt。面の左端・下端、凡例の右・上から） */
  marginPt: 5,
  /** プロットの上・右の余白（pt。タイトル・凡例の下／面の右・凡例の左から）。円・ドーナツは四方 */
  plotPadPt: 11,
  /** タイトルの上端（pt） */
  titleTopPt: 6,
  /** 文字の高さ ＝ 大きさ × この値 */
  lineHeightRatio: 1.2,
  /** タイトルの高さ ＝ 大きさ × この値（Excel の Calibri 14pt で 1.435・10pt 太字で 1.52） */
  titleLineHeightRatio: 1.45,
  /** 目盛の長さ（主・補助。pt）と、目盛の文字との間隔（px） */
  tickLengthPt: 5,
  minorTickLengthPt: 3,
  labelGapPx: 3,
  /** 軸タイトルの帯の余白（px） */
  axisTitleGapPx: 4,
  /** 凡例 */
  legendSymbolRatio: 0.55,
  legendLineSymbolRatio: 1.33,
  legendGapRatio: 0.4,
  /**
   * 横並び（1 行）で項目と項目の間（C6）。Excel の項目の間隔は、記号＋文字の幅（実測）+ 約 0.47〜0.49em。
   * 文字の幅の推定が広めに出る名前（円 5 項目の凡例）があるので、1 回目・2 回目の照合の両方が 3% 以内に収まる 0.4 にした
   */
  legendEntryGapRatio: 0.4,
  /** 横並びの格子の桁の余白（最も長い項目の幅に足す。Excel の桁の幅は文字の幅の推定より少し広い。C6） */
  legendColPadRatio: 0.1,
  legendRowGapRatio: 0.675,
  /** 凡例の箱の余白（左右・上下） */
  legendPadRatio: 0.34,
  legendPadYRatio: 0,
  legendMaxFraction: 0.4,
  /** 右上の凡例でプロットを凡例の下まで縮める条件：縮めた後のプロットの高さが縮める前のこの割合以上（Excel 16 の COM：小さい面は 0.50・標準は 0.71） */
  legendCornerShrinkMin: 0.6,
  /** 凡例と面の縁（タイトルの下）の間（pt） */
  legendEdgeGapPt: 6,
  /** 回転した文字の本文の最大の長さの打ち切り（1 行の文字数） */
  maxLineChars: 255,
  /**
   * 項目軸の文字の形の自動判定（縦棒・折れ線の横の項目軸。ISSUE-107・ADR-0031 §決定 5。Excel 16 の c7sw_*・c7w_* の基準値から）。
   * 水平 → 折り返し → 45° → 90° の順に最初に当たった形にする。値の「暫定」は基準値の無い境目（外れた組は照合テストの KNOWN）。
   * - `catRotateMinSlotRatio`：45° と 90° の境目（項目の間隔 ≧ 文字の大きさ × この値なら 45°）。観測は 90° の最大 10.7pt・45° の最小 14.4pt（9pt 文字で 1.6 倍）。暫定
   * - `catRotatedBandExtraPt`：回した文字の帯の厚さ ＝ (幅 ＋ 行の高さ)·sin45° ＋ この値（pt。目盛・文字との間隔を含む。実測 40・50・75pt と ±1pt）。90° も暫定で同じ式
   * - `catWrapRowRatio`：折り返した行の送り ＝ 文字の大きさ × この値（9pt 文字で 11pt。帯は 1 行 18pt・2 行 29pt・3 行 40pt）
   * - `catRotatedSkipPadPt`：90° に回した文字の間引きの余白（pt）。4px のままだと 1 回目の cat30_small が 3 個おきになる（Excel は 2 個おき）
   * - `catReachGapPt`：回した文字が左へ出る分の下限 ＝ (幅 ＋ 行の高さ)·cos45° − 最初の項目の中心の位置 ＋ この値（pt。面の左端から見たプロットの左）
   * - `catBandMaxRatio`：推定の帯が面の高さのこの割合を超える小さい面は、回さず今までどおり間引く（安全弁。暫定。1 回目の catlong_small・catlong_std は推定の幅が広めに出て戻る。0.45 にすると std は回るが、プロットの高さが Excel より 17px 低く数値軸の刻みが変わって 1 回目の KNOWN が増える）
   */
  catRotateMinSlotRatio: 1.6,
  catRotatedBandExtraPt: 6.5,
  catWrapRowRatio: 11 / 9,
  catRotatedSkipPadPt: 1,
  catReachGapPt: 2,
  catBandMaxRatio: 0.4,
} as const;

/** 文字の幅の係数（P4・暫定）：半角の英数字と記号・半角カナ・全角・その他（em に対する割合） */
export const TEXT_WIDTH_EM = { ascii: 0.55, halfKana: 0.5, full: 1, other: 0.6 } as const;

/** 既定の色の変化（C3）：周回 k（9 周で繰り返す）の lumMod・lumOff（1/1000 %）。k=0 は変化なし */
export const CHART_COLOR_VARIATIONS: readonly { lumMod?: number; lumOff?: number }[] = [
  {},
  { lumMod: 60000 },
  { lumMod: 80000, lumOff: 20000 },
  { lumMod: 80000 },
  { lumMod: 60000, lumOff: 40000 },
  { lumMod: 50000 },
  { lumMod: 70000, lumOff: 30000 },
  { lumMod: 70000 },
  { lumMod: 50000, lumOff: 50000 },
];

const ACCENT_SLOTS: readonly ThemeSlot[] = [
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
];

// ---- 色 ----

/** 変換の子の列（`color.ts` の `applyTransforms` に渡す本文） */
function transformBody(lumMod?: number, lumOff?: number): string | null {
  let s = '';
  if (lumMod !== undefined) s += `<a:lumMod val="${lumMod}"/>`;
  if (lumOff !== undefined) s += `<a:lumOff val="${lumOff}"/>`;
  return s === '' ? null : s;
}

/** テーマの色（無ければ既定のテーマ） */
function themeRgb(theme: Theme, slot: ThemeSlot): number {
  return theme.get(slot) ?? DEFAULT_THEME.get(slot) ?? 0;
}

/** 既定の色の指定を色にする */
function specColor(theme: Theme, spec: ColorSpec): Rgba {
  return applyTransforms(
    { rgb: themeRgb(theme, spec.slot), alpha: 1 },
    transformBody(spec.lumMod, spec.lumOff),
  );
}

/**
 * 既定の系列（・点）の色（C3）：位置 `index`（系列は `c:idx`・点は点の位置）に対し、テーマの `accent((index mod 6)+1)` に
 * 周回 k＝floor(index / 6) の変化（{@link CHART_COLOR_VARIATIONS}。9 周で繰り返す）を掛ける。
 */
export function autoColor(theme: Theme, index: number): Rgba {
  const i = Number.isFinite(index) && index > 0 ? Math.floor(index) : 0;
  const base: Rgba = { rgb: themeRgb(theme, ACCENT_SLOTS[i % 6]!), alpha: 1 };
  const v = CHART_COLOR_VARIATIONS[Math.floor(i / 6) % CHART_COLOR_VARIATIONS.length]!;
  return applyTransforms(base, transformBody(v.lumMod, v.lumOff));
}

const hex = (rgb: number): string => `#${(rgb & 0xffffff).toString(16).padStart(6, '0')}`;
const clamp01 = (x: number): number => (x > 0 ? (x < 1 ? x : 1) : 0);
const toColor = (c: Rgba): ChartColor => ({ color: hex(c.rgb), opacity: clamp01(c.alpha) });

// ---- 文字の幅・切り詰め ----

/** 1 文字の幅（em）。半角の英数字と記号 0.55・半角カナ 0.5・全角 1・その他 0.6（P4・暫定） */
function charEm(cp: number): number {
  if (cp >= 0x20 && cp <= 0x7e) return TEXT_WIDTH_EM.ascii;
  if (cp >= 0xff61 && cp <= 0xff9f) return TEXT_WIDTH_EM.halfKana;
  if (
    (cp >= 0x1100 && cp <= 0x115f) ||
    (cp >= 0x2e80 && cp <= 0xa4cf) ||
    (cp >= 0xac00 && cp <= 0xd7a3) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xfe30 && cp <= 0xfe6f) ||
    (cp >= 0xff01 && cp <= 0xff60) ||
    (cp >= 0xffe0 && cp <= 0xffe6) ||
    (cp >= 0x20000 && cp <= 0x3fffd)
  )
    return TEXT_WIDTH_EM.full;
  return TEXT_WIDTH_EM.other;
}

/** 文字の幅の推定（px）。`sizePx` は文字の大きさ（px） */
export function estimateTextWidth(text: string, sizePx: number): number {
  let em = 0;
  for (const ch of text) em += charEm(ch.codePointAt(0)!);
  return em * sizePx;
}

/** 1 行を最大 `max` 文字（既定 255）に切る（超えたら末尾を `…` にする） */
export function truncateLine(text: string, max: number = CHART_LAYOUT.maxLineChars): string {
  if (text.length <= max) return text;
  const cps = Array.from(text);
  return cps.length <= max ? text : `${cps.slice(0, max - 1).join('')}…`;
}

/** 推定の幅が `maxWidth` に入るように末尾を `…` で切る */
function fitText(text: string, sizePx: number, maxWidth: number): string {
  if (estimateTextWidth(text, sizePx) <= maxWidth) return text;
  const cps = Array.from(text);
  let lo = 0;
  let hi = cps.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (estimateTextWidth(`${cps.slice(0, mid).join('')}…`, sizePx) <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return lo === 0 ? '' : `${cps.slice(0, lo).join('')}…`;
}

// ---- 数値の文字列 ----

/** 経路の数値：有限に直し、小数 2 桁に丸める（指数表記にしない） */
function f(x: number): string {
  const v = Number.isFinite(x) ? Math.max(-1e7, Math.min(1e7, x)) : 0;
  const r = Math.round(v * 100) / 100;
  return String(r === 0 ? 0 : r);
}

/** 有限でなければ 0 にする（場面の数値は常に有限） */
const fin = (x: number): number => (Number.isFinite(x) ? x : 0);

// ---- 矩形 ----

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** manualLayout を矩形に当てる（`edge`＝グラフの面に対する位置の割合・`factor`＝自動の位置からのずれ。ADR-0031 §決定 6） */
function applyManual(l: ManualLayout, W: number, H: number, auto: Rect): Rect {
  const x = l.x === undefined ? auto.x : l.xMode === 'edge' ? l.x * W : auto.x + l.x * W;
  const y = l.y === undefined ? auto.y : l.yMode === 'edge' ? l.y * H : auto.y + l.y * H;
  const w = l.w === undefined ? auto.w : l.wMode === 'edge' ? l.w * W - x : l.w * W;
  const h = l.h === undefined ? auto.h : l.hMode === 'edge' ? l.h * H - y : l.h * H;
  return { x: fin(x), y: fin(y), w: Math.max(1, fin(w)), h: Math.max(1, fin(h)) };
}

const hasManual = (l: ManualLayout | undefined): l is ManualLayout =>
  l !== undefined &&
  (l.x !== undefined || l.y !== undefined || l.w !== undefined || l.h !== undefined);

// ---- 場面の構築 ----

/** 要素数の上限を超えた（内部で投げ、layoutChart が null にする） */
class SceneOverflow extends Error {}

/** 解決した文字の見た目 */
interface TextLook {
  px: number;
  bold: boolean;
  color: ChartColor;
}

/** 描画の文脈 */
interface Ctx {
  model: ChartModel;
  theme: Theme;
  W: number;
  H: number;
  items: ChartItem[];
  /** 点の多いグラフの幾何を間引くか（ADR-0031 §決定 11.1。既定は真。密でないグラフでは何も変わらない） */
  thin: boolean;
  /** グラフの面の角丸の半径（px。0＝直角。ISSUE-109） */
  radius: number;
}

/** `c:style` が 41〜48（暗い背景）か（ISSUE-109） */
function isDarkStyle(m: ChartModel): boolean {
  return (
    m.style !== undefined &&
    m.style >= CHART_DEFAULTS.darkStyleMin &&
    m.style <= CHART_DEFAULTS.darkStyleMax
  );
}

/** 目盛線・軸線の既定の色：`c14:style` を持てば #898989、無ければ `normal`（ISSUE-109・CHART_DEFAULTS.c14Line） */
function lineSpec(cx: Ctx, normal: ColorSpec): ColorSpec {
  return cx.model.c14Style !== undefined ? CHART_DEFAULTS.c14Line : normal;
}

function push(cx: Ctx, item: ChartItem): void {
  if (cx.items.length >= MAX_CHART_SCENE_ITEMS) throw new SceneOverflow();
  cx.items.push(item);
}

/** 文字の見た目（部品 → グラフ全体 → 既定の順・§決定 8） */
function textLook(
  cx: Ctx,
  part: ChartTextStyle | undefined,
  defaultPt: number,
  defaultBold = false,
): TextLook {
  const pt = part?.sizePt ?? cx.model.text.sizePt ?? defaultPt;
  const color =
    part?.color ??
    cx.model.text.color ??
    (isDarkStyle(cx.model) ? CHART_DEFAULTS.darkText : specColor(cx.theme, CHART_DEFAULTS.text));
  return {
    px: pt * PX_PER_PT,
    bold: part?.bold ?? cx.model.text.bold ?? defaultBold,
    color: toColor(color),
  };
}

/** 線を解く（明示の `a:ln` ＞ 既定。`none` と既定が無ければ null） */
function resolveStroke(
  cx: Ctx,
  line: ChartLine | undefined,
  def: { spec: ColorSpec | Rgba; widthPt: number } | null,
  caps: { cap?: 'butt' | 'round'; join?: 'miter' | 'round' } = {},
): ChartStroke | null {
  if (line?.paint?.kind === 'none') return null;
  let rgba: Rgba | null = null;
  if (line?.paint?.kind === 'solid') rgba = line.paint.color;
  else if (def !== null) rgba = 'slot' in def.spec ? specColor(cx.theme, def.spec) : def.spec;
  if (rgba === null) return null;
  const width =
    line?.widthEmu !== undefined ? line.widthEmu / 9525 : (def?.widthPt ?? 0.75) * PX_PER_PT;
  return {
    ...toColor(rgba),
    width: Math.max(0, fin(width)),
    dash: line?.dash ?? 'solid',
    cap: caps.cap ?? 'butt',
    join: caps.join ?? 'miter',
  };
}

/** 塗りを解く（`none` は null・明示が無ければ `def`） */
function resolveFill(paint: ChartPaint | undefined, def: ChartColor | null): ChartColor | null {
  if (paint === undefined) return def;
  return paint.kind === 'none' ? null : toColor(paint.color);
}

function pushText(
  cx: Ctx,
  x: number,
  y: number,
  text: string,
  look: TextLook,
  anchor: ChartAnchor,
  baseline: ChartBaseline,
  rotate = 0,
): void {
  push(cx, {
    kind: 'text',
    x: fin(x),
    y: fin(y),
    text,
    size: look.px,
    bold: look.bold,
    color: look.color,
    anchor,
    baseline,
    rotate: fin(rotate),
  });
}

// ---- マーカー ----

/** マーカー 1 つの経路（中心 (x, y)・半径 r）。`filled` が偽の形（線だけ）は塗らない */
function markerPath(
  sym: ChartMarkerSymbol,
  x: number,
  y: number,
  r: number,
): { d: string; filled: boolean } {
  const a = f(x - r);
  const b = f(x + r);
  const t = f(y - r);
  const u = f(y + r);
  switch (sym) {
    case 'square':
      return { d: `M${a} ${t}L${b} ${t}L${b} ${u}L${a} ${u}Z`, filled: true };
    case 'diamond':
      return { d: `M${f(x)} ${t}L${b} ${f(y)}L${f(x)} ${u}L${a} ${f(y)}Z`, filled: true };
    case 'triangle':
      return { d: `M${f(x)} ${t}L${b} ${u}L${a} ${u}Z`, filled: true };
    case 'x':
      return { d: `M${a} ${t}L${b} ${u}M${b} ${t}L${a} ${u}`, filled: false };
    case 'plus':
      return { d: `M${a} ${f(y)}L${b} ${f(y)}M${f(x)} ${t}L${f(x)} ${u}`, filled: false };
    case 'star':
      return {
        d: `M${a} ${t}L${b} ${u}M${b} ${t}L${a} ${u}M${a} ${f(y)}L${b} ${f(y)}M${f(x)} ${t}L${f(x)} ${u}`,
        filled: false,
      };
    case 'dash':
      return { d: `M${a} ${f(y)}L${b} ${f(y)}`, filled: false };
    case 'dot': {
      const h = r * 0.5;
      return {
        d: `M${f(x - h)} ${f(y - h)}L${f(x + h)} ${f(y - h)}L${f(x + h)} ${f(y + h)}L${f(x - h)} ${f(y + h)}Z`,
        filled: true,
      };
    }
    default:
      // circle・auto・picture
      return {
        d: `M${a} ${f(y)}A${f(r)} ${f(r)} 0 1 1 ${b} ${f(y)}A${f(r)} ${f(r)} 0 1 1 ${a} ${f(y)}Z`,
        filled: true,
      };
  }
}

/**
 * 系列のマーカーの形（`none` の時だけ null）。自動・無指定は {@link CHART_DEFAULTS.autoMarker}（C8・C13）。
 * グループの `c:marker`（`lineMarkers`）は見ない：Excel 16 は、系列に `c:marker` が無ければ、グループの `c:marker` が 0 でも無くても、自動のマーカーを描く（COM の MarkerStyle が自動）。
 */
function markerSymbol(s: ChartSeries): ChartMarkerSymbol | null {
  const sym = s.marker?.symbol;
  if (sym === 'none') return null;
  if (sym !== undefined && sym !== 'auto') return sym;
  return CHART_DEFAULTS.autoMarker;
}

// ---- 系列・点の色 ----

/** 棒・折れ線で `varyColors` が点の位置の色になるか（1 系列の時だけ・C3） */
const varyByPoint = (model: ChartModel): boolean => model.varyColors && model.series.length === 1;

/** 系列の既定の色（明示が無い時） */
const seriesAuto = (cx: Ctx, s: ChartSeries): ChartColor => toColor(autoColor(cx.theme, s.idx));

/**
 * 点 i の塗り：`c:dPt` ＞ 系列の明示の塗り ＞ 既定（`varyColors` の効く時は点の位置・そうでなければ系列の `idx`）。
 * `none` は null（描かない）。
 */
function pointFill(cx: Ctx, s: ChartSeries, i: number, vary: boolean): ChartColor | null {
  const dpt = s.points?.get(i);
  if (dpt !== undefined) return resolveFill(dpt, null);
  if (s.fill !== undefined) return resolveFill(s.fill, null);
  const auto = vary ? autoColor(cx.theme, i) : autoColor(cx.theme, s.idx);
  // 暗い背景（41〜48）の棒：Excel のグラデーションの 50% の位置の色に近い単色（彩度 110%。ISSUE-109）
  if (cx.model.type === 'bar' && isDarkStyle(cx.model))
    return toColor(applyTransforms(auto, `<a:satMod val="${CHART_DEFAULTS.darkBarSatMod}"/>`));
  return toColor(auto);
}

// ---- タイトル・凡例 ----

/**
 * グラフのタイトルの行（無ければ空）。文字の無い `c:title` は系列が 1 つなら系列名（C10）。
 * `c:title` が無くても `c:autoTitleDeleted` が偽で系列が 1 つなら、Excel は系列名を自動のタイトルに出す（Excel 16 の照合）。
 */
function titleLines(model: ChartModel): string[] {
  const t = model.title;
  if (t === undefined) {
    const only = model.series.length === 1 ? model.series[0]!.name : '';
    return model.autoTitleDeleted || only === '' ? [] : [truncateLine(only)];
  }
  if (t.lines.length > 0) return t.lines.map((l) => truncateLine(l));
  const only = model.series.length === 1 ? model.series[0]!.name : '';
  return only === '' ? [] : [truncateLine(only)];
}

/**
 * 文字の無い `c:title`（自動）で系列が 2 つ以上か。Excel は言語ごとの代替の文字（日本語 UI では「グラフ タイトル」）を出して帯を取る（C10）。
 * 代替の文字は言語に依るので描かず、帯（1 行ぶんの高さ）だけ取る。
 */
function isBlankAutoTitle(model: ChartModel): boolean {
  const t = model.title;
  return t !== undefined && t.lines.length === 0 && model.series.length >= 2;
}

/** 凡例の項目 */
interface LegendEntry {
  text: string;
  /** 記号：四角（棒・円）・線（折れ線） */
  symbol: 'square' | 'line';
  fill: ChartColor | null;
  stroke: ChartStroke | null;
  marker: {
    sym: ChartMarkerSymbol;
    fill: ChartColor | null;
    stroke: ChartStroke | null;
    sizePx: number;
  } | null;
}

/** 凡例の項目を作る（系列名／`varyColors` の項目・`c:legendEntry` の削除を除く。並びは系列の順・C6） */
function legendEntries(cx: Ctx, legend: ChartLegend): LegendEntry[] {
  const m = cx.model;
  const out: LegendEntry[] = [];
  const del = new Set(legend.deleted);
  const raw: LegendEntry[] = [];
  // 円は varyColors に関わらず項目（点）ごと（偽なら全部が系列の色。Excel 16 の COM は 5 項目を返す・C3）。
  // ドーナツは varyColors が真の時だけ項目ごと（偽なら系列ごと＝ 1 系列でも系列名 1 つ）
  const byPoint = m.type === 'pie' ? true : m.type === 'doughnut' ? m.varyColors : varyByPoint(m);
  if (byPoint) {
    const s = m.series[0];
    if (s === undefined) return out;
    m.categories.forEach((name, i) => {
      raw.push(entryFor(cx, m.type, s, truncateLine(name), pointFill(cx, s, i, m.varyColors)));
    });
  } else {
    m.series.forEach((s, pos) => {
      const name = s.name === '' ? `Series ${pos + 1}` : truncateLine(s.name);
      raw.push(entryFor(cx, m.type, s, name, pointFill(cx, s, 0, false)));
    });
  }
  // 縦並びの凡例（右・左・右上）は、積み上げの縦棒・集合の横棒で項目の並びを逆にする（積み・棒の見た目の並びと合わせる。
  // 積み上げの横棒・集合の縦棒・折れ線は逆にしない。Excel 16 の COM で確認・C11）。`c:legendEntry` の idx は表示の並びの位置
  const stacked = m.grouping === 'stacked' || m.grouping === 'percentStacked';
  const verticalLegend = legend.pos === 'r' || legend.pos === 'l' || legend.pos === 'tr';
  if (m.type === 'bar' && verticalLegend && (m.barDir === 'bar' ? !stacked : stacked))
    raw.reverse();
  raw.forEach((e, i) => {
    if (!del.has(i)) out.push(e);
  });
  return out;
}

/** 凡例の項目 1 つ（棒・円は四角・折れ線は線＋マーカー） */
function entryFor(
  cx: Ctx,
  type: ChartModel['type'],
  s: ChartSeries,
  text: string,
  fill: ChartColor | null,
): LegendEntry {
  if (type !== 'line') {
    const stroke =
      type === 'bar'
        ? resolveStroke(cx, s.line, null)
        : resolveStroke(cx, s.line, {
            spec: CHART_DEFAULTS.pieLine,
            widthPt: CHART_DEFAULTS.pieLinePt,
          });
    return { text, symbol: 'square', fill, stroke, marker: null };
  }
  const lineColor = s.line?.paint?.kind === 'solid' ? toColor(s.line.paint.color) : fill;
  const stroke = resolveStroke(
    cx,
    s.line,
    { spec: autoColor(cx.theme, s.idx), widthPt: CHART_DEFAULTS.seriesLinePt },
    { cap: 'round', join: 'round' },
  );
  const sym = markerSymbol(s);
  let marker: LegendEntry['marker'] = null;
  if (sym !== null) {
    const mfill = resolveFill(s.marker?.fill, lineColor ?? fill);
    marker = {
      sym,
      fill: mfill,
      stroke: resolveStroke(cx, s.marker?.line, {
        spec: autoColor(cx.theme, s.idx),
        widthPt: CHART_DEFAULTS.thinLinePt,
      }),
      sizePx: (s.marker?.size ?? CHART_DEFAULTS.markerSizePt) * PX_PER_PT,
    };
  }
  return { text, symbol: 'line', fill, stroke, marker };
}

interface PlacedEntry {
  entry: LegendEntry;
  /** 記号の左端・行の中心 y（凡例の左上が原点） */
  x: number;
  y: number;
  symW: number;
  text: string;
}

interface LegendBox {
  w: number;
  h: number;
  placed: PlacedEntry[];
}

/**
 * 凡例の項目を箱の中に並べる（右・左・右上は縦、上・下は横に並べて折り返す。入らない項目は出さない・§決定 7）。
 * 項目が 1 つも入らなければ null。
 */
function buildLegend(
  entries: readonly LegendEntry[],
  pos: ChartLegend['pos'],
  look: TextLook,
  maxW: number,
  maxH: number,
): LegendBox | null {
  const L = CHART_LAYOUT;
  const s = look.px;
  const pad = s * L.legendPadRatio;
  const padY = s * L.legendPadYRatio;
  const gap = s * L.legendGapRatio;
  const rowH = s * L.lineHeightRatio + s * L.legendRowGapRatio;
  const maxRows = Math.floor((maxH - 2 * padY) / rowH);
  if (maxRows < 1 || entries.length === 0) return null;
  const vertical = pos === 'r' || pos === 'l' || pos === 'tr';
  const symOf = (e: LegendEntry): number =>
    s * (e.symbol === 'line' ? L.legendLineSymbolRatio : L.legendSymbolRatio);
  const placed: PlacedEntry[] = [];
  if (vertical) {
    const innerW = Math.max(0, maxW - 2 * pad);
    let widest = 0;
    const shown = entries.slice(0, maxRows);
    shown.forEach((e, k) => {
      const symW = symOf(e);
      const text = fitText(e.text, s, Math.max(0, innerW - symW - gap));
      const w = symW + gap + estimateTextWidth(text, s);
      widest = Math.max(widest, w);
      placed.push({ entry: e, x: pad, y: padY + k * rowH + rowH / 2, symW, text });
    });
    return { w: widest + 2 * pad, h: shown.length * rowH + 2 * padY, placed };
  }
  const innerW = Math.max(0, maxW - 2 * pad);
  const gapX = s * L.legendEntryGapRatio;
  // 項目ごとの幅（記号＋間隔＋文字）
  const cells = entries.map((e) => {
    const symW = symOf(e);
    const text = fitText(e.text, s, Math.max(0, innerW - symW - gap));
    return { entry: e, symW, text, w: symW + gap + estimateTextWidth(text, s) };
  });
  const flowW = cells.reduce((a, c) => a + c.w, 0) + (cells.length - 1) * gapX;
  if (flowW <= innerW) {
    // 1 行に収まる：項目ごとの幅で左から並べる（行は箱の中央）
    let x = pad;
    for (const c of cells) {
      placed.push({ entry: c.entry, x, y: padY + rowH / 2, symW: c.symW, text: c.text });
      x += c.w + gapX;
    }
    return { w: flowW + 2 * pad, h: rowH + 2 * padY, placed };
  }
  // 1 行に収まらない：等幅の桁の格子（桁の幅＝最も長い項目の幅）。Excel は 1 行の桁数の上限から行数を決め、
  // 行数のまま桁数を均す（13 項目・上限 8 桁 → 2 行・7 桁。C6）。行の上限を超える項目は出さない
  const colW = cells.reduce((m, c) => Math.max(m, c.w), 0) + s * L.legendColPadRatio;
  const maxCols = Math.max(1, Math.floor(innerW / colW));
  let rowCount = Math.ceil(cells.length / maxCols);
  let shown = cells.length;
  if (rowCount > maxRows) {
    rowCount = maxRows;
    shown = rowCount * maxCols;
  }
  const cols = Math.ceil(Math.min(shown, cells.length) / rowCount);
  cells.slice(0, shown).forEach((c, k) => {
    placed.push({
      entry: c.entry,
      x: pad + (k % cols) * colW,
      y: padY + Math.floor(k / cols) * rowH + rowH / 2,
      symW: c.symW,
      text: c.text,
    });
  });
  return { w: cols * colW + 2 * pad, h: rowCount * rowH + 2 * padY, placed };
}

/** 凡例を場面へ出す（箱の左上 (ox, oy)） */
function emitLegend(cx: Ctx, box: LegendBox, look: TextLook, ox: number, oy: number): void {
  const s = look.px;
  const gap = s * CHART_LAYOUT.legendGapRatio;
  for (const p of box.placed) {
    const e = p.entry;
    const x = ox + p.x;
    const y = oy + p.y;
    if (e.symbol === 'square') {
      const q = s * CHART_LAYOUT.legendSymbolRatio;
      if (e.fill !== null || e.stroke !== null) {
        push(cx, {
          kind: 'rect',
          x: fin(x),
          y: fin(y - q / 2),
          width: q,
          height: q,
          fill: e.fill,
          stroke: e.stroke,
        });
      }
    } else {
      if (e.stroke !== null) {
        push(cx, {
          kind: 'line',
          x1: fin(x),
          y1: fin(y),
          x2: fin(x + p.symW),
          y2: fin(y),
          stroke: e.stroke,
        });
      }
      if (e.marker !== null) {
        const m = markerPath(
          e.marker.sym,
          x + p.symW / 2,
          y,
          Math.min(e.marker.sizePx, s * 1.2) / 2,
        );
        push(cx, {
          kind: 'path',
          d: m.d,
          fill: m.filled ? e.marker.fill : null,
          stroke:
            e.marker.stroke ?? (m.filled ? null : e.fill === null ? null : strokeFrom(e.fill)),
        });
      }
    }
    pushText(cx, x + p.symW + gap, y, p.text, look, 'start', 'middle');
  }
}

/** 色から細い線（線だけの記号の既定） */
function strokeFrom(c: ChartColor): ChartStroke {
  return {
    ...c,
    width: CHART_DEFAULTS.thinLinePt * PX_PER_PT,
    dash: 'solid',
    cap: 'butt',
    join: 'miter',
  };
}

// ---- 円・ドーナツの経路 ----

const TAU = Math.PI * 2;

/** 12 時から時計回りの角度 θ（rad）の円周上の点 */
function polar(cx0: number, cy0: number, r: number, theta: number): [number, number] {
  return [cx0 + r * Math.sin(theta), cy0 - r * Math.cos(theta)];
}

function circlePath(cx0: number, cy0: number, r: number, sweep: 0 | 1): string {
  return `M${f(cx0 - r)} ${f(cy0)}A${f(r)} ${f(r)} 0 1 ${sweep} ${f(cx0 + r)} ${f(cy0)}A${f(r)} ${f(r)} 0 1 ${sweep} ${f(cx0 - r)} ${f(cy0)}Z`;
}

/** 扇（`a0` から `a1` へ時計回り）。全周なら円 */
function sectorPath(cx0: number, cy0: number, r: number, a0: number, a1: number): string {
  if (a1 - a0 >= TAU - 1e-6) return circlePath(cx0, cy0, r, 1);
  const [x0, y0] = polar(cx0, cy0, r, a0);
  const [x1, y1] = polar(cx0, cy0, r, a1);
  return `M${f(cx0)} ${f(cy0)}L${f(x0)} ${f(y0)}A${f(r)} ${f(r)} 0 ${a1 - a0 > Math.PI ? 1 : 0} 1 ${f(x1)} ${f(y1)}Z`;
}

/** 輪の扇（外半径 ro・内半径 ri）。全周なら外円と逆回りの内円（穴） */
function ringPath(
  cx0: number,
  cy0: number,
  ro: number,
  ri: number,
  a0: number,
  a1: number,
): string {
  if (a1 - a0 >= TAU - 1e-6) return circlePath(cx0, cy0, ro, 1) + circlePath(cx0, cy0, ri, 0);
  const [ox0, oy0] = polar(cx0, cy0, ro, a0);
  const [ox1, oy1] = polar(cx0, cy0, ro, a1);
  const [ix0, iy0] = polar(cx0, cy0, ri, a0);
  const [ix1, iy1] = polar(cx0, cy0, ri, a1);
  const large = a1 - a0 > Math.PI ? 1 : 0;
  return `M${f(ox0)} ${f(oy0)}A${f(ro)} ${f(ro)} 0 ${large} 1 ${f(ox1)} ${f(oy1)}L${f(ix1)} ${f(iy1)}A${f(ri)} ${f(ri)} 0 ${large} 0 ${f(ix0)} ${f(iy0)}Z`;
}

/** 円・ドーナツの扇を場面へ出す（円は最初の系列だけ・ドーナツは系列ごとの同心の輪で最初の系列が内側・C9） */
function emitPie(cx: Ctx, plot: Rect): void {
  const m = cx.model;
  const cxp = plot.x + plot.w / 2;
  const cyp = plot.y + plot.h / 2;
  const R = Math.max(0, Math.min(plot.w, plot.h) / 2);
  const start = ((m.firstSliceAng ?? CHART_DEFAULTS.firstSliceAng) * Math.PI) / 180;
  const rings = m.type === 'doughnut' ? m.series : m.series.slice(0, 1);
  const hole =
    m.type === 'doughnut' ? (R * (m.holeSize ?? CHART_DEFAULTS.holeSizePercent)) / 100 : 0;
  const ringW = (R - hole) / Math.max(1, rings.length);
  rings.forEach((s, j) => {
    const abs = s.values.map((v) =>
      typeof v === 'number' && Number.isFinite(v) ? Math.abs(v) : 0,
    );
    const total = abs.reduce((a, b) => a + b, 0);
    if (!(total > 0)) return;
    // Excel は最初の系列が内側（中心寄り）の輪
    const ro = m.type === 'doughnut' ? hole + (j + 1) * ringW : R;
    const ri = m.type === 'doughnut' ? ro - ringW : 0;
    const stroke = resolveStroke(cx, s.line, {
      spec: CHART_DEFAULTS.pieLine,
      widthPt: CHART_DEFAULTS.pieLinePt,
    });
    const wedges: { i: number; a0: number; a1: number }[] = [];
    let a0 = start;
    abs.forEach((v, i) => {
      if (!(v > 0)) return;
      const a1 = a0 + (v / total) * TAU;
      wedges.push({ i, a0, a1 });
      a0 = a1;
    });
    const path = (x0: number, x1: number): string =>
      m.type === 'doughnut' ? ringPath(cxp, cyp, ro, ri, x0, x1) : sectorPath(cxp, cyp, ro, x0, x1);
    // 点の多い円（ISSUE-111）：外周の弧が線の太さ以下の扇の連なりは、境の線で内側まで覆われる（線が不透明・実線の時）。
    // その区間を線の色で塗った扇 1 つ（同じ線付き）へ束ねる。c:dPt で上書きした扇は束ねない
    const runs =
      cx.thin && stroke !== null && stroke.opacity >= 1 && stroke.dash === 'solid'
        ? thinSectorRuns(
            wedges.map((w) => ({ a0: w.a0, a1: w.a1, fixed: s.points?.has(w.i) === true })),
            ro,
            stroke.width,
          )
        : [];
    let r = 0;
    for (let k = 0; k < wedges.length; k++) {
      const w = wedges[k]!;
      const run = r < runs.length && runs[r]![0] === k ? runs[r]! : null;
      if (run !== null && stroke !== null) {
        const last = wedges[run[1]]!;
        push(cx, {
          kind: 'path',
          d: path(w.a0, last.a1),
          fill: { color: stroke.color, opacity: stroke.opacity },
          stroke,
        });
        k = run[1];
        r++;
        continue;
      }
      // varyColors が偽なら点ごとには塗らない（Excel は系列の色の 1 色。c:dPt の明示は pointFill が優先する）
      push(cx, {
        kind: 'path',
        d: path(w.a0, w.a1),
        fill: pointFill(cx, s, w.i, m.varyColors),
        stroke,
      });
    }
  });
}

// ---- 数値軸・項目軸の配置（直交座標のグラフ） ----

interface Insets {
  l: number;
  r: number;
  t: number;
  b: number;
}

/** 目盛の文字・向きなど、プロットの大きさで決まる軸の結果 */
interface Axes {
  scale: ValueScale;
  valLabels: string[];
  showValLabels: boolean;
  /** 数値軸の文字の（軸に直交する向きの）大きさ：縦軸なら最長の幅・横軸なら高さ */
  valPerp: number;
  catLabels: string[];
  /** 項目ごとの行（折り返した時だけ複数行。それ以外は 1 行＝`catLabels[i]`）。ISSUE-107 */
  catLines: string[][];
  showCatLabels: boolean;
  catIdx: number[];
  catPerp: number;
  catRot: number;
  /** 回した文字が左へ出る分：プロットの左端が面の左端（area）から最低これだけ内側に要る（px。出ない形は 0）。ISSUE-107 */
  catReach: number;
  /** 数値軸の線の側・数値軸と項目軸の文字を置く側（`low`＝下・左、`high`＝上・右） */
  valLineHigh: boolean;
  valSide: 'low' | 'high';
  catSide: 'low' | 'high';
  /** 項目軸の文字を軸の線の隣に置く（`tickLblPos` が `nextTo`・無指定）。`low`・`high` の時は端に置く */
  catAtLine: boolean;
  /** 項目軸の文字の分の余白を取るか（軸の線の隣に置く時は、線が端にある時だけ） */
  catReserve: boolean;
  /** 項目軸の線の位置の割合（数値軸上・下端／左端から） */
  crossFrac: number;
  midCat: boolean;
  catMaxMin: boolean;
}

/** 項目軸の文字の形（ISSUE-107。`rotate45` は右上がり＝−45°・`rotate90` は下から上＝−90°） */
export type CategoryLabelForm = 'horizontal' | 'wrap' | 'rotate45' | 'rotate90';

/**
 * 1 つの項目名を、幅 `slotPx` に収まるよう空白で折り返した行にする（語を左から詰める）。
 * 1 行に収まれば `[text]`。どれかの語の幅が `slotPx` を超える（空白が無い・語が長い）時は折り返せないので `null`。
 */
export function wrapCategoryLabel(text: string, sizePx: number, slotPx: number): string[] | null {
  if (estimateTextWidth(text, sizePx) <= slotPx) return [text];
  const words = text.split(/\s+/).filter((w) => w !== '');
  if (words.length < 2) return null;
  const lines: string[] = [];
  let cur = '';
  for (const w of words) {
    if (estimateTextWidth(w, sizePx) > slotPx) return null;
    const next = cur === '' ? w : `${cur} ${w}`;
    if (cur === '' || estimateTextWidth(next, sizePx) <= slotPx) cur = next;
    else {
      lines.push(cur);
      cur = w;
    }
  }
  lines.push(cur);
  return lines;
}

/**
 * 縦棒・折れ線の項目軸の文字の形を選ぶ（ISSUE-107・ADR-0031 §決定 5。`a:bodyPr rot` の明示が無い時だけ呼ぶ）。
 * 上から順に判定し、最初に当たった形にする。
 * 1. 水平：どの項目名の幅（推定）も `slotPx`（項目の間隔）以下。
 * 2. 折り返し：幅が超える項目名が全部、空白を含み、どの語も `slotPx` に入る（`rows` は行数の最大）。
 * 3. 45°：`slotPx` が 45° の境目（{@link CHART_LAYOUT.catRotateMinSlotRatio} × 文字の大きさ）以上。
 * 4. 90°：それより狭い（呼び出し側が入らない分を間引く）。
 */
export function planCategoryLabelForm(
  labels: readonly string[],
  sizePx: number,
  slotPx: number,
): { form: CategoryLabelForm; rows: number } {
  if (Number.isNaN(slotPx)) return { form: 'horizontal', rows: 1 };
  let over = false;
  let wrappable = true;
  let rows = 1;
  for (const t of labels) {
    if (estimateTextWidth(t, sizePx) <= slotPx) continue;
    over = true;
    const ls = wrapCategoryLabel(t, sizePx, slotPx);
    if (ls === null) {
      wrappable = false;
      break;
    }
    rows = Math.max(rows, ls.length);
  }
  if (!over) return { form: 'horizontal', rows: 1 };
  if (wrappable) return { form: 'wrap', rows };
  return {
    form: slotPx >= CHART_LAYOUT.catRotateMinSlotRatio * sizePx ? 'rotate45' : 'rotate90',
    rows: 1,
  };
}

/** 回転した文字の枠（幅 w・高さ h を θ 度回す）の軸に沿った大きさと直交する大きさ */
function rotatedExtent(w: number, h: number, deg: number): { along: number; perp: number } {
  const t = (deg * Math.PI) / 180;
  const c = Math.abs(Math.cos(t));
  const s = Math.abs(Math.sin(t));
  return { along: w * c + h * s, perp: w * s + h * c };
}

/** 軸のタイトルの行（無ければ空）。軸が消されていても出す（Excel の挙動に合わせない・暫定） */
function axisTitleLines(axis: ChartAxis | undefined): string[] {
  const t = axis?.title;
  return t === undefined ? [] : t.lines.map((l) => truncateLine(l));
}

/** 直交座標のグラフの軸まわりの定常の値 */
interface CartCtx {
  horizontalBars: boolean;
  valueVertical: boolean;
  catLook: TextLook;
  valLook: TextLook;
  catTitleLook: TextLook;
  valTitleLook: TextLook;
  catAxisTitle: string[];
  valAxisTitle: string[];
  catRot: number;
  /** `a:bodyPr rot` の明示がある（自動の形の判定をしない）。ISSUE-107 */
  catRotExplicit: boolean;
  catLabelsAll: string[];
}

function cartCtx(cx: Ctx): CartCtx {
  const m = cx.model;
  const horizontalBars = m.type === 'bar' && m.barDir === 'bar';
  return {
    horizontalBars,
    valueVertical: !horizontalBars,
    catLook: textLook(cx, m.catAxis?.text, CHART_DEFAULTS.axisPt),
    valLook: textLook(cx, m.valAxis?.text, CHART_DEFAULTS.axisPt),
    catTitleLook: textLook(cx, m.catAxis?.title?.text, CHART_DEFAULTS.axisTitlePt),
    valTitleLook: textLook(cx, m.valAxis?.title?.text, CHART_DEFAULTS.axisTitlePt),
    catAxisTitle: axisTitleLines(m.catAxis),
    valAxisTitle: axisTitleLines(m.valAxis),
    catRot: m.catAxis?.text.rotDeg ?? CHART_DEFAULTS.autoRotation,
    catRotExplicit: m.catAxis?.text.rotDeg !== undefined,
    catLabelsAll: m.categories.map((c) => truncateLine(c)),
  };
}

/** 項目軸の文字の自動の形の結果（ISSUE-107） */
interface AutoCatLabels {
  rot: number;
  idx: number[];
  lines: string[][];
  /** 軸の線に直交する向きの文字の大きさ（px。目盛・文字との間隔を除く） */
  perp: number;
  /** プロットの左端が面（area）の左端から最低これだけ内側に要る（px） */
  reach: number;
}

/**
 * 縦棒・折れ線の項目軸の文字の形を決め、回転・間引き・行・帯の厚さ・左の張り出しにする（ISSUE-107。規則は {@link planCategoryLabelForm}）。
 * `a:bodyPr rot` の明示が無い時だけ呼ぶ。`c:tickLblSkip` の明示は最優先で、形は表示する項目の間隔（間隔 × skip）で選ぶ。
 * 回した・折り返した帯が面の高さの {@link CHART_LAYOUT.catBandMaxRatio} を超える小さい面は `null`（今までどおり回さず間引く）。
 */
function autoCatLabels(
  cx: Ctx,
  cc: CartCtx,
  n: number,
  catLen: number,
  midCat: boolean,
  reversed: boolean,
  explicitSkip: number | undefined,
): AutoCatLabels | null {
  if (n <= 0) return null;
  const L = CHART_LAYOUT;
  const px = cc.catLook.px;
  const lineH = px * L.lineHeightRatio;
  const spacing = categorySpacing(n, catLen, midCat);
  const ex =
    explicitSkip !== undefined && Number.isFinite(explicitSkip) && explicitSkip >= 1
      ? Math.min(n, Math.floor(explicitSkip))
      : undefined;
  const step = ex ?? 1;
  const slot = spacing * step;
  const labels = cc.catLabelsAll;
  const plan = planCategoryLabelForm(
    categoryLabelIndices(n, step).map((i) => labels[i]!),
    px,
    slot,
  );
  // 90° に回した文字だけ、行の高さが間隔に入る個数に間引く（余白は小さく）。それ以外は全部（明示の skip はその値）
  const skip =
    plan.form === 'rotate90' && ex === undefined
      ? categoryLabelSkip(n, spacing, lineH, undefined, L.catRotatedSkipPadPt * PX_PER_PT)
      : step;
  const idx = categoryLabelIndices(n, skip);
  const lines = labels.map((t) =>
    plan.form === 'wrap' ? (wrapCategoryLabel(t, px, slot) ?? [t]) : [t],
  );
  const widths = idx.map((i) => estimateTextWidth(labels[i]!, px));
  const tickGap = L.tickLengthPt * PX_PER_PT + L.labelGapPx;
  const rotated = plan.form === 'rotate45' || plan.form === 'rotate90';
  let perp = lineH;
  if (plan.form === 'wrap') perp = lineH + (plan.rows - 1) * px * L.catWrapRowRatio;
  else if (rotated) {
    const wMax = widths.reduce((a, b) => Math.max(a, b), 0);
    perp = Math.max(
      lineH,
      (wMax + lineH) * Math.SQRT1_2 + L.catRotatedBandExtraPt * PX_PER_PT - tickGap,
    );
  }
  if (plan.form !== 'horizontal' && perp + tickGap > cx.H * L.catBandMaxRatio) return null;
  let reach = 0;
  if (rotated) {
    const centers = categoryPositions(n, catLen, { midCat, reversed });
    idx.forEach((i, k) => {
      reach = Math.max(reach, (widths[k]! + lineH) * Math.SQRT1_2 - centers[i]!);
    });
    reach = Math.max(0, reach + L.catReachGapPt * PX_PER_PT);
  }
  return { rot: rotated ? (plan.form === 'rotate45' ? -45 : -90) : 0, idx, lines, perp, reach };
}

/**
 * プロットの大きさ（px）から、軸の範囲・目盛の文字・項目の文字の間引き・文字を置く側を決める。
 * `scaleW` は横向きの数値軸（横棒）の目盛の数を決める長さ（px。省略はプロットの幅）。
 * Excel は横棒の刻みを、項目軸の文字・目盛の文字の余白を引く前の幅（自動配置の外側）で決める（C1・ISSUE-110）。
 */
function solveAxes(cx: Ctx, cc: CartCtx, plotW: number, plotH: number, scaleW?: number): Axes {
  const m = cx.model;
  const L = CHART_LAYOUT;
  const valAxis = m.valAxis;
  const catAxis = m.catAxis;
  const valLen = cc.valueVertical ? plotH : (scaleW ?? plotW);
  const catLen = cc.valueVertical ? plotW : plotH;
  const lineH = (look: TextLook): number => look.px * L.lineHeightRatio;
  const scaleInput = {
    axis: valAxis,
    extent: dataExtent(m.grouping, m.series),
    percentStacked: m.grouping === 'percentStacked',
    lengthPx: valLen,
    labelHeightPx: lineH(cc.valLook),
  };
  const labelsOf = (sc: ValueScale): string[] => {
    const code = tickFormatCode(valAxis, m.series[0]?.formatCode, sc.percent);
    return sc.majorTicks.map((t) => truncateLine(tickLabel(code, t)));
  };
  const widest = (ls: readonly string[]): number =>
    ls.reduce((a, t) => Math.max(a, estimateTextWidth(t, cc.valLook.px)), 0);
  let scale = computeValueScale(scaleInput);
  let valLabels = labelsOf(scale);
  if (!cc.valueVertical) {
    // 横向きの数値軸（横棒）：目盛の数の上限は最長の目盛の文字の幅で決める
    scale = computeValueScale({ ...scaleInput, labelWidthPx: widest(valLabels) });
    valLabels = labelsOf(scale);
  }
  const showValLabels = valAxis?.deleted !== true && valAxis?.tickLblPos !== 'none';
  const valPerp = showValLabels ? (cc.valueVertical ? widest(valLabels) : lineH(cc.valLook)) : 0;

  const midCat = m.type === 'line' && (valAxis?.crossBetween ?? catAxis?.crossBetween) === 'midCat';
  const catMaxMin = catAxis?.orientation === 'maxMin';
  const n = m.categories.length;
  const showCatLabels = catAxis?.deleted !== true && catAxis?.tickLblPos !== 'none';
  // 縦棒・折れ線で `a:bodyPr rot` の明示が無ければ、形（水平・折り返し・45°・90°）を自動で選ぶ（ISSUE-107）。それ以外・小さい面は従来どおり
  const auto =
    cc.valueVertical && !cc.catRotExplicit && showCatLabels
      ? autoCatLabels(cx, cc, n, catLen, midCat, catMaxMin, catAxis?.tickLblSkip)
      : null;
  let catIdx: number[];
  let catPerp: number;
  if (auto !== null) {
    catIdx = auto.idx;
    catPerp = auto.perp;
  } else {
    let maxAlong = 0;
    const dims = cc.catLabelsAll.map((t) => {
      const e = rotatedExtent(estimateTextWidth(t, cc.catLook.px), lineH(cc.catLook), cc.catRot);
      // 横軸なら水平方向が「軸に沿う」向き、縦軸（横棒）なら垂直方向
      const d = cc.valueVertical ? e : { along: e.perp, perp: e.along };
      maxAlong = Math.max(maxAlong, d.along);
      return d;
    });
    const skip = categoryLabelSkip(
      n,
      categorySpacing(n, catLen, midCat),
      maxAlong,
      catAxis?.tickLblSkip,
    );
    catIdx = showCatLabels ? categoryLabelIndices(n, skip) : [];
    catPerp = catIdx.reduce((a, i) => Math.max(a, dims[i]!.perp), 0);
  }

  const crossFrac = valueFraction(scale, crossValue(scale, catAxis));
  const pos = catAxis?.tickLblPos;
  const atMax = crossFrac >= 1 - 1e-9 && crossFrac > 1e-9;
  const atMin = crossFrac <= 1e-9;
  const catAtLine = pos !== 'low' && pos !== 'high';
  // 項目軸が数値軸の最大の端に来るのが「データだけの理由」（自動の交差で、全部が負・数値軸の反転なし）の時は、文字を軸の線の低い側（縦軸の下・横軸の左）に
  // 置いて余白を取らない（Excel 16 の COM：全部が負の 100% 積み上げで、項目軸の文字の帯がプロットの内側に重なる・C2）。
  // `crosses`・`crossesAt` の明示や数値軸の反転で端に来る時は従来どおり外側（高い側）
  const crossAuto =
    (catAxis?.crosses ?? 'autoZero') === 'autoZero' && catAxis?.crossesAt === undefined;
  const maxByData = atMax && crossAuto && !scale.reversed;
  const catSide: 'low' | 'high' =
    pos === 'high' ? 'high' : pos === 'low' ? 'low' : atMax && !maxByData ? 'high' : 'low';
  // 線の隣に置く時は、線が端（文字の側）にある時だけ余白を取る
  const catReserve = !catAtLine || (catSide === 'high' ? atMax : atMin);
  const valLineHigh = catMaxMin !== (valAxis?.crosses === 'max');
  const vpos = valAxis?.tickLblPos;
  const valSide: 'low' | 'high' =
    vpos === 'high' ? 'high' : vpos === 'low' ? 'low' : valLineHigh ? 'high' : 'low';
  return {
    scale,
    valLabels,
    showValLabels,
    valPerp,
    catLabels: cc.catLabelsAll,
    catLines: auto?.lines ?? cc.catLabelsAll.map((t) => [t]),
    showCatLabels,
    catIdx,
    catPerp,
    catRot: auto?.rot ?? cc.catRot,
    catReach: auto?.reach ?? 0,
    valLineHigh,
    valSide,
    catSide,
    catAtLine,
    catReserve,
    crossFrac,
    midCat,
    catMaxMin,
  };
}

/** 軸の文字・軸タイトルが外側に取る余白 */
function axisInsets(cc: CartCtx, ax: Axes): Insets {
  const L = CHART_LAYOUT;
  const tick = L.tickLengthPt * PX_PER_PT;
  const ins: Insets = { l: 0, r: 0, t: 0, b: 0 };
  const band = (horizontalAxis: boolean, side: 'low' | 'high', size: number): void => {
    if (size <= 0) return;
    const key = horizontalAxis ? (side === 'low' ? 'b' : 't') : side === 'low' ? 'l' : 'r';
    ins[key] += size;
  };
  const valBand = ax.showValLabels ? ax.valPerp + tick + L.labelGapPx : 0;
  const catBand = ax.showCatLabels && ax.catReserve ? ax.catPerp + tick + L.labelGapPx : 0;
  // 横軸：縦棒・折れ線は項目軸、横棒は数値軸
  band(true, cc.valueVertical ? ax.catSide : ax.valSide, cc.valueVertical ? catBand : valBand);
  band(false, cc.valueVertical ? ax.valSide : ax.catSide, cc.valueVertical ? valBand : catBand);
  // 横向きの数値軸（横棒）：両端の目盛の文字は軸の端の外へ半分はみ出す（Excel 16 の照合：両端とも他の余白より大きい時だけ。右は余白 11pt の外に足す）
  if (!cc.valueVertical && ax.showValLabels && ax.valLabels.length > 0) {
    const half = (t: string | undefined): number =>
      t === undefined ? 0 : estimateTextWidth(t, cc.valLook.px) / 2;
    const first = ax.valLabels[0];
    const last = ax.valLabels[ax.valLabels.length - 1];
    // 項目軸の文字の帯がある側（crosses=max 等で右）は、帯とはみ出しの半幅の大きい方（足すと二重に取る）
    ins.l = Math.max(ins.l, half(ax.scale.reversed ? last : first));
    ins.r = Math.max(ins.r, half(ax.scale.reversed ? first : last));
  }
  // 回した項目軸の文字（縦棒・折れ線）：最初の項目の文字が左へ出る分、プロットの左端を内側へ（ISSUE-107）
  if (cc.valueVertical && ax.showCatLabels && ax.catReserve && ax.catReach > 0)
    ins.l = Math.max(ins.l, ax.catReach - L.marginPt * PX_PER_PT);
  const tl = (lines: string[], look: TextLook): number =>
    lines.length === 0 ? 0 : lines.length * look.px * L.lineHeightRatio + L.axisTitleGapPx;
  // 軸タイトル：横軸は下・縦軸は左（軸ごとの文字の大きさ）
  ins.b += cc.valueVertical
    ? tl(cc.catAxisTitle, cc.catTitleLook)
    : tl(cc.valAxisTitle, cc.valTitleLook);
  ins.l += cc.valueVertical
    ? tl(cc.valAxisTitle, cc.valTitleLook)
    : tl(cc.catAxisTitle, cc.catTitleLook);
  return ins;
}

/** 自動配置の外側の矩形からプロットの内側（軸の文字・軸タイトルを引いた矩形）を解く */
function solvePlot(cx: Ctx, cc: CartCtx, outer: Rect): { plot: Rect; axes: Axes } {
  let plot: Rect = { x: outer.x, y: outer.y, w: outer.w, h: outer.h };
  let axes = solveAxes(cx, cc, plot.w, plot.h, outer.w);
  for (let k = 0; k < 2; k++) {
    const ins = axisInsets(cc, axes);
    plot = {
      x: outer.x + ins.l,
      y: outer.y + ins.t,
      w: Math.max(1, outer.w - ins.l - ins.r),
      h: Math.max(1, outer.h - ins.t - ins.b),
    };
    axes = solveAxes(cx, cc, plot.w, plot.h, outer.w);
  }
  return { plot, axes };
}

/** 矩形の経路 */
const rectD = (x0: number, y0: number, x1: number, y1: number): string =>
  `M${f(x0)} ${f(y0)}L${f(x1)} ${f(y0)}L${f(x1)} ${f(y1)}L${f(x0)} ${f(y1)}Z`;

/** 点列を直線（または Catmull-Rom を 3 次ベジェにした近似）でつなぐ経路 */
function polylineD(pts: readonly (readonly [number, number])[], smooth: boolean): string {
  const p0 = pts[0]!;
  let d = `M${f(p0[0])} ${f(p0[1])}`;
  if (!smooth || pts.length < 3) {
    for (let i = 1; i < pts.length; i++) d += `L${f(pts[i]![0])} ${f(pts[i]![1])}`;
    return d;
  }
  for (let i = 0; i < pts.length - 1; i++) {
    const a = pts[Math.max(0, i - 1)]!;
    const b = pts[i]!;
    const c = pts[i + 1]!;
    const e = pts[Math.min(pts.length - 1, i + 2)]!;
    d += `C${f(b[0] + (c[0] - a[0]) / 6)} ${f(b[1] + (c[1] - a[1]) / 6)} ${f(c[0] - (e[0] - b[0]) / 6)} ${f(c[1] - (e[1] - b[1]) / 6)} ${f(c[0])} ${f(c[1])}`;
  }
  return d;
}

/** 直交座標のグラフを場面へ出す（プロットの面 → 目盛線 → 系列 → 軸線・目盛・文字 → 軸タイトル） */
function emitCartesian(cx: Ctx, cc: CartCtx, outer: Rect, plot: Rect, ax: Axes): void {
  const m = cx.model;
  const L = CHART_LAYOUT;
  const { scale } = ax;
  const vv = cc.valueVertical;
  const n = m.categories.length;
  const catLen = vv ? plot.w : plot.h;
  /** 項目の間隔が 1px 未満（点の多いグラフ）。系列の幾何を間引く（ADR-0031 §決定 11.1） */
  const dense = cx.thin && n > 0 && catLen / n < CHART_THIN.denseSpacingPx;
  const valAxis = m.valAxis;
  const catAxis = m.catAxis;

  // プロットの面（暗い背景は #3F3F3F。明示の c:spPr が優先）
  const pFill = resolveFill(
    m.plotArea.fill,
    isDarkStyle(m) ? toColor(CHART_DEFAULTS.darkPlotFill) : null,
  );
  const pStroke = resolveStroke(cx, m.plotArea.line, null);
  if (pFill !== null || pStroke !== null) {
    push(cx, {
      kind: 'rect',
      x: plot.x,
      y: plot.y,
      width: plot.w,
      height: plot.h,
      fill: pFill,
      stroke: pStroke,
    });
  }

  const valCoord = (v: number): number => {
    const fr = valueFraction(scale, v);
    return vv ? plot.y + plot.h * (1 - fr) : plot.x + plot.w * fr;
  };
  const catCoord = (off: number): number => (vv ? plot.x + off : plot.y + plot.h - off);
  /** 範囲外を端に収めない数値軸の座標（折れ線をプロットの矩形で切るため） */
  const valCoordRaw = (v: number): number => {
    const fr = valueFraction(scale, v, false);
    return vv ? plot.y + plot.h * (1 - fr) : plot.x + plot.w * fr;
  };
  const xy = (off: number, v: number): [number, number] =>
    vv ? [catCoord(off), valCoordRaw(v)] : [valCoordRaw(v), catCoord(off)];
  /** 数値軸に垂直な線（値 v の目盛線）・項目軸に垂直な線（位置 off） */
  const valLineAt = (v: number): string => {
    const c = valCoord(v);
    return vv
      ? `M${f(plot.x)} ${f(c)}L${f(plot.x + plot.w)} ${f(c)}`
      : `M${f(c)} ${f(plot.y)}L${f(c)} ${f(plot.y + plot.h)}`;
  };
  const catLineAt = (off: number): string => {
    const c = catCoord(off);
    return vv
      ? `M${f(c)} ${f(plot.y)}L${f(c)} ${f(plot.y + plot.h)}`
      : `M${f(plot.x)} ${f(c)}L${f(plot.x + plot.w)} ${f(c)}`;
  };

  // 項目の位置
  const centers = categoryPositions(n, catLen, { midCat: ax.midCat, reversed: ax.catMaxMin });
  const bandCenters = categoryPositions(n, catLen, { reversed: ax.catMaxMin });
  const boundaries: number[] = [];
  if (ax.midCat) boundaries.push(...centers);
  else if (n + 1 <= MAX_CHART_SCENE_ITEMS)
    for (let i = 0; i <= n; i++) boundaries.push((i * catLen) / Math.max(1, n));

  // 目盛線（補助 → 主）
  const grid = (line: ChartLine | undefined, spec: ColorSpec, ds: string[]): void => {
    if (line === undefined || ds.length === 0) return;
    const stroke = resolveStroke(cx, line, { spec, widthPt: CHART_DEFAULTS.axisLinePt });
    if (stroke !== null) push(cx, { kind: 'path', d: ds.join(''), fill: null, stroke });
  };
  grid(valAxis?.minorGridlines, CHART_DEFAULTS.gridMinor, scale.minorTicks.map(valLineAt));
  const gridMajor = lineSpec(cx, CHART_DEFAULTS.gridMajor);
  grid(valAxis?.majorGridlines, gridMajor, scale.majorTicks.map(valLineAt));
  grid(catAxis?.majorGridlines, gridMajor, boundaries.map(catLineAt));

  // 系列
  if (m.type === 'bar') emitBars(cx, cc, ax, bandCenters, catLen, valCoord, catCoord, dense);
  else emitLines(cx, centers, xy, plot, dense ? { catAxis: vv ? 0 : 1 } : null);

  // 軸線
  const crossC = valCoord(crossValue(scale, catAxis));
  const catStroke =
    catAxis?.deleted === true
      ? null
      : resolveStroke(cx, catAxis?.line, {
          spec: lineSpec(cx, CHART_DEFAULTS.catAxisLine),
          widthPt: CHART_DEFAULTS.axisLinePt,
        });
  if (catStroke !== null) {
    if (vv)
      push(cx, {
        kind: 'line',
        x1: plot.x,
        y1: crossC,
        x2: plot.x + plot.w,
        y2: crossC,
        stroke: catStroke,
      });
    else
      push(cx, {
        kind: 'line',
        x1: crossC,
        y1: plot.y,
        x2: crossC,
        y2: plot.y + plot.h,
        stroke: catStroke,
      });
  }
  const valStroke =
    valAxis?.deleted === true
      ? null
      : resolveStroke(cx, valAxis?.line, {
          spec: lineSpec(cx, CHART_DEFAULTS.valAxisLine),
          widthPt: CHART_DEFAULTS.axisLinePt,
        });
  const valLinePos = vv
    ? ax.valLineHigh
      ? plot.x + plot.w
      : plot.x
    : ax.valLineHigh
      ? plot.y
      : plot.y + plot.h;
  if (valStroke !== null) {
    if (vv)
      push(cx, {
        kind: 'line',
        x1: valLinePos,
        y1: plot.y,
        x2: valLinePos,
        y2: plot.y + plot.h,
        stroke: valStroke,
      });
    else
      push(cx, {
        kind: 'line',
        x1: plot.x,
        y1: valLinePos,
        x2: plot.x + plot.w,
        y2: valLinePos,
        stroke: valStroke,
      });
  }

  // 目盛の印（要素が無ければ {@link CHART_DEFAULTS.tickMark}＝交差。C12）
  const tickMarks = (
    markOrNone: ChartAxis['majorTickMark'],
    positions: readonly number[],
    horizontalAxis: boolean,
    linePos: number,
    labelsHigh: boolean,
    len: number,
    stroke: ChartStroke | null,
  ): void => {
    const mark = markOrNone ?? CHART_DEFAULTS.tickMark;
    if (stroke === null || mark === 'none' || positions.length === 0) return;
    // 外向き＝文字を置く側
    const out = horizontalAxis ? (labelsHigh ? -1 : 1) : labelsHigh ? 1 : -1;
    const signs = mark === 'out' ? [out] : mark === 'in' ? [-out] : [out, -out];
    let d = '';
    for (const p of positions) {
      for (const sg of signs) {
        d += horizontalAxis
          ? `M${f(p)} ${f(linePos)}L${f(p)} ${f(linePos + sg * len)}`
          : `M${f(linePos)} ${f(p)}L${f(linePos + sg * len)} ${f(p)}`;
      }
    }
    push(cx, { kind: 'path', d, fill: null, stroke });
  };
  const tickLen = L.tickLengthPt * PX_PER_PT;
  const minorLen = L.minorTickLengthPt * PX_PER_PT;
  const tickStroke = (axis: ChartAxis | undefined, def: ColorSpec): ChartStroke | null =>
    axis?.deleted === true
      ? null
      : resolveStroke(cx, axis?.line, { spec: def, widthPt: CHART_DEFAULTS.axisLinePt });
  const valTickStroke = tickStroke(valAxis, lineSpec(cx, CHART_DEFAULTS.catAxisLine));
  const catTickStroke = tickStroke(catAxis, lineSpec(cx, CHART_DEFAULTS.catAxisLine));
  // 項目軸の印：点が画素より細かい（間引きが効く）グラフでは、帯のようになるので出さない（要素が既定で増えすぎない）
  const catPosAbs = dense ? [] : boundaries.map(catCoord);
  tickMarks(
    catAxis?.majorTickMark,
    catPosAbs,
    vv,
    crossC,
    ax.catSide === 'high',
    tickLen,
    catTickStroke,
  );
  tickMarks(
    valAxis?.majorTickMark,
    scale.majorTicks.map(valCoord),
    !vv,
    valLinePos,
    ax.valSide === 'high',
    tickLen,
    valTickStroke,
  );
  tickMarks(
    valAxis?.minorTickMark,
    scale.minorTicks.map(valCoord),
    !vv,
    valLinePos,
    ax.valSide === 'high',
    minorLen,
    valTickStroke,
  );

  // 数値軸の文字
  const gap = L.labelGapPx;
  if (ax.showValLabels) {
    const hi = ax.valSide === 'high';
    ax.valLabels.forEach((t, i) => {
      const v = scale.majorTicks[i]!;
      if (vv) {
        const x = hi ? plot.x + plot.w + tickLen + gap : plot.x - tickLen - gap;
        pushText(cx, x, valCoord(v), t, cc.valLook, hi ? 'start' : 'end', 'middle');
      } else {
        const y = hi ? plot.y - tickLen - gap : plot.y + plot.h + tickLen + gap;
        pushText(cx, valCoord(v), y, t, cc.valLook, 'middle', hi ? 'alphabetic' : 'hanging');
      }
    });
  }
  // 項目軸の文字
  if (ax.showCatLabels) {
    const hi = ax.catSide === 'high';
    const rot = ax.catRot;
    // 基準：軸の線の隣（nextTo）か、プロットの端（low・high）
    const edgeBase = vv ? (hi ? plot.y : plot.y + plot.h) : hi ? plot.x + plot.w : plot.x;
    const base = ax.catAtLine ? crossC : edgeBase;
    for (const i of ax.catIdx) {
      const off = centers[i]!;
      const t = ax.catLabels[i]!;
      if (vv) {
        const y = hi ? base - tickLen - gap : base + tickLen + gap;
        if (rot === 0) {
          // 折り返した行は、軸の線に近い行から順に外へ（下の軸は上から下へ・上の軸は下から上へ。行の送りは文字の大きさ × catWrapRowRatio）
          const ls = ax.catLines[i] ?? [t];
          const pitch = cc.catLook.px * L.catWrapRowRatio;
          ls.forEach((line, k) =>
            pushText(
              cx,
              catCoord(off),
              hi ? y - (ls.length - 1 - k) * pitch : y + k * pitch,
              line,
              cc.catLook,
              'middle',
              hi ? 'alphabetic' : 'hanging',
            ),
          );
        } else
          pushText(
            cx,
            catCoord(off),
            y,
            t,
            cc.catLook,
            rot < 0 !== hi ? 'end' : 'start',
            'middle',
            rot,
          );
      } else {
        const x = hi ? base + tickLen + gap : base - tickLen - gap;
        pushText(cx, x, catCoord(off), t, cc.catLook, hi ? 'start' : 'end', 'middle', rot);
      }
    }
  }
  // 軸タイトル：横軸は下・縦軸は左（縦の軸のタイトルは −90°）
  const emitAxisTitle = (
    axis: ChartAxis | undefined,
    lines: string[],
    horizontalAxis: boolean,
  ): void => {
    if (lines.length === 0) return;
    const look = textLook(cx, axis?.title?.text, CHART_DEFAULTS.axisTitlePt);
    const lineH = look.px * L.lineHeightRatio;
    lines.forEach((t, k) => {
      if (horizontalAxis) {
        const y = outer.y + outer.h - (lines.length - k - 0.5) * lineH;
        pushText(
          cx,
          plot.x + plot.w / 2,
          y,
          t,
          look,
          'middle',
          'middle',
          axis?.title?.text.rotDeg ?? 0,
        );
      } else {
        const x = outer.x + (k + 0.5) * lineH;
        pushText(
          cx,
          x,
          plot.y + plot.h / 2,
          t,
          look,
          'middle',
          'middle',
          axis?.title?.text.rotDeg ?? -90,
        );
      }
    });
  };
  emitAxisTitle(catAxis, cc.catAxisTitle, vv);
  emitAxisTitle(valAxis, cc.valAxisTitle, !vv);
}

type XY = (off: number, v: number) => [number, number];

/** 棒（集合・積み上げ・100% 積み上げ）。系列ごとに塗りの違いごとの経路にまとめる */
function emitBars(
  cx: Ctx,
  cc: CartCtx,
  ax: Axes,
  centers: readonly number[],
  catLen: number,
  valCoord: (v: number) => number,
  catCoord: (off: number) => number,
  dense: boolean,
): void {
  const m = cx.model;
  const vv = cc.valueVertical;
  const n = m.categories.length;
  const stacked = m.grouping === 'stacked' || m.grouping === 'percentStacked';
  const nBars = stacked ? 1 : m.series.length;
  const g = (m.gapWidth ?? CHART_DEFAULTS.gapWidth) / 100;
  const o =
    (m.overlap ?? (stacked ? CHART_DEFAULTS.overlapStacked : CHART_DEFAULTS.overlapClustered)) /
    100;
  const band = catLen / Math.max(1, n);
  const denom = nBars - (nBars - 1) * o + g;
  const w = denom > 0 ? band / denom : band;
  const step = w * (1 - o);
  const groupW = w + (nBars - 1) * step;
  const base = crossValue(ax.scale, cx.model.catAxis);
  const pos = new Array<number>(n).fill(0);
  const neg = new Array<number>(n).fill(0);
  const totals: number[] = [];
  if (m.grouping === 'percentStacked') {
    for (let i = 0; i < n; i++) {
      totals.push(
        m.series.reduce(
          (a, s) =>
            a +
            (typeof s.values[i] === 'number' && Number.isFinite(s.values[i]!)
              ? Math.abs(s.values[i]!)
              : 0),
          0,
        ),
      );
    }
  }
  const vary = varyByPoint(m);
  // 点の多い集合の棒（重なり < 100）は、列の幅を系列の数に合わせて広げる（ISSUE-111）。積み上げ・100% は 1px のまま
  const barColPx =
    !stacked && o < 1 ? clusteredBarColumnPx(m.series.length) : CHART_THIN.barColumnPx;
  m.series.forEach((s, j) => {
    const groups = new Map<string, { fill: ChartColor | null; ds: string[]; boxes: BarBoxIn[] }>();
    for (let i = 0; i < n; i++) {
      const raw = s.values[i];
      if (typeof raw !== 'number' || !Number.isFinite(raw)) continue;
      let v0: number;
      let v1: number;
      if (stacked) {
        const v = m.grouping === 'percentStacked' ? (totals[i]! > 0 ? raw / totals[i]! : 0) : raw;
        if (v >= 0) {
          v0 = pos[i]!;
          v1 = v0 + v;
          pos[i] = v1;
        } else {
          v0 = neg[i]!;
          v1 = v0 + v;
          neg[i] = v1;
        }
      } else {
        v0 = base;
        v1 = raw;
      }
      const slot = stacked ? 0 : ax.catMaxMin ? nBars - 1 - j : j;
      const lo = centers[i]! - groupW / 2 + slot * step;
      const hi = lo + w;
      const a = valCoord(v0);
      const b = valCoord(v1);
      // 軸の範囲から完全に外れた棒は出さない（かかる棒は端に収めて描く）
      if (Math.max(v0, v1) < ax.scale.min || Math.min(v0, v1) > ax.scale.max) continue;
      if (Math.abs(a - b) < 1e-9) continue;
      const c1 = catCoord(lo);
      const c2 = catCoord(hi);
      const fill = pointFill(cx, s, i, vary);
      const key = fill === null ? 'none' : `${fill.color}/${fill.opacity}`;
      let grp = groups.get(key);
      if (grp === undefined) {
        grp = { fill, ds: [], boxes: [] };
        groups.set(key, grp);
      }
      if (dense) {
        // 点の多いグラフ：列ごとにまとめるため、箱を集める（正の側と負の側は別）
        grp.boxes.push({
          lo: Math.min(c1, c2),
          hi: Math.max(c1, c2),
          a: Math.min(a, b),
          b: Math.max(a, b),
          neg: v1 < v0,
        });
      } else {
        grp.ds.push(
          vv
            ? rectD(Math.min(c1, c2), Math.min(a, b), Math.max(c1, c2), Math.max(a, b))
            : rectD(Math.min(a, b), Math.min(c1, c2), Math.max(a, b), Math.max(c1, c2)),
        );
      }
    }
    const stroke = resolveStroke(cx, s.line, null);
    for (const grp of groups.values()) {
      if (grp.fill === null && stroke === null) continue;
      const d = dense
        ? mergeBarColumns(grp.boxes, barColPx)
            .map((bx: BarBox) =>
              vv ? rectD(bx.lo, bx.a, bx.hi, bx.b) : rectD(bx.a, bx.lo, bx.b, bx.hi),
            )
            .join('')
        : grp.ds.join('');
      push(cx, { kind: 'path', d, fill: grp.fill, stroke });
    }
  });
}

/** 線分 p→q を矩形（境界を含む）で切る（Liang–Barsky）。矩形の外なら null。t0・t1 は元の線分上の位置（0〜1） */
function clipSegment(
  p: readonly [number, number],
  q: readonly [number, number],
  r: Rect,
): { a: [number, number]; b: [number, number]; t0: number; t1: number } | null {
  const eps = 1e-6;
  const dx = q[0] - p[0];
  const dy = q[1] - p[1];
  let t0 = 0;
  let t1 = 1;
  const edges: [number, number][] = [
    [-dx, p[0] - (r.x - eps)],
    [dx, r.x + r.w + eps - p[0]],
    [-dy, p[1] - (r.y - eps)],
    [dy, r.y + r.h + eps - p[1]],
  ];
  for (const [pp, qq] of edges) {
    if (pp === 0) {
      if (qq < 0) return null;
      continue;
    }
    const t = qq / pp;
    if (pp < 0) {
      if (t > t1) return null;
      if (t > t0) t0 = t;
    } else {
      if (t < t0) return null;
      if (t < t1) t1 = t;
    }
  }
  return {
    a: [p[0] + t0 * dx, p[1] + t0 * dy],
    b: [p[0] + t1 * dx, p[1] + t1 * dy],
    t0,
    t1,
  };
}

/** 点列の折れ線を矩形で切る。矩形の外を通る区間で分かれ、2 点以上の部分だけを返す（端に寄せず、傾きを保つ） */
function clipPolyline(pts: readonly (readonly [number, number])[], r: Rect): [number, number][][] {
  const out: [number, number][][] = [];
  let cur: [number, number][] = [];
  const flush = (): void => {
    const p0 = cur[0];
    if (p0 !== undefined && cur.some((p) => p[0] !== p0[0] || p[1] !== p0[1])) out.push(cur);
    cur = [];
  };
  /** 直前の線分が終点まで矩形の中に入っていた（＝次の線分は続きとしてつなぐ） */
  let open = false;
  for (let i = 0; i + 1 < pts.length; i++) {
    const c = clipSegment(pts[i]!, pts[i + 1]!, r);
    if (c === null) {
      flush();
      open = false;
      continue;
    }
    if (!(open && c.t0 === 0)) {
      flush();
      cur = [c.a];
    }
    cur.push(c.b);
    open = c.t1 === 1;
  }
  flush();
  return out;
}

/** 点が矩形の中（境界を含む）か */
function inRect(p: readonly [number, number], r: Rect): boolean {
  const eps = 1e-6;
  return (
    p[0] >= r.x - eps && p[0] <= r.x + r.w + eps && p[1] >= r.y - eps && p[1] <= r.y + r.h + eps
  );
}

/** 滑らかな線（Catmull-Rom）を、切るために折れ線に直す（区間ごとに `steps` 分割） */
function flattenSmooth(
  pts: readonly (readonly [number, number])[],
  steps: number,
): [number, number][] {
  const out: [number, number][] = [[pts[0]![0], pts[0]![1]]];
  for (let i = 0; i < pts.length - 1; i++) {
    const a = pts[Math.max(0, i - 1)]!;
    const b = pts[i]!;
    const c = pts[i + 1]!;
    const e = pts[Math.min(pts.length - 1, i + 2)]!;
    const c1: [number, number] = [b[0] + (c[0] - a[0]) / 6, b[1] + (c[1] - a[1]) / 6];
    const c2: [number, number] = [c[0] - (e[0] - b[0]) / 6, c[1] - (e[1] - b[1]) / 6];
    for (let k = 1; k <= steps; k++) {
      const t = k / steps;
      const u = 1 - t;
      const w0 = u * u * u;
      const w1 = 3 * u * u * t;
      const w2 = 3 * u * t * t;
      const w3 = t * t * t;
      out.push([
        w0 * b[0] + w1 * c1[0] + w2 * c2[0] + w3 * c[0],
        w0 * b[1] + w1 * c1[1] + w2 * c2[1] + w3 * c[1],
      ]);
    }
  }
  return out;
}

/**
 * 折れ線（空の点 3 種・積み上げの累積・滑らか・マーカー）。
 * `xy` は範囲外を端に収めない座標。軸の範囲の外へ出る区間は `clip`（プロットの矩形）で幾何的に切り（傾きを保つ）、
 * 外を通る区間で線は複数の部分に分かれる。プロットの外の点のマーカーは出さない。
 * `thin` が非 null なら点の多いグラフ：線は部分ごとに M4→Douglas–Peucker で間引き（滑らかな線は 4 分割の折れ線にして `C` を出さない）、
 * マーカーは既定の見た目の円を帯に・他を間引く（`c:dPt` の点・`varyColors` の点ごとの色は間引かない。ADR-0031 §決定 11.1）。
 */
function emitLines(
  cx: Ctx,
  centers: readonly number[],
  xy: XY,
  clip: Rect,
  thin: { catAxis: 0 | 1 } | null,
): void {
  const m = cx.model;
  const n = m.categories.length;
  const stacked = m.grouping === 'stacked' || m.grouping === 'percentStacked';
  const cum = new Array<number>(n).fill(0);
  const totals: number[] = [];
  if (m.grouping === 'percentStacked') {
    for (let i = 0; i < n; i++) {
      totals.push(
        m.series.reduce(
          (a, s) =>
            a +
            (typeof s.values[i] === 'number' && Number.isFinite(s.values[i]!)
              ? Math.abs(s.values[i]!)
              : 0),
          0,
        ),
      );
    }
  }
  const vary = varyByPoint(m);
  m.series.forEach((s) => {
    const pts: ([number, number] | null)[] = [];
    for (let i = 0; i < n; i++) {
      const raw = s.values[i];
      const ok = typeof raw === 'number' && Number.isFinite(raw);
      if (!ok && m.dispBlanksAs !== 'zero') {
        pts.push(null);
        continue;
      }
      let v = ok ? (raw as number) : 0;
      if (m.grouping === 'percentStacked') v = totals[i]! > 0 ? v / totals[i]! : 0;
      if (stacked) {
        cum[i] = cum[i]! + v;
        v = cum[i]!;
      }
      pts.push(xy(centers[i]!, v));
    }
    // 線：空の点で分ける（span は飛ばしてつなぐ）
    const segs: [number, number][][] = [];
    let cur: [number, number][] = [];
    for (const p of pts) {
      if (p === null) {
        if (m.dispBlanksAs === 'gap' && cur.length > 0) {
          segs.push(cur);
          cur = [];
        }
        continue;
      }
      cur.push(p);
    }
    if (cur.length > 0) segs.push(cur);
    const lineSpecColor = autoColor(cx.theme, s.idx);
    const stroke = resolveStroke(
      cx,
      s.line,
      { spec: lineSpecColor, widthPt: CHART_DEFAULTS.seriesLinePt },
      { cap: 'round', join: 'round' },
    );
    if (stroke !== null) {
      const d = segs
        .filter((sg) => sg.length >= 2)
        .map((sg) => {
          const smooth = s.smooth === true && sg.length >= 3;
          const inside = sg.every((p) => inRect(p, clip));
          if (thin !== null) {
            // 点の多いグラフ：折れ線にして（範囲外は切って）、部分ごとに間引く。`C` は出さない
            const flat = smooth ? flattenSmooth(sg, CHART_THIN.smoothSteps) : sg;
            const pieces = inside ? [flat] : clipPolyline(flat, clip);
            return pieces
              .map((piece) => polylineD(thinPolyline(piece, thin.catAxis), false))
              .join('');
          }
          if (inside) return polylineD(sg, smooth);
          // 範囲外を含む：（滑らかなら折れ線に直して）矩形で切り、部分ごとに M から
          return clipPolyline(smooth ? flattenSmooth(sg, 16) : sg, clip)
            .map((piece) => polylineD(piece, false))
            .join('');
        })
        .join('');
      if (d !== '') push(cx, { kind: 'path', d, fill: null, stroke });
    }
    // マーカー：塗りの違いごとの経路
    const sym = markerSymbol(s);
    if (sym === null) return;
    const r = ((s.marker?.size ?? CHART_DEFAULTS.markerSizePt) * PX_PER_PT) / 2;
    const base: ChartColor =
      s.line?.paint?.kind === 'solid' ? toColor(s.line.paint.color) : seriesAuto(cx, s);
    const filled = markerPath(sym, 0, 0, r).filled;
    const groups = new Map<
      string,
      {
        fill: ChartColor | null;
        stroke: ChartStroke | null;
        /** 間引かない点（と密でない時の全点）のマーカーの経路 */
        d: string;
        /** 密な時に間引く点の中心 */
        centers: Pt[];
      }
    >();
    pts.forEach((p, i) => {
      if (p === null || !inRect(p, clip)) return;
      const dpt = s.points?.get(i);
      let fill: ChartColor | null;
      if (dpt !== undefined) fill = resolveFill(dpt, null);
      else if (s.marker?.fill !== undefined) fill = resolveFill(s.marker.fill, null);
      else fill = vary ? toColor(autoColor(cx.theme, i)) : base;
      const key = fill === null ? 'none' : `${fill.color}/${fill.opacity}`;
      let grp = groups.get(key);
      if (grp === undefined) {
        const ms =
          s.marker?.line !== undefined
            ? resolveStroke(
                cx,
                s.marker.line,
                fill === null
                  ? null
                  : {
                      spec: { rgb: parseInt(fill.color.slice(1), 16), alpha: fill.opacity },
                      widthPt: CHART_DEFAULTS.thinLinePt,
                    },
              )
            : fill === null
              ? null
              : strokeFrom(fill);
        grp = { fill: filled ? fill : null, stroke: ms, d: '', centers: [] };
        groups.set(key, grp);
      }
      // 点ごとの上書き（c:dPt）・点ごとの色（varyColors）の点は間引かない
      if (thin === null || dpt !== undefined || vary) grp.d += markerPath(sym, p[0], p[1], r).d;
      else grp.centers.push(p);
    });
    for (const grp of groups.values()) {
      if (grp.fill === null && grp.stroke === null) continue;
      let band: ChartPathItem | null = null;
      if (thin !== null && grp.centers.length > 0) {
        const st = grp.stroke;
        const fl = grp.fill;
        const circle = markerPath(sym, 0, 0, r).d.includes('A');
        if (
          circle &&
          fl !== null &&
          st !== null &&
          fl.opacity === 1 &&
          st.opacity === 1 &&
          st.color === fl.color &&
          st.dash === 'solid'
        ) {
          // 既定の見た目の円：続く中心を太さ 2R・round の線（帯）にする（R＝半径＋枠の太さの半分）
          const R = r + st.width / 2;
          const { runs, singles } = splitMarkerBands(grp.centers, bandGap(R));
          for (const p of singles) grp.d += markerPath(sym, p[0], p[1], r).d;
          const bd = runs.map((run) => polylineD(thinPolyline(run, thin.catAxis), false)).join('');
          if (bd !== '')
            band = {
              kind: 'path',
              d: bd,
              fill: null,
              stroke: {
                color: fl.color,
                opacity: 1,
                width: 2 * R,
                dash: 'solid',
                cap: 'round',
                join: 'round',
              },
            };
        } else {
          // その他のマーカー：Douglas–Peucker の点は残し、間は直前に残した点から 1px 以内を省く
          for (const i of thinMarkerIndices(grp.centers)) {
            const p = grp.centers[i]!;
            grp.d += markerPath(sym, p[0], p[1], r).d;
          }
        }
      }
      if (grp.d !== '') push(cx, { kind: 'path', d: grp.d, fill: grp.fill, stroke: grp.stroke });
      if (band !== null) push(cx, band);
    }
  });
}

// ---- 公開 ----

/**
 * グラフの場面を計算する（ISSUE-102・ADR-0031 §決定 3・6）。
 * @param model `parseChart` の結果のグラフ
 * @param widthPx グラフの面の幅（px）
 * @param heightPx グラフの面の高さ（px）
 * @param theme テーマの色（既定の色に使う）
 * @param opts `thin`＝点の多いグラフ（項目の間隔 1px 未満）の幾何を間引くか（既定 真。ADR-0031 §決定 11.1。偽は照合・試験で間引きの有無を比べる用）
 * @returns 場面。大きさが正の有限値でない・場面の要素が {@link MAX_CHART_SCENE_ITEMS} を超えるなら null（表示側は破線枠）
 */
export function layoutChart(
  model: ChartModel,
  widthPx: number,
  heightPx: number,
  theme: Theme = DEFAULT_THEME,
  opts?: { thin?: boolean },
): ChartScene | null {
  if (!(Number.isFinite(widthPx) && widthPx > 0 && Number.isFinite(heightPx) && heightPx > 0))
    return null;
  const pie = model.type === 'pie' || model.type === 'doughnut';
  if (
    pie &&
    (model.type === 'pie' ? 1 : model.series.length) * model.categories.length >
      MAX_CHART_SCENE_ITEMS
  )
    return null;
  const rounded = model.roundedCorners ?? CHART_DEFAULTS.roundedCorners;
  const radius = rounded
    ? Math.max(0, Math.min(CHART_DEFAULTS.cornerRadiusPt * PX_PER_PT, widthPx / 2, heightPx / 2))
    : 0;
  const cx: Ctx = {
    model,
    theme,
    W: widthPx,
    H: heightPx,
    items: [],
    thin: opts?.thin ?? true,
    radius,
  };
  try {
    build(cx, pie);
  } catch (e) {
    if (e instanceof SceneOverflow) return null;
    throw e;
  }
  return {
    width: widthPx,
    height: heightPx,
    backdrop: model.chartArea.fill?.kind === 'none',
    cornerRadius: radius,
    items: cx.items,
  };
}

/** 場面を組み立てる（グラフの面 → プロット・系列 → 凡例 → タイトルの順に出す） */
function build(cx: Ctx, pie: boolean): void {
  const m = cx.model;
  const L = CHART_LAYOUT;
  const { W, H } = cx;

  // グラフの面
  const dark = isDarkStyle(m);
  const caFill = resolveFill(
    m.chartArea.fill,
    toColor(
      dark ? CHART_DEFAULTS.darkChartAreaFill : specColor(cx.theme, CHART_DEFAULTS.chartAreaFill),
    ),
  );
  const caStroke = resolveStroke(cx, m.chartArea.line, {
    spec: dark ? CHART_DEFAULTS.darkChartAreaLine : CHART_DEFAULTS.chartAreaLine,
    widthPt: CHART_DEFAULTS.chartAreaLinePt,
  });
  if (caFill !== null || caStroke !== null) {
    const rect: ChartRectItem = {
      kind: 'rect',
      x: 0,
      y: 0,
      width: W,
      height: H,
      fill: caFill,
      stroke: caStroke,
    };
    if (cx.radius > 0) rect.radius = cx.radius;
    push(cx, rect);
  }

  const margin = L.marginPt * PX_PER_PT;
  const pad = L.plotPadPt * PX_PER_PT;
  // タイトル・凡例が引いた後の残りの面（グラフの面全体から始める。プロットの余白は後で足す）
  const area: Rect = { x: 0, y: 0, w: W, h: H };

  // タイトル（自動：上端中央・帯を取る。文字の無い自動のタイトルは帯だけ取り、文字は描かない）
  const blankTitle = isBlankAutoTitle(m);
  const tLines = blankTitle ? [''] : titleLines(m);
  let emitTitle: (() => void) | null = null;
  if (tLines.length > 0) {
    const look = textLook(
      cx,
      m.title?.text,
      m.c14Style !== undefined ? CHART_DEFAULTS.c14TitlePt : CHART_DEFAULTS.titlePt,
      CHART_DEFAULTS.titleBold,
    );
    const lineH = look.px * L.titleLineHeightRatio;
    const tw = tLines.reduce((a, t) => Math.max(a, estimateTextWidth(t, look.px)), 0);
    const th = tLines.length * lineH;
    const auto: Rect = { x: (W - tw) / 2, y: L.titleTopPt * PX_PER_PT, w: tw, h: th };
    const manual = hasManual(m.title?.layout);
    const box = manual ? applyManual(m.title!.layout!, W, H, auto) : auto;
    if (!manual && m.title?.overlay !== true) {
      area.y = auto.y + th;
      area.h = Math.max(1, H - area.y);
    }
    if (!blankTitle) {
      emitTitle = (): void => {
        tLines.forEach((t, k) =>
          pushText(
            cx,
            box.x + tw / 2,
            box.y + (k + 0.5) * lineH,
            t,
            look,
            'middle',
            'middle',
            m.title?.text.rotDeg ?? 0,
          ),
        );
      };
    }
  }

  // 凡例（自動：縁に寄せて帯を取る）
  let emitLeg: (() => void) | null = null;
  const legend = m.legend;
  if (legend !== undefined) {
    const look = textLook(cx, legend.text, CHART_DEFAULTS.legendPt);
    const entries = legendEntries(cx, legend);
    const vertical = legend.pos === 'r' || legend.pos === 'l' || legend.pos === 'tr';
    const gap = L.legendEdgeGapPt * PX_PER_PT;
    const maxW = vertical ? area.w * L.legendMaxFraction : area.w;
    // 縦並びの高さは、面の上（タイトルの下）・下の縁との間（gap）を除いた高さ（入らない項目は描かない。C3：高さ 216pt・13 項目で 10 項目）
    const maxH = vertical ? Math.max(0, area.h - 2 * gap) : area.h * L.legendMaxFraction;
    let box = buildLegend(entries, legend.pos, look, maxW, maxH);
    if (box !== null) {
      let ox: number;
      let oy: number;
      switch (legend.pos) {
        case 'l':
          ox = gap;
          oy = area.y + (area.h - box.h) / 2;
          break;
        case 't':
          ox = area.x + (area.w - box.w) / 2;
          oy = area.y + gap;
          break;
        case 'b':
          ox = area.x + (area.w - box.w) / 2;
          oy = H - gap - box.h;
          break;
        case 'tr':
          ox = W - gap - box.w;
          oy = area.y + gap;
          break;
        default:
          ox = W - gap - box.w;
          oy = area.y + (area.h - box.h) / 2;
      }
      const manual = hasManual(legend.layout);
      if (manual) {
        const rect = applyManual(legend.layout!, W, H, { x: ox, y: oy, w: box.w, h: box.h });
        if (legend.layout!.w !== undefined || legend.layout!.h !== undefined) {
          box = buildLegend(entries, legend.pos, look, rect.w, rect.h) ?? box;
        }
        ox = rect.x;
        oy = rect.y;
      } else if (!legend.overlay) {
        // 残りの面から凡例を除く（面の縁からの余白を含む）
        if (legend.pos === 'l') {
          area.w = Math.max(1, area.x + area.w - (ox + box.w));
          area.x = ox + box.w;
        } else if (legend.pos === 't') {
          area.h = Math.max(1, area.y + area.h - (oy + box.h));
          area.y = oy + box.h;
        } else if (legend.pos === 'b') area.h = Math.max(1, oy - area.y);
        else {
          area.w = Math.max(1, ox - area.x);
          if (legend.pos === 'tr' && !pie) {
            // 右上（corner）の凡例：Excel はプロットを凡例の下（凡例の下端 + marginPt）まで縮める。縮めた後のプロットの高さが
            // 縮める前の legendCornerShrinkMin 以上残らない小さい面では縮めず、凡例とプロットを legendEdgeGapPt だけ下げる（C5）
            const bottom = area.y + area.h;
            const below = oy + box.h - gap;
            const kept = (bottom - below - pad - margin) / Math.max(1, area.h - pad - margin);
            if (kept >= L.legendCornerShrinkMin) {
              area.h = Math.max(1, bottom - below);
              area.y = below;
            } else {
              oy += gap;
              area.y += gap;
              area.h = Math.max(1, area.h - gap);
            }
          }
        }
      }
      const done = box;
      emitLeg = (): void => emitLegend(cx, done, look, ox, oy);
    }
  }

  // プロット
  const pl = m.plotLayout;
  if (pie) {
    const free: Rect = {
      x: area.x + pad,
      y: area.y + pad,
      w: Math.max(1, area.w - 2 * pad),
      h: Math.max(1, area.h - 2 * pad),
    };
    const plot = hasManual(pl) ? applyManual(pl, W, H, free) : free;
    emitPieArea(cx, plot);
  } else {
    const cc = cartCtx(cx);
    const free: Rect = {
      x: area.x + margin,
      y: area.y + pad,
      w: Math.max(1, area.w - margin - pad),
      h: Math.max(1, area.h - pad - margin),
    };
    const auto = solvePlot(cx, cc, free);
    let outer = free;
    let plot = auto.plot;
    let axes = auto.axes;
    if (hasManual(pl)) {
      if (pl.target === 'inner') {
        plot = applyManual(pl, W, H, auto.plot);
        axes = solveAxes(cx, cc, plot.w, plot.h);
      } else {
        outer = applyManual(pl, W, H, free);
        const s = solvePlot(cx, cc, outer);
        plot = s.plot;
        axes = s.axes;
      }
    }
    emitCartesian(cx, cc, outer, plot, axes);
  }
  emitLeg?.();
  emitTitle?.();
}

/** 円・ドーナツ：プロットの面の矩形（塗りがあれば）と扇 */
function emitPieArea(cx: Ctx, plot: Rect): void {
  const m = cx.model;
  // 暗い背景でも円・ドーナツのプロットは塗りなし（黒い面が見える。明示の c:spPr が優先）
  const pFill = resolveFill(m.plotArea.fill, null);
  const pStroke = resolveStroke(cx, m.plotArea.line, null);
  if (pFill !== null || pStroke !== null) {
    push(cx, {
      kind: 'rect',
      x: plot.x,
      y: plot.y,
      width: plot.w,
      height: plot.h,
      fill: pFill,
      stroke: pStroke,
    });
  }
  emitPie(cx, plot);
}
