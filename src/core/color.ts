// 正: InterHub docs/plan/adr/0028-xlsx-preset-shape-geometry.md（§決定 6.2〜6.5・§決定 7）
/**
 * DrawingML の色の解釈（画面に依存しない純関数）：色の要素 1 つ（＋変換の子）→ RGB の整数と不透明度、
 * 面の濃淡（darken・lighten）、2 色の混色。
 * - 色の要素は 6 種（`srgbClr`・`schemeClr`・`sysClr`・`scrgbClr`・`hslClr`・`prstClr`）。
 * - 変換は子の要素を出現順に掛ける。HSL 系は HSL で、`tint`・`shade` は線形 RGB で、`alpha` 系は不透明度。
 * - 値は整数（0〜0xFFFFFF）と不透明度（0〜1）で返し、ファイルの文字列を返さない。名前は `Map` のキーにだけ使う。
 * - 解けない・不正な値は例外にせず null（解けない）か、その変換だけ飛ばす。図形のどこから色を取るかは ISSUE-95。
 */
import { MAX_COLOR_TRANSFORMS } from './limits';
import { DEFAULT_THEME, parseHex6, sysClrValue, type Theme, type ThemeSlot } from './theme';
import { NS, parseAttrs } from './xml';

export { parseHex6, sysClrValue };

/** 色：RGB の整数（0〜0xFFFFFF）と不透明度（0〜1） */
export interface Rgba {
  rgb: number;
  alpha: number;
}

/** 見つけた色の要素 */
export interface ColorElement {
  /** 要素名（接頭辞を除く。例 `srgbClr`） */
  name: string;
  /** 開きタグ全体（`parseAttrs` に渡せる） */
  open: string;
  /** 中身（子の変換）。自己閉じ・閉じタグの無いときは null */
  body: string | null;
}

// ---- 線形化（U7：Excel の画素と合わなければ、ここの 1 か所だけ直す） ----

/**
 * sRGB の成分（0〜1）と線形 RGB を行き来する式の指定。null は sRGB の区分式（IEC 61966-2-1）。
 * 数値なら単純なべき乗（γ 2.2・2.3 など）。`tint`・`shade`・`scrgbClr` の変換が使う（U7）。
 */
export const LINEAR_GAMMA: number | null = null;

const clamp01 = (x: number): number => (x > 0 ? (x < 1 ? x : 1) : 0);

/** sRGB の成分（0〜1）→ 線形 RGB（0〜1）。範囲外は収め、NaN は 0 */
export function toLinear(c: number, gamma: number | null = LINEAR_GAMMA): number {
  const v = clamp01(c);
  if (gamma !== null) return Math.pow(v, gamma);
  return v <= 0.04045 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
}

/** 線形 RGB（0〜1）→ sRGB の成分（0〜1）。範囲外は収め、NaN は 0 */
export function fromLinear(l: number, gamma: number | null = LINEAR_GAMMA): number {
  const v = clamp01(l);
  if (gamma !== null) return Math.pow(v, 1 / gamma);
  return v <= 0.0031308 ? v * 12.92 : 1.055 * Math.pow(v, 1 / 2.4) - 0.055;
}

// ---- 面の濃淡（U8。ISSUE-95 で Excel 16 の画素（ref-colors.xlsx の can・cube・bevel の面 4 色）と照合して決めた） ----

/**
 * path の `fill` の濃淡：`shade`（黒へ寄せる）か `tint`（白へ寄せる）の係数（0〜1）。
 * Excel の画素から決めた値：darken＝0.6・lighten＝0.6 は成分ごとの比がちょうど 0.6、darkenLess・lightenLess は 0.802〜0.809 に揃うので 0.805。
 * 濃淡は `applyFillMode` が **sRGB の値のまま**掛ける（線形 RGB にすると Excel の面より暗く／明るくなりすぎる）。
 */
export const FILL_MODE_FACTORS: ReadonlyMap<string, { kind: 'shade' | 'tint'; factor: number }> =
  new Map([
    ['darken', { kind: 'shade', factor: 0.6 }],
    ['darkenLess', { kind: 'shade', factor: 0.805 }],
    ['lighten', { kind: 'tint', factor: 0.6 }],
    ['lightenLess', { kind: 'tint', factor: 0.805 }],
  ]);

// ---- 基本の数値 ----

/** 整数だけの文字列（`-` 可）を数にする。整数でない・巨大で有限でない・未指定は null */
function parseIntAttr(s: string | undefined): number | null {
  if (s === undefined || !/^-?\d+$/.test(s)) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

/** 成分（0〜1）を 0〜255 の整数にする */
const to255 = (c: number): number => Math.round(clamp01(c) * 255);
/** RGB の整数 → 成分 3 つ（0〜1） */
const unpack = (rgb: number): [number, number, number] => [
  ((rgb >> 16) & 0xff) / 255,
  ((rgb >> 8) & 0xff) / 255,
  (rgb & 0xff) / 255,
];
/** 成分 3 つ（0〜1）→ RGB の整数 */
const pack = (r: number, g: number, b: number): number =>
  (to255(r) << 16) | (to255(g) << 8) | to255(b);

// ---- HSL ----

/** RGB（0〜1）→ HSL（色相は度 0〜360・彩度・明るさは 0〜1） */
function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  const d = max - min;
  if (d === 0) return [0, 0, l];
  const s = d / (1 - Math.abs(2 * l - 1));
  let hue: number;
  if (max === r) hue = ((g - b) / d) % 6;
  else if (max === g) hue = (b - r) / d + 2;
  else hue = (r - g) / d + 4;
  return [wrapHue(hue * 60), clamp01(s), l];
}

/** 色相を 0 以上 360 未満に回す（有限でなければ 0） */
function wrapHue(deg: number): number {
  if (!Number.isFinite(deg)) return 0;
  const m = deg % 360;
  return m < 0 ? m + 360 : m;
}

/** HSL → RGB（0〜1）。色相は 360° で回し、彩度・明るさは収める */
function hslToRgb(hue: number, sat: number, lum: number): [number, number, number] {
  const s = clamp01(sat);
  const l = clamp01(lum);
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const hp = wrapHue(hue) / 60;
  const x = c * (1 - Math.abs((hp % 2) - 1));
  let r: number;
  let g: number;
  let b: number;
  if (hp < 1) [r, g, b] = [c, x, 0];
  else if (hp < 2) [r, g, b] = [x, c, 0];
  else if (hp < 3) [r, g, b] = [0, c, x];
  else if (hp < 4) [r, g, b] = [0, x, c];
  else if (hp < 5) [r, g, b] = [x, 0, c];
  else [r, g, b] = [c, 0, x];
  const m = l - c / 2;
  return [r + m, g + m, b + m];
}

// ---- プリセット色（ST_PresetColorVal） ----

/** `ST_PresetColorVal` の色（名前 → RGB）。dk/lt/med の別名（dark/light/medium）と Gray の別綴り（Grey）は下で足す */
const PRESET_BASE: readonly (readonly [string, number])[] = [
  ['aliceBlue', 0xf0f8ff],
  ['antiqueWhite', 0xfaebd7],
  ['aqua', 0x00ffff],
  ['aquamarine', 0x7fffd4],
  ['azure', 0xf0ffff],
  ['beige', 0xf5f5dc],
  ['bisque', 0xffe4c4],
  ['black', 0x000000],
  ['blanchedAlmond', 0xffebcd],
  ['blue', 0x0000ff],
  ['blueViolet', 0x8a2be2],
  ['brown', 0xa52a2a],
  ['burlyWood', 0xdeb887],
  ['cadetBlue', 0x5f9ea0],
  ['chartreuse', 0x7fff00],
  ['chocolate', 0xd2691e],
  ['coral', 0xff7f50],
  ['cornflowerBlue', 0x6495ed],
  ['cornsilk', 0xfff8dc],
  ['crimson', 0xdc143c],
  ['cyan', 0x00ffff],
  ['deepPink', 0xff1493],
  ['deepSkyBlue', 0x00bfff],
  ['dimGray', 0x696969],
  ['dkBlue', 0x00008b],
  ['dkCyan', 0x008b8b],
  ['dkGoldenrod', 0xb8860b],
  ['dkGray', 0xa9a9a9],
  ['dkGreen', 0x006400],
  ['dkKhaki', 0xbdb76b],
  ['dkMagenta', 0x8b008b],
  ['dkOliveGreen', 0x556b2f],
  ['dkOrange', 0xff8c00],
  ['dkOrchid', 0x9932cc],
  ['dkRed', 0x8b0000],
  ['dkSalmon', 0xe9967a],
  ['dkSeaGreen', 0x8fbc8b],
  ['dkSlateBlue', 0x483d8b],
  ['dkSlateGray', 0x2f4f4f],
  ['dkTurquoise', 0x00ced1],
  ['dkViolet', 0x9400d3],
  ['dodgerBlue', 0x1e90ff],
  ['firebrick', 0xb22222],
  ['floralWhite', 0xfffaf0],
  ['forestGreen', 0x228b22],
  ['fuchsia', 0xff00ff],
  ['gainsboro', 0xdcdcdc],
  ['ghostWhite', 0xf8f8ff],
  ['gold', 0xffd700],
  ['goldenrod', 0xdaa520],
  ['gray', 0x808080],
  ['green', 0x008000],
  ['greenYellow', 0xadff2f],
  ['honeydew', 0xf0fff0],
  ['hotPink', 0xff69b4],
  ['indianRed', 0xcd5c5c],
  ['indigo', 0x4b0082],
  ['ivory', 0xfffff0],
  ['khaki', 0xf0e68c],
  ['lavender', 0xe6e6fa],
  ['lavenderBlush', 0xfff0f5],
  ['lawnGreen', 0x7cfc00],
  ['lemonChiffon', 0xfffacd],
  ['lime', 0x00ff00],
  ['limeGreen', 0x32cd32],
  ['linen', 0xfaf0e6],
  ['ltBlue', 0xadd8e6],
  ['ltCoral', 0xf08080],
  ['ltCyan', 0xe0ffff],
  ['ltGoldenrodYellow', 0xfafad2],
  ['ltGray', 0xd3d3d3],
  ['ltGreen', 0x90ee90],
  ['ltPink', 0xffb6c1],
  ['ltSalmon', 0xffa07a],
  ['ltSeaGreen', 0x20b2aa],
  ['ltSkyBlue', 0x87cefa],
  ['ltSlateGray', 0x778899],
  ['ltSteelBlue', 0xb0c4de],
  ['ltYellow', 0xffffe0],
  ['magenta', 0xff00ff],
  ['maroon', 0x800000],
  ['medAquamarine', 0x66cdaa],
  ['medBlue', 0x0000cd],
  ['medOrchid', 0xba55d3],
  ['medPurple', 0x9370db],
  ['medSeaGreen', 0x3cb371],
  ['medSlateBlue', 0x7b68ee],
  ['medSpringGreen', 0x00fa9a],
  ['medTurquoise', 0x48d1cc],
  ['medVioletRed', 0xc71585],
  ['midnightBlue', 0x191970],
  ['mintCream', 0xf5fffa],
  ['mistyRose', 0xffe4e1],
  ['moccasin', 0xffe4b5],
  ['navajoWhite', 0xffdead],
  ['navy', 0x000080],
  ['oldLace', 0xfdf5e6],
  ['olive', 0x808000],
  ['oliveDrab', 0x6b8e23],
  ['orange', 0xffa500],
  ['orangeRed', 0xff4500],
  ['orchid', 0xda70d6],
  ['paleGoldenrod', 0xeee8aa],
  ['paleGreen', 0x98fb98],
  ['paleTurquoise', 0xafeeee],
  ['paleVioletRed', 0xdb7093],
  ['papayaWhip', 0xffefd5],
  ['peachPuff', 0xffdab9],
  ['peru', 0xcd853f],
  ['pink', 0xffc0cb],
  ['plum', 0xdda0dd],
  ['powderBlue', 0xb0e0e6],
  ['purple', 0x800080],
  ['red', 0xff0000],
  ['rosyBrown', 0xbc8f8f],
  ['royalBlue', 0x4169e1],
  ['saddleBrown', 0x8b4513],
  ['salmon', 0xfa8072],
  ['sandyBrown', 0xf4a460],
  ['seaGreen', 0x2e8b57],
  ['seaShell', 0xfff5ee],
  ['sienna', 0xa0522d],
  ['silver', 0xc0c0c0],
  ['skyBlue', 0x87ceeb],
  ['slateBlue', 0x6a5acd],
  ['slateGray', 0x708090],
  ['snow', 0xfffafa],
  ['springGreen', 0x00ff7f],
  ['steelBlue', 0x4682b4],
  ['tan', 0xd2b48c],
  ['teal', 0x008080],
  ['thistle', 0xd8bfd8],
  ['tomato', 0xff6347],
  ['turquoise', 0x40e0d0],
  ['violet', 0xee82ee],
  ['wheat', 0xf5deb3],
  ['white', 0xffffff],
  ['whiteSmoke', 0xf5f5f5],
  ['yellow', 0xffff00],
  ['yellowGreen', 0x9acd32],
];

/** プリセット色の表（別名を足した後）。キーは名前なので Map（`__proto__` などで汚染されない） */
const PRESET_COLORS: ReadonlyMap<string, number> = (() => {
  const m = new Map<string, number>();
  for (const [name, rgb] of PRESET_BASE) {
    m.set(name, rgb);
    // 仕様の別名：dk→dark・lt→light・med→medium、Gray→Grey
    const long = name
      .replace(/^dk(?=[A-Z])/, 'dark')
      .replace(/^lt(?=[A-Z])/, 'light')
      .replace(/^med(?=[A-Z])/, 'medium');
    m.set(long, rgb);
    for (const n of [name, long]) if (n.endsWith('Gray')) m.set(n.slice(0, -4) + 'Grey', rgb);
  }
  return m;
})();

/** プリセット色の数（別名を含む。試験用） */
export const PRESET_COLOR_COUNT: number = PRESET_COLORS.size;

/** `prstClr` の名前から RGB の整数を引く。未知の名前は null */
export function lookupPresetColor(name: string): number | null {
  return PRESET_COLORS.get(name) ?? null;
}

// ---- schemeClr の名前 ----

/** `schemeClr` の名前 → テーマの色の位置（スプレッドシートは `clrMap` を持たないため、既定の対応で固定）。`phClr`・未知の名前は無い */
const SCHEME_SLOTS: ReadonlyMap<string, ThemeSlot> = new Map<string, ThemeSlot>([
  ['dk1', 'dk1'],
  ['lt1', 'lt1'],
  ['dk2', 'dk2'],
  ['lt2', 'lt2'],
  ['accent1', 'accent1'],
  ['accent2', 'accent2'],
  ['accent3', 'accent3'],
  ['accent4', 'accent4'],
  ['accent5', 'accent5'],
  ['accent6', 'accent6'],
  ['hlink', 'hlink'],
  ['folHlink', 'folHlink'],
  ['bg1', 'lt1'],
  ['tx1', 'dk1'],
  ['bg2', 'lt2'],
  ['tx2', 'dk2'],
]);

// ---- 要素の切り出し ----

/** 色の要素 6 種の開きタグ（名前を捕まえる）。`[^<>]*>` で打ち切るので、`>` の無いタグを並べても時間は線形 */
const COLOR_OPEN_RE = new RegExp(
  String.raw`<(${NS})(srgbClr|schemeClr|sysClr|scrgbClr|hslClr|prstClr)(?=[\s>/])[^<>]*>`,
);

/** 子の要素の開きタグ（名前を捕まえる）。変換の子の列を読む */
const CHILD_OPEN_RE = new RegExp(String.raw`<${NS}([A-Za-z][\w.-]*)(?=[\s>/])[^<>]*>`, 'g');

/**
 * `xml`（例 `<a:solidFill>` の中身）で最初に見つかる色の要素を切り出す。無ければ null。
 * 閉じタグは `indexOf` で探す（閉じが無ければ body は null＝子の変換を読まない）。
 */
export function findColorElement(xml: string): ColorElement | null {
  const m = COLOR_OPEN_RE.exec(xml);
  if (m === null) return null;
  const open = m[0];
  const name = m[2]!;
  if (open.endsWith('/>')) return { name, open, body: null };
  const start = m.index + open.length;
  const close = xml.indexOf(`</${m[1] ?? ''}${name}>`, start);
  return { name, open, body: close < 0 ? null : xml.slice(start, close) };
}

// ---- 解釈 ----

/** 色の要素の基本の色（変換の前）。解けなければ null */
function baseColor(el: ColorElement, theme: Theme): number | null {
  const a = parseAttrs(el.open);
  switch (el.name) {
    case 'srgbClr':
      return parseHex6(a.val);
    case 'schemeClr': {
      const slot = a.val === undefined ? undefined : SCHEME_SLOTS.get(a.val);
      return slot === undefined ? null : (theme.get(slot) ?? DEFAULT_THEME.get(slot) ?? null);
    }
    case 'sysClr':
      return sysClrValue(a);
    case 'scrgbClr': {
      const r = parseIntAttr(a.r);
      const g = parseIntAttr(a.g);
      const b = parseIntAttr(a.b);
      if (r === null || g === null || b === null) return null;
      // 線形 RGB（1/1000 %）→ sRGB
      return pack(fromLinear(r / 100000), fromLinear(g / 100000), fromLinear(b / 100000));
    }
    case 'hslClr': {
      const hue = parseIntAttr(a.hue);
      const sat = parseIntAttr(a.sat);
      const lum = parseIntAttr(a.lum);
      if (hue === null || sat === null || lum === null) return null;
      const [r, g, b] = hslToRgb(hue / 60000, sat / 100000, lum / 100000);
      return pack(r, g, b);
    }
    case 'prstClr':
      return a.val === undefined ? null : lookupPresetColor(a.val);
    default:
      return null;
  }
}

/** 変換の途中の色（成分は 0〜1 の小数。最後に 0〜255 へ丸める） */
interface Work {
  r: number;
  g: number;
  b: number;
  alpha: number;
}

/** HSL の 1 成分の変換（lum・sat・hue の系統）。値は 1/1000 %（色相は 1/60000 度）の整数 */
function applyHsl(
  w: Work,
  kind: 'hue' | 'sat' | 'lum',
  op: 'set' | 'mod' | 'off',
  v: number,
): void {
  const hsl = rgbToHsl(w.r, w.g, w.b);
  const i = kind === 'hue' ? 0 : kind === 'sat' ? 1 : 2;
  const x = hsl[i]!;
  // 色相は度（1/60000 度）、彩度・明るさは割合（1/1000 %）。mod は割合で掛ける
  const unit = kind === 'hue' ? 60000 : 100000;
  if (op === 'set') hsl[i] = v / unit;
  else if (op === 'mod') hsl[i] = (x * v) / 100000;
  else hsl[i] = x + v / unit;
  if (i === 0) hsl[0] = wrapHue(hsl[0]);
  else hsl[i] = clamp01(hsl[i]!);
  [w.r, w.g, w.b] = hslToRgb(hsl[0], hsl[1], hsl[2]);
}

/** 線形 RGB で `c′ = c × shade`（黒へ寄せる）。係数は 0 以上 */
function applyShade(w: Work, factor: number): void {
  const k = Math.max(0, factor);
  w.r = fromLinear(toLinear(w.r) * k);
  w.g = fromLinear(toLinear(w.g) * k);
  w.b = fromLinear(toLinear(w.b) * k);
}

/** 線形 RGB で `c′ = 1 − (1 − c) × tint`（白へ寄せる）。係数は 0 以上 */
function applyTint(w: Work, factor: number): void {
  const k = Math.max(0, factor);
  w.r = fromLinear(1 - (1 - toLinear(w.r)) * k);
  w.g = fromLinear(1 - (1 - toLinear(w.g)) * k);
  w.b = fromLinear(1 - (1 - toLinear(w.b)) * k);
}

/**
 * 基本の色に、子の変換を出現順に掛ける（ADR-0028 §決定 6.3）。
 * 読む子は最初の `MAX_COLOR_TRANSFORMS` 個まで（無視する要素・不正な値も 1 個と数える）。
 * 各段の結果は 0〜1 に収める。値が整数でない変換は飛ばす。無視する変換（`comp`・`inv`・`gray`・`gamma`・`invGamma`・`red`/`green`/`blue` の系統）は色を変えない。
 */
export function applyTransforms(base: Rgba, body: string | null): Rgba {
  if (body === null) return base;
  const [r, g, b] = unpack(base.rgb);
  const w: Work = { r, g, b, alpha: base.alpha };
  CHILD_OPEN_RE.lastIndex = 0;
  for (let count = 0; count < MAX_COLOR_TRANSFORMS; count++) {
    const m = CHILD_OPEN_RE.exec(body);
    if (m === null) break;
    const v = parseIntAttr(parseAttrs(m[0]).val);
    if (v === null) continue;
    switch (m[1]) {
      case 'lum':
        applyHsl(w, 'lum', 'set', v);
        break;
      case 'lumMod':
        applyHsl(w, 'lum', 'mod', v);
        break;
      case 'lumOff':
        applyHsl(w, 'lum', 'off', v);
        break;
      case 'sat':
        applyHsl(w, 'sat', 'set', v);
        break;
      case 'satMod':
        applyHsl(w, 'sat', 'mod', v);
        break;
      case 'satOff':
        applyHsl(w, 'sat', 'off', v);
        break;
      case 'hue':
        applyHsl(w, 'hue', 'set', v);
        break;
      case 'hueMod':
        applyHsl(w, 'hue', 'mod', v);
        break;
      case 'hueOff':
        applyHsl(w, 'hue', 'off', v);
        break;
      case 'tint':
        applyTint(w, v / 100000);
        break;
      case 'shade':
        applyShade(w, v / 100000);
        break;
      case 'alpha':
        w.alpha = clamp01(v / 100000);
        break;
      case 'alphaMod':
        w.alpha = clamp01((w.alpha * v) / 100000);
        break;
      case 'alphaOff':
        w.alpha = clamp01(w.alpha + v / 100000);
        break;
      default:
        break; // 無視する変換（comp・inv・gray・gamma・invGamma・red・green・blue の系統・未知）
    }
  }
  CHILD_OPEN_RE.lastIndex = 0;
  return { rgb: pack(w.r, w.g, w.b), alpha: clamp01(w.alpha) };
}

/**
 * 切り出した色の要素 1 つを解く（基本の色 → 子の変換）。解けなければ null（例外にしない）。
 * 解けないときに次の出どころへ進むのは呼び出し側（ISSUE-95）。
 */
export function resolveColorElement(el: ColorElement, theme: Theme): Rgba | null {
  const rgb = baseColor(el, theme);
  if (rgb === null) return null;
  return applyTransforms({ rgb, alpha: 1 }, el.body);
}

/**
 * `xml`（`<a:solidFill>` の中身など）の最初の色の要素を解く。色の要素が無い・解けなければ null。
 * 色は RGB の整数と不透明度で返す（ファイルの文字列は返さない）。
 */
export function resolveColor(xml: string, theme: Theme): Rgba | null {
  const el = findColorElement(xml);
  return el === null ? null : resolveColorElement(el, theme);
}

// ---- 面の濃淡・混色 ----

/**
 * path の `fill`（`norm`・`darken`・`darkenLess`・`lighten`・`lightenLess`）を色に掛ける。
 * 係数は `FILL_MODE_FACTORS`（U8）。未知の名前・`norm` は色のまま。不透明度は変えない。
 * 掛けるのは sRGB の値のまま（shade＝`c × k`・tint＝`1 − (1 − c) × k`）。変換の `shade`・`tint`（線形 RGB）とは別の式で、Excel の面の画素と合う方を採った。
 */
export function applyFillMode(color: Rgba, mode: string): Rgba {
  const f = FILL_MODE_FACTORS.get(mode);
  if (f === undefined) return color;
  const k = Math.max(0, f.factor);
  const [r, g, b] = unpack(color.rgb);
  const one = (c: number): number => (f.kind === 'shade' ? c * k : 1 - (1 - c) * k);
  return { rgb: pack(one(r), one(g), one(b)), alpha: color.alpha };
}

/**
 * 2 色を **線形 RGB** で割合を混ぜる（`t`＝0 で a、1 で b。0〜1 に収め、NaN は 0）。不透明度は sRGB と同じ線形補間。
 * Excel の 2 stop のグラデーションの中央の色がこの混色と合う（ISSUE-95・U7。3 stop 以上は sRGB の混色 `mixColors` と合う）。
 */
export function mixColorsLinear(a: Rgba, b: Rgba, t: number): Rgba {
  const k = clamp01(t);
  const [ar, ag, ab] = unpack(a.rgb).map((c) => toLinear(c)) as [number, number, number];
  const [br, bg, bb] = unpack(b.rgb).map((c) => toLinear(c)) as [number, number, number];
  return {
    rgb: pack(
      fromLinear(ar + (br - ar) * k),
      fromLinear(ag + (bg - ag) * k),
      fromLinear(ab + (bb - ab) * k),
    ),
    alpha: clamp01(a.alpha + (b.alpha - a.alpha) * k),
  };
}

/**
 * 2 色を割合で混ぜる（`t`＝0 で a、1 で b。0〜1 に収め、NaN は 0）。各成分は sRGB の値の線形補間で、不透明度も混ぜる。
 * グラデーションの補間・模様の混色に使う（適用は ISSUE-95）。
 */
export function mixColors(a: Rgba, b: Rgba, t: number): Rgba {
  const k = clamp01(t);
  const [ar, ag, ab] = unpack(a.rgb);
  const [br, bg, bb] = unpack(b.rgb);
  return {
    rgb: pack(ar + (br - ar) * k, ag + (bg - ag) * k, ab + (bb - ab) * k),
    alpha: clamp01(a.alpha + (b.alpha - a.alpha) * k),
  };
}
