// 正: InterHub docs/plan/adr/0029-xlsx-cell-borders.md（§決定 6）
/**
 * SpreadsheetML の色（styles.xml の `<color auto rgb theme tint indexed>`）の解釈（画面に依存しない純関数）。
 * DrawingML の色（color.ts・ADR-0028 §決定 6）とは書き方も `tint` の式も別なので、別の小さな解釈器にする。
 * 罫線（borders.ts）のほか、塗り・フォントの色（次の候補）でも使い回す。
 * - `auto="1"`・属性なし・`indexed="64"` は null（自動。描く側が既定の色にする）。解けない・不正な値も null。例外にしない。
 * - 値は整数（0〜0xFFFFFF）と不透明度で返し、ファイルの文字列を返さない。
 */
import type { Rgba } from './color';
import { type Theme, type ThemeSlot, parseHex6 } from './theme';
import { type Attrs, findElement, forEachElement, parseAttrs } from './xml';

/** 旧 64 色の表（ECMA-376 §18.8.27。`<indexedColors>` が無いブックの `indexed` 0〜63） */
export const INDEXED_COLORS: readonly number[] = [
  0x000000, 0xffffff, 0xff0000, 0x00ff00, 0x0000ff, 0xffff00, 0xff00ff, 0x00ffff, 0x000000,
  0xffffff, 0xff0000, 0x00ff00, 0x0000ff, 0xffff00, 0xff00ff, 0x00ffff, 0x800000, 0x008000,
  0x000080, 0x808000, 0x800080, 0x008080, 0xc0c0c0, 0x808080, 0x9999ff, 0x993366, 0xffffcc,
  0xccffff, 0x660066, 0xff8080, 0x0066cc, 0xccccff, 0x000080, 0xff00ff, 0xffff00, 0x00ffff,
  0x800080, 0x800000, 0x008080, 0x0000ff, 0x00ccff, 0xccffff, 0xccffcc, 0xffff99, 0x99ccff,
  0xff99cc, 0xcc99ff, 0xffcc99, 0x3366ff, 0x33cccc, 0x99cc00, 0xffcc00, 0xff9900, 0xff6600,
  0x666699, 0x969696, 0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399,
  0x333333,
];

/** `indexed` の特別な値：64＝システムの前景色（自動）、65＝システムの背景色（白） */
const INDEXED_FOREGROUND = 64;
const INDEXED_BACKGROUND = 65;

/** `<indexedColors>` の件数の上限（旧パレットは 64 色。ADR-0029 §決定 9） */
const MAX_INDEXED = INDEXED_COLORS.length;

/** `theme` の番号（SpreadsheetML の順）→ テーマの位置。0 と 1・2 と 3 が THEME_SLOTS（dk1・lt1・dk2・lt2）と入れ替わる */
const THEME_BY_INDEX: readonly ThemeSlot[] = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];

/** 属性の値を整数にする（空・数でない・小数は null） */
function intAttr(v: string | undefined): number | null {
  if (v === undefined || v.trim() === '') return null;
  const n = Number(v);
  return Number.isInteger(n) ? n : null;
}

/** `rgb` 属性（ARGB 8 桁か RGB 6 桁の 16 進）を整数にする。先頭のアルファは無視。不正は null */
function parseRgbAttr(raw: string): number | null {
  if (raw.length === 8 && parseHex6(raw.slice(0, 6)) !== null) return parseHex6(raw.slice(2));
  if (raw.length === 6) return parseHex6(raw);
  return null;
}

/**
 * styles.xml の `<colors><indexedColors>` を読み、64 色の表を返す（既定の表を、書いてある位置だけ上書きする）。
 * 無い・空なら null（既定の表を使う）。`<rgbColor rgb>` は 64 件まで。不正な値はその位置だけ既定のまま。
 */
export function parseIndexedColors(xml: string): readonly number[] | null {
  const block = findElement(xml, 'indexedColors');
  if (block === null || block.body === null) return null;
  const palette = INDEXED_COLORS.slice();
  let n = 0;
  forEachElement(block.body, 'rgbColor', (open) => {
    if (n >= MAX_INDEXED) return;
    const raw = parseAttrs(open).rgb;
    const rgb = raw === undefined ? null : parseRgbAttr(raw);
    if (rgb !== null) palette[n] = rgb;
    n++;
  });
  return n === 0 ? null : palette;
}

/** RGB の整数 → HLS（H・L・S とも 0〜1） */
function rgbToHls(rgb: number): [number, number, number] {
  const r = ((rgb >> 16) & 255) / 255;
  const g = ((rgb >> 8) & 255) / 255;
  const b = (rgb & 255) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, l, 0];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, l, s];
}

/** HLS（0〜1）→ RGB の整数（各成分は四捨五入） */
function hlsToRgb(h: number, l: number, s: number): number {
  const channel = (t0: number, p: number, q: number): number => {
    let t = t0;
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r = l;
  let g = l;
  let b = l;
  if (s !== 0) {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = channel(h + 1 / 3, p, q);
    g = channel(h, p, q);
    b = channel(h - 1 / 3, p, q);
  }
  const to8 = (v: number): number => Math.min(255, Math.max(0, Math.round(v * 255)));
  return (to8(r) << 16) | (to8(g) << 8) | to8(b);
}

/**
 * SpreadsheetML の `tint`（HLS の明度）を当てる。負は `L × (1 + tint)`、正は `L × (1 − tint) + tint`。
 * `tint` は -1〜1 に丸める（-1 は黒・1 は白）。0 は色を変えない。DrawingML の `tint`（線形 RGB）とは別の式。
 */
function applyTint(rgb: number, tint: number): number {
  const t = Math.max(-1, Math.min(1, tint));
  if (t === 0) return rgb;
  const [h, l, s] = rgbToHls(rgb);
  const l2 = t < 0 ? l * (1 + t) : l * (1 - t) + t;
  return hlsToRgb(h, l2, s);
}

/**
 * `<color>` の属性 → 色。null は自動（`auto="1"`・属性なし・`indexed="64"`）か解けない値。
 * @param attrs `parseAttrs` で読んだ `<color>` の属性（無ければ undefined＝自動）
 * @param theme ブックのテーマ（`theme` の番号を解く）
 * @param palette `parseIndexedColors` の結果（null なら既定の 64 色）
 */
export function resolveCellColor(
  attrs: Attrs | undefined,
  theme: Theme,
  palette: readonly number[] | null,
): Rgba | null {
  if (attrs === undefined) return null;
  if (attrs.auto === '1' || attrs.auto === 'true') return null;

  let rgb: number | null = null;
  if (attrs.rgb !== undefined) {
    rgb = parseRgbAttr(attrs.rgb);
  } else if (attrs.indexed !== undefined) {
    const i = intAttr(attrs.indexed);
    if (i === INDEXED_BACKGROUND) rgb = 0xffffff;
    else if (i !== null && i >= 0 && i < MAX_INDEXED && i !== INDEXED_FOREGROUND)
      rgb = (palette ?? INDEXED_COLORS)[i] ?? null;
  } else if (attrs.theme !== undefined) {
    const i = intAttr(attrs.theme);
    const slot = i === null || i < 0 ? undefined : THEME_BY_INDEX[i];
    if (slot !== undefined) rgb = theme.get(slot) ?? null;
  }
  if (rgb === null) return null;

  if (attrs.tint !== undefined && attrs.tint.trim() !== '') {
    const tint = Number(attrs.tint);
    if (Number.isFinite(tint)) rgb = applyTint(rgb, tint);
  }
  return { rgb, alpha: 1 };
}
