// 正: InterHub docs/plan/adr/0030-xlsx-cell-format.md（§決定 2・3・4・11）
/**
 * styles.xml の書式の読み取り：`<fills>`（塗り）・`<fonts>`（文字の装飾）と、cellXfs の `fontId`・`fillId`。
 * ブックで 1 回だけ読み、セルの `s` 属性（cellXfs の位置）→ 書式の番号 → 塗りと文字の装飾、と引けるようにする（`FormatTable`）。
 * - 見える書式（塗り・太字・斜体・下線・取り消し線・自動でない色・比が 1 でない文字サイズのどれか）が無い xf は「書式なし」にまとめる。
 * - 同じ見た目の書式は 1 つにまとめる（キーは数値だけ。ファイルの文字列は持たない）。まとめた後の数は `maxCellFormats` まで。
 * - 塗りは 1 色に畳む：`solid`＝`fgColor`、ほかの模様＝前景と背景の混色、グラデーション＝50% の位置の色。
 * - `applyFill`・`applyFont` は見ない（F5）。cellStyleXfs は読まない。フォント名は読まない。
 * - 色は SpreadsheetML の `<color>`（cell-color.ts）をテーマで解いてから持つ。解けない値は例外にせず既定の色・書式なしに倒す。
 * - 走査は線形（開きタグは `[^<>]*>`、閉じは `indexOf`）。
 */
import { type Rgba, mixColors, mixColorsLinear } from './color';
import { parseIndexedColors, resolveCellColor } from './cell-color';
import { MAX_GRADIENT_STOPS, type XlsxLimits } from './limits';
import type { Theme } from './theme';
import type { CellFormat } from './types';
import { NS, findElement, findOpenTag, forEachElement, parseAttrs } from './xml';

/** 書式の表 */
export interface FormatTable {
  /** 書式の番号 → 書式（見える書式だけ・重複なし。`SheetRow.formats` の値・`XlsxSheet.formats` の位置） */
  readonly formats: readonly CellFormat[];
  /** cellXfs の位置（セルの `s`）→ 書式の番号。書式なし・不正・上限の外は -1 */
  readonly xfFormat: Int32Array;
}

/** 書式の無い表（styles.xml が無い・壊れているとき。共有する・変更しない） */
export const EMPTY_FORMAT_TABLE: FormatTable = Object.freeze({
  formats: Object.freeze([] as CellFormat[]),
  xfFormat: new Int32Array(0),
});

/** 整数の属性（空・数でない・小数は null） */
function intAttr(v: string | undefined): number | null {
  if (v === undefined || v.trim() === '') return null;
  const n = Number(v);
  return Number.isInteger(n) ? n : null;
}

/** 数の属性（空・数でない・無限は null） */
function numAttr(v: string | undefined): number | null {
  if (v === undefined || v.trim() === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

/** 模様（`patternType`）→ 前景色の割合（ADR-0030 §決定 3 の表。`solid`・`none`・未知は含めない） */
const PATTERN_RATIO: ReadonlyMap<string, number> = new Map<string, number>([
  ['darkGray', 0.75],
  ['mediumGray', 0.5],
  ['lightGray', 0.25],
  ['gray125', 0.125],
  ['gray0625', 0.0625],
  ['darkHorizontal', 0.5],
  ['darkVertical', 0.5],
  ['darkDown', 0.5],
  ['darkUp', 0.5],
  ['darkGrid', 0.75],
  ['darkTrellis', 0.75],
  ['lightHorizontal', 0.25],
  ['lightVertical', 0.25],
  ['lightDown', 0.25],
  ['lightUp', 0.25],
  ['lightGrid', 7 / 16],
  ['lightTrellis', 6 / 16],
]);

/** 色の既定：前景・`solid` の塗りは黒、背景は白（F3） */
const BLACK: Rgba = { rgb: 0x000000, alpha: 1 };
const WHITE: Rgba = { rgb: 0xffffff, alpha: 1 };

/** `<fgColor>`・`<bgColor>`（要素名 name）の色。要素が無い・自動・解けない色なら fallback */
function patternColor(
  body: string | null,
  name: string,
  fallback: Rgba,
  theme: Theme,
  palette: readonly number[] | null,
): Rgba {
  const tag = body === null ? null : findOpenTag(body, name);
  return (tag === null ? null : resolveCellColor(parseAttrs(tag), theme, palette)) ?? fallback;
}

/** `<patternFill>` 1 件（開きタグと中身）→ 塗りの色。塗りなし（無い・none・未知の模様）は undefined */
function readPatternFill(
  open: string,
  body: string | null,
  theme: Theme,
  palette: readonly number[] | null,
): Rgba | undefined {
  const type = parseAttrs(open).patternType;
  if (type === undefined) return undefined;
  const fg = patternColor(body, 'fgColor', BLACK, theme, palette);
  if (type === 'solid') return { rgb: fg.rgb, alpha: 1 };
  const ratio = PATTERN_RATIO.get(type);
  if (ratio === undefined) return undefined;
  const bg = patternColor(body, 'bgColor', WHITE, theme, palette);
  return { rgb: mixColors(bg, fg, ratio).rgb, alpha: 1 };
}

/**
 * `<gradientFill>` の中身 → 50% の位置の色。`<stop>`（16 個まで）を `position` の順に並べ、前後の stop の補間。
 * 2 stop は線形 RGB・3 stop 以上は sRGB（ADR-0028 §決定 6.4 と同じ）。stop が無ければ undefined。解けない色は黒。
 */
function readGradientFill(
  body: string | null,
  theme: Theme,
  palette: readonly number[] | null,
): Rgba | undefined {
  if (body === null) return undefined;
  const stops: { pos: number; color: Rgba }[] = [];
  let seen = 0;
  forEachElement(body, 'stop', (open, inner) => {
    if (seen >= MAX_GRADIENT_STOPS) return;
    seen++;
    const pos = Math.min(Math.max(numAttr(parseAttrs(open).position) ?? 0, 0), 1);
    stops.push({ pos, color: patternColor(inner, 'color', BLACK, theme, palette) });
  });
  if (stops.length === 0) return undefined;
  stops.sort((a, b) => a.pos - b.pos);
  const mid = 0.5;
  let color = stops[stops.length - 1]!.color;
  if (stops[0]!.pos >= mid) {
    color = stops[0]!.color;
  } else {
    for (let i = 1; i < stops.length; i++) {
      const hi = stops[i]!;
      if (hi.pos < mid) continue;
      const lo = stops[i - 1]!;
      const t = (mid - lo.pos) / (hi.pos - lo.pos);
      color =
        stops.length === 2
          ? mixColorsLinear(lo.color, hi.color, t)
          : mixColors(lo.color, hi.color, t);
      break;
    }
  }
  return { rgb: color.rgb, alpha: 1 };
}

/** `<fill>` 1 件の中身 → 塗りの色（`patternFill` か `gradientFill`。どちらも無ければ undefined） */
function readFill(
  body: string | null,
  theme: Theme,
  palette: readonly number[] | null,
): Rgba | undefined {
  if (body === null) return undefined;
  const pattern = findElement(body, 'patternFill');
  if (pattern !== null) return readPatternFill(pattern.open, pattern.body, theme, palette);
  const gradient = findElement(body, 'gradientFill');
  return gradient === null ? undefined : readGradientFill(gradient.body, theme, palette);
}

/** 文字の装飾（`<font>` 1 件から読んだもの。塗りは含まない） */
interface FontStyle {
  color?: Rgba;
  bold?: true;
  italic?: true;
  strike?: true;
  underline?: 'single' | 'double';
  sizeRatio?: number;
}

/** 真偽の子要素（`<b>`・`<i>`・`<strike>`）。要素が無ければ偽、`val` が無ければ真、`val="0"`・`false` は偽 */
function boolChild(body: string, name: string): boolean {
  const tag = findOpenTag(body, name);
  if (tag === null) return false;
  const v = parseAttrs(tag).val;
  return v !== '0' && v !== 'false';
}

/** 下線の種類（`<u val>`）の対応。`none`・未知は含めない（下線なし）。会計用は普通の下線・二重下線に寄せる（F8） */
const UNDERLINE: ReadonlyMap<string, 'single' | 'double'> = new Map<string, 'single' | 'double'>([
  ['single', 'single'],
  ['singleAccounting', 'single'],
  ['double', 'double'],
  ['doubleAccounting', 'double'],
]);

/** 文字サイズの比の範囲（ADR-0030 §決定 4） */
const MIN_SIZE_RATIO = 0.5;
const MAX_SIZE_RATIO = 4;

/** `<font>` の `<sz val>`（pt）。無い・数でない・0 以下は null */
function fontSize(body: string): number | null {
  const tag = findOpenTag(body, 'sz');
  const v = tag === null ? null : numAttr(parseAttrs(tag).val);
  return v !== null && v > 0 ? v : null;
}

/**
 * `<font>` 1 件の中身 → 文字の装飾。
 * @param baseSize 先頭の `<font>` の `sz`（pt。無ければ null＝比を持たない）
 */
function readFont(
  body: string | null,
  baseSize: number | null,
  theme: Theme,
  palette: readonly number[] | null,
): FontStyle {
  const font: FontStyle = {};
  if (body === null) return font;
  if (boolChild(body, 'b')) font.bold = true;
  if (boolChild(body, 'i')) font.italic = true;
  if (boolChild(body, 'strike')) font.strike = true;
  const u = findOpenTag(body, 'u');
  if (u !== null) {
    const kind = UNDERLINE.get(parseAttrs(u).val ?? 'single');
    if (kind !== undefined) font.underline = kind;
  }
  const size = fontSize(body);
  if (size !== null && baseSize !== null) {
    const ratio =
      Math.round(Math.min(Math.max(size / baseSize, MIN_SIZE_RATIO), MAX_SIZE_RATIO) * 100) / 100;
    if (ratio !== 1) font.sizeRatio = ratio;
  }
  const colorTag = findOpenTag(body, 'color');
  const color = colorTag === null ? null : resolveCellColor(parseAttrs(colorTag), theme, palette);
  // 不透明な #000000 は自動と同じ描き方なので、自動に寄せる（書式の数を減らす。ADR-0030 §決定 2）
  if (color !== null && color.rgb !== 0x000000) font.color = color;
  return font;
}

/**
 * styles.xml から書式の表を作る（ブックで 1 回）。
 * @param xml styles.xml（無ければ null＝書式なし）
 * @param theme ブックのテーマ（色の `theme` を解く）
 * @param limits `maxFonts`・`maxFills` を超えた要素は読まない（それを参照するセルは装飾なし・塗りなし）。
 *   重複を除いた書式の数が `maxCellFormats` を超えた書式のセルは書式なし
 */
export function parseCellFormats(
  xml: string | null,
  theme: Theme,
  limits: XlsxLimits,
): FormatTable {
  if (xml === null) return EMPTY_FORMAT_TABLE;
  const cellXfs = findElement(xml, 'cellXfs');
  if (cellXfs === null || cellXfs.body === null) return EMPTY_FORMAT_TABLE;

  const palette = parseIndexedColors(xml);

  // <fill> の位置（fillId）→ 塗りの色（塗りなしは undefined）
  const fills: (Rgba | undefined)[] = [];
  const fillBlock = findElement(xml, 'fills');
  if (fillBlock?.body)
    forEachElement(fillBlock.body, 'fill', (_open, body) => {
      if (fills.length >= limits.maxFills) return;
      fills.push(readFill(body, theme, palette));
    });

  // <font> の位置（fontId）→ 文字の装飾（文字サイズの比は、先頭の <font> の sz が基準）
  const fonts: FontStyle[] = [];
  const fontBlock = findElement(xml, 'fonts');
  if (fontBlock?.body) {
    let baseSize: number | null = null;
    forEachElement(fontBlock.body, 'font', (_open, body) => {
      if (fonts.length >= limits.maxFonts) return;
      if (fonts.length === 0) baseSize = body === null ? null : fontSize(body);
      fonts.push(readFont(body, baseSize, theme, palette));
    });
  }

  const formats: CellFormat[] = [];
  const keyToIndex = new Map<string, number>();
  const xfFormat: number[] = [];
  for (const m of cellXfs.body.matchAll(new RegExp(String.raw`<${NS}xf\b[^<>]*>`, 'g'))) {
    const a = parseAttrs(m[0]);
    const fontId = intAttr(a.fontId);
    const fillId = intAttr(a.fillId);
    const font = fontId !== null && fontId >= 0 ? fonts[fontId] : undefined;
    const fill = fillId !== null && fillId >= 0 ? fills[fillId] : undefined;
    const format: CellFormat = {};
    if (fill) format.fill = fill;
    if (font?.color) format.color = font.color;
    if (font?.bold) format.bold = true;
    if (font?.italic) format.italic = true;
    if (font?.strike) format.strike = true;
    if (font?.underline) format.underline = font.underline;
    if (font?.sizeRatio !== undefined) format.sizeRatio = font.sizeRatio;
    if (Object.keys(format).length === 0) {
      xfFormat.push(-1);
      continue;
    }
    const key = formatKey(format);
    let idx = keyToIndex.get(key);
    if (idx === undefined) {
      if (formats.length >= limits.maxCellFormats) {
        xfFormat.push(-1);
        continue;
      }
      idx = formats.length;
      formats.push(format);
      keyToIndex.set(key, idx);
    }
    xfFormat.push(idx);
  }
  return { formats, xfFormat: Int32Array.from(xfFormat) };
}

/** 書式の重複の除去のキー（数値だけで組む：塗り｜文字の色｜装飾のビット｜サイズ比×100） */
function formatKey(f: CellFormat): string {
  const bits =
    (f.bold ? 1 : 0) |
    (f.italic ? 2 : 0) |
    (f.strike ? 4 : 0) |
    (f.underline === 'single' ? 8 : 0) |
    (f.underline === 'double' ? 16 : 0);
  return `${f.fill?.rgb ?? -1}|${f.color?.rgb ?? -1}|${bits}|${Math.round((f.sizeRatio ?? 1) * 100)}`;
}

/**
 * 「紙」の塗りか：解決後の塗りが不透明な `#FFFFFF`（`rgb`・`theme`・`indexed` など出どころは問わない）。
 * 紙は塗りなしと同じに描く（ADR-0030 1.1）ため、表の範囲を広げる判定（sheet.ts）と描き分け（viewer）の両方がこれを使う。
 * `#FEFEFE`・`#F2F2F2`・半透明の白は紙ではない（alpha < 1）。
 */
export function isPaperFill(fill: Rgba | undefined): boolean {
  return fill !== undefined && fill.rgb === 0xffffff && fill.alpha >= 1;
}

/**
 * 書式の色に `theme` の指定があるか（`<fonts>`・`<fills>`・`<borders>` の中だけを見る）。無ければテーマを読まなくても
 * 色は決まるので、呼び出し側（index.ts）はテーマの読み取りを省ける（図の無いブックでテーマの読み取りを増やさない。ISSUE-94）。
 */
export function stylesUseTheme(xml: string | null): boolean {
  if (xml === null) return false;
  for (const name of ['fonts', 'fills', 'borders']) {
    const body = findElement(xml, name)?.body;
    if (body !== null && body !== undefined && /\stheme\s*=/.test(body)) return true;
  }
  return false;
}
