// 正: InterHub docs/plan/adr/0028-xlsx-preset-shape-geometry.md（§決定 6.1・6.4）
/**
 * 図形・線・文字の色の読み取り（画面に依存しない）：色の出どころの順と、塗りの代表色。
 * - 塗り：`spPr` 直下の塗り → `a:fillRef` の子の色 → テーマの `lt1`。
 * - 枠線・線：`a:ln` の塗り → `a:lnRef` の子の色 → テーマの `tx1`。
 * - 文字：ランの `a:rPr` の塗り → `a:fontRef` の子の色 → テーマの `tx1`。
 * - 解けない色（不正な値・`phClr`・色の要素が無い）はその出どころを飛ばして次へ進む。`a:grpFill` は未指定として扱う。
 * - グラデーション・模様・画像の塗りは代表色の 1 色にする（`a:gradFill`＝50% の位置・`a:pattFill`＝前景と背景の混色・`a:blipFill`＝#D9D9D9）。
 * 塗り・枠線の「有無」は `drawing-shape.ts`（ADR-0027）が決める。ここは「ある」と決まった塗り・線の色を返す。
 * 色の要素そのものの解釈（変換・テーマ）は `color.ts`。結果は RGB の整数と不透明度だけで、ファイルの文字列を持たない。
 * 閉じの無い要素・大量の stop・変換の列でも例外にせず、走査は線形時間（上限は `limits.ts`）。
 */
import { mixColors, mixColorsLinear, resolveColor, type Rgba } from './color';
import { intAttr, type FillEl, type ShapeProps, type StyleRef } from './drawing-shape';
import { MAX_GRADIENT_STOPS } from './limits';
import { DEFAULT_THEME, type Theme, type ThemeSlot } from './theme';
import { NS, findElement, forEachElement, parseAttrs } from './xml';

/** 画像の塗り（`a:blipFill`）の代表色（画像は読まない） */
const BLIP_COLOR = 0xd9d9d9;

/** グラデーションの代表色にする位置（1/1000 %＝50%） */
const GRADIENT_MID = 50000;

/** 模様の名前（`pct5`〜`pct90`）→ 前景色の割合。この表に無い模様は 50% */
const PATTERN_RATIO: ReadonlyMap<string, number> = new Map(
  [5, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 90].map((n) => [`pct${n}`, n / 100]),
);

/** テーマの色の位置の色（テーマに無ければ既定のテーマ）。毎回新しい色を返す */
function themeColor(theme: Theme, slot: ThemeSlot): Rgba {
  return { rgb: theme.get(slot) ?? DEFAULT_THEME.get(slot) ?? 0, alpha: 1 };
}

/**
 * `a:gradFill` の代表色：`a:gsLst` の stop（16 個まで）を `pos` の順に並べ、50% の位置の色（前後の stop の補間）。解ける stop が無ければ null。
 * 補間は、解けた stop が 2 つなら線形 RGB、3 つ以上なら sRGB の値（Excel の画素に合わせた。ISSUE-95・U7：2 stop の赤→青の中央が #BC00B8、3 stop の黒→赤→白の 50% の G が #46＝sRGB の補間）。
 */
function gradientColor(body: string, theme: Theme): Rgba | null {
  const list = findElement(body, 'gsLst')?.body;
  if (list === null || list === undefined) return null;
  const stops: { pos: number; color: Rgba }[] = [];
  let seen = 0;
  forEachElement(list, 'gs', (open, b) => {
    if (seen >= MAX_GRADIENT_STOPS) return;
    seen++;
    if (b === null) return;
    const color = resolveColor(b, theme);
    if (color === null) return;
    const pos = intAttr(parseAttrs(open).pos) ?? 0;
    stops.push({ pos: Math.min(Math.max(pos, 0), 100000), color });
  });
  if (stops.length === 0) return null;
  stops.sort((a, b) => a.pos - b.pos);
  if (stops[0]!.pos >= GRADIENT_MID) return stops[0]!.color;
  for (let i = 1; i < stops.length; i++) {
    const hi = stops[i]!;
    if (hi.pos < GRADIENT_MID) continue;
    const lo = stops[i - 1]!;
    const t = (GRADIENT_MID - lo.pos) / (hi.pos - lo.pos);
    return stops.length === 2
      ? mixColorsLinear(lo.color, hi.color, t)
      : mixColors(lo.color, hi.color, t);
  }
  return stops[stops.length - 1]!.color;
}

/** `a:pattFill` の代表色：前景（`a:fgClr`）と背景（`a:bgClr`）を、前景の割合で混ぜる。片方しか解けなければその色、どちらも解けなければ null */
function patternColor(open: string, body: string, theme: Theme): Rgba | null {
  const side = (name: 'fgClr' | 'bgClr'): Rgba | null => {
    const inner = findElement(body, name)?.body;
    return inner === null || inner === undefined ? null : resolveColor(inner, theme);
  };
  const fg = side('fgClr');
  const bg = side('bgClr');
  if (fg === null || bg === null) return fg ?? bg;
  return mixColors(bg, fg, PATTERN_RATIO.get(parseAttrs(open).prst ?? '') ?? 0.5);
}

/**
 * 塗りの要素 1 つの色（ADR-0028 §決定 6.4）。`a:solidFill`＝その色・`a:gradFill`・`a:pattFill`＝代表色・`a:blipFill`＝#D9D9D9。
 * `a:noFill`・`a:grpFill`・色の要素が無い・解けないときは null（次の出どころへ進む）。
 */
export function fillElementColor(el: FillEl, theme: Theme): Rgba | null {
  if (el.name === 'blipFill') return { rgb: BLIP_COLOR, alpha: 1 };
  if (el.body === null) return null;
  switch (el.name) {
    case 'solidFill':
      return resolveColor(el.body, theme);
    case 'gradFill':
      return gradientColor(el.body, theme);
    case 'pattFill':
      return patternColor(el.open, el.body, theme);
    default:
      return null;
  }
}

/**
 * `xdr:style` の参照の子の色。idx が 1 以上（1001 以上の背景の塗りも同じ）の参照を見る。idx が 0 なら見ない。
 * idx が数値でない・欠けている（null。Excel はほぼ書かないが安全側）ときは、子の色があればそれを使う。参照が無い・子が無い・解けないときは null
 */
function refColor(ref: StyleRef | null, theme: Theme): Rgba | null {
  if (ref === null || ref.body === null || (ref.idx !== null && ref.idx < 1)) return null;
  return resolveColor(ref.body, theme);
}

/**
 * 図形の塗りの色：`spPr` 直下の塗り → `a:fillRef` の子の色 → テーマの `lt1`。
 * 塗りがあると決まった図形（`props.hasFill`）にだけ呼ぶ（有無はここで判定しない）。
 * `xdr:style` も `spPr` の塗りも無い図形は、有無の規則（ADR-0027）で「塗り・線あり」なので、色は既定の `lt1`（枠線・線は `tx1`）になる。ISSUE-95 の「既定」どおり。
 */
export function readFillColor(props: ShapeProps, theme: Theme): Rgba {
  const direct = props.fillEl === null ? null : fillElementColor(props.fillEl, theme);
  return direct ?? refColor(props.fillRef, theme) ?? themeColor(theme, 'lt1');
}

/**
 * 図形の枠線・線の色：`a:ln` の `a:solidFill`・`a:gradFill`・`a:pattFill` → `a:lnRef` の子の色 → テーマの `tx1`。
 * 枠線が見えると決まったもの（`props.hasOutline`）にだけ呼ぶ。矢じりも同じ色。
 */
export function readOutlineColor(props: ShapeProps, theme: Theme): Rgba {
  const el = props.lnFillEl;
  const direct = el === null || el.name === 'blipFill' ? null : fillElementColor(el, theme);
  return direct ?? refColor(props.lnRef, theme) ?? themeColor(theme, 'dk1');
}

/** 図形の文字の既定の色：`xdr:style` の `a:fontRef` の子の色 → テーマの `tx1`（`spBody` は `xdr:sp` の中身） */
export function readFontColor(spBody: string, theme: Theme): Rgba {
  const body = findElement(spBody, 'fontRef')?.body;
  const c = body === null || body === undefined ? null : resolveColor(body, theme);
  return c ?? themeColor(theme, 'dk1');
}

/** 塗りの要素の名前（`a:rPr` の直下で最初に出るもの） */
const TEXT_FILL_NAMES: ReadonlySet<string> = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
]);

/** 開き・閉じタグ 1 つ（名前空間の接頭辞可。`[^<>]*>` で打ち切る）。`a:rPr` の直下の塗りを探す深さの計算用 */
const TAG_RE = new RegExp(String.raw`<(/?)${NS}([A-Za-z_][\w.-]*)(?=[\s>/])[^<>]*>`, 'g');

/**
 * `a:rPr` の中身の直下（深さ 0）で最初に出る塗りの要素。無ければ null。
 * `a:ln`（文字の輪郭）・`a:uFill`（下線の色）の中の塗りを文字の色にしないよう、深さを開き・閉じタグの数で数える（線形時間）。
 */
function firstDirectFill(rPrBody: string): FillEl | null {
  let depth = 0;
  TAG_RE.lastIndex = 0;
  for (let m = TAG_RE.exec(rPrBody); m !== null; m = TAG_RE.exec(rPrBody)) {
    if (m[1] === '/') {
      if (depth > 0) depth--;
      continue;
    }
    const selfClosed = m[0].endsWith('/>');
    if (depth === 0 && TEXT_FILL_NAMES.has(m[2]!)) {
      const el = findElement(rPrBody, m[2]!, m.index);
      TAG_RE.lastIndex = 0;
      return el === null ? null : { name: m[2]!, open: el.open, body: el.body };
    }
    if (!selfClosed) depth++;
  }
  TAG_RE.lastIndex = 0;
  return null;
}

/**
 * 文字のラン（`a:r`・`a:fld`・`a:br`）の色：`a:rPr` の `a:solidFill`・`a:gradFill`（代表色）→ `fallback`（図形の `fontRef` の色。`readFontColor`）。
 * `a:noFill` は透明（不透明度 0）。解けない色・`a:pattFill` などは `fallback`。`a:ln`（文字の輪郭）・`a:highlight` は読まない。
 * @param rPrBody `a:rPr` の中身（`a:rPr` が無い・自己閉じのときは null）
 */
export function readRunColor(rPrBody: string | null, fallback: Rgba, theme: Theme): Rgba {
  const el = rPrBody === null ? null : firstDirectFill(rPrBody);
  if (el === null) return { ...fallback };
  if (el.name === 'noFill') return { rgb: fallback.rgb, alpha: 0 };
  if (el.name === 'solidFill' || el.name === 'gradFill') {
    return fillElementColor(el, theme) ?? { ...fallback };
  }
  return { ...fallback };
}
