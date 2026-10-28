// 正: InterHub docs/plan/adr/0028-xlsx-preset-shape-geometry.md（§決定 6.2・§決定 7）
/**
 * ブックのテーマの 12 色（`xl/theme/theme1.xml` の `a:clrScheme`）の読み取り。
 * テーマの書式（`a:fmtScheme`）・フォント（`a:fontScheme`）は読まない。
 * テーマが無い・読めない・上限を超える・色が欠けているときは、Office 2013〜2022 の既定の 12 色で補う
 * （`XlsxError` にもバナーにもしない）。テーマはブックに 1 つで、styles.xml と同じくブックで 1 回だけ読む（index.ts）。
 */
import type { XlsxLimits } from './limits';
import type { XlsxPackage } from './open';
import { type Attrs, NS, findElement, parseAttrs } from './xml';
import { XlsxError } from './types';

/** テーマの色の位置（`a:clrScheme` の子の要素名）。この順で 12 色 */
export const THEME_SLOTS = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;

export type ThemeSlot = (typeof THEME_SLOTS)[number];

/** テーマの 12 色（位置 → RGB の整数）。`Map` なので位置の名前が `__proto__` などでも汚染されない。常に 12 色そろう */
export type Theme = ReadonlyMap<ThemeSlot, number>;

/** 既定の 12 色（Office 2013〜2022。Excel 16 が書く theme1.xml と同じ値。`THEME_SLOTS` の順） */
const DEFAULT_VALUES: readonly number[] = [
  0x000000, 0xffffff, 0x44546a, 0xe7e6e6, 0x4472c4, 0xed7d31, 0xa5a5a5, 0xffc000, 0x5b9bd5,
  0x70ad47, 0x0563c1, 0x954f72,
];

/** 既定のテーマ（共有する。変更しない） */
export const DEFAULT_THEME: Theme = new Map(THEME_SLOTS.map((s, i) => [s, DEFAULT_VALUES[i]!]));

/** 16 進 6 桁（`[0-9A-Fa-f]{6}`）だけを整数にする。それ以外・未指定は null */
export function parseHex6(s: string | undefined): number | null {
  return s !== undefined && /^[0-9A-Fa-f]{6}$/.test(s) ? parseInt(s, 16) : null;
}

/** `sysClr` の色：`lastClr`（16 進 6 桁）。無い・不正なら `windowText`＝黒・`window`＝白、その他は null */
export function sysClrValue(attrs: Attrs): number | null {
  const last = parseHex6(attrs.lastClr);
  if (last !== null) return last;
  if (attrs.val === 'windowText') return 0x000000;
  if (attrs.val === 'window') return 0xffffff;
  return null;
}

/** 位置の中の色の要素（`srgbClr`・`sysClr`）の開きタグ。`[^<>]*>` で打ち切るので線形時間 */
const SLOT_COLOR_RE = new RegExp(String.raw`<${NS}(srgbClr|sysClr)(?=[\s>/])[^<>]*>`);

/** 位置の中身（`<a:srgbClr val="…"/>` など）から色を取る。解けなければ null */
function slotColor(body: string): number | null {
  const m = SLOT_COLOR_RE.exec(body);
  if (m === null) return null;
  const a = parseAttrs(m[0]);
  return m[1] === 'srgbClr' ? parseHex6(a.val) : sysClrValue(a);
}

/**
 * `theme1.xml` の `a:clrScheme` から 12 色を読む。欠けた色・不正な色はその色だけ既定で補う。
 * xml が null・`clrScheme` が無い・閉じが無いときは既定の 12 色（新しい `Map`）。
 */
export function parseTheme(xml: string | null): Theme {
  const out = new Map<ThemeSlot, number>(DEFAULT_THEME);
  if (xml === null) return out;
  const scheme = findElement(xml, 'clrScheme')?.body;
  if (scheme === null || scheme === undefined) return out;
  for (const slot of THEME_SLOTS) {
    const body = findElement(scheme, slot)?.body;
    if (body === null || body === undefined) continue;
    const rgb = slotColor(body);
    if (rgb !== null) out.set(slot, rgb);
  }
  return out;
}

/**
 * テーマを読む。`path` が無い・部品が zip に無い・上限（小さな部品）を超える・読めないときは既定の 12 色。
 * 画面を閉じた後の読み取り（`broken`・`closed`）だけは呼び出し側へ投げる（index.ts が閉じたかを確かめる）。
 */
export async function loadTheme(
  pkg: XlsxPackage,
  path: string | null,
  limits: XlsxLimits,
): Promise<Theme> {
  if (path === null || !pkg.has(path)) return parseTheme(null);
  try {
    return parseTheme(await pkg.readText(path, limits.maxSmallPartBytes));
  } catch (e) {
    if (e instanceof XlsxError && e.kind === 'broken' && /closed/.test(e.message)) throw e;
    return parseTheme(null);
  }
}
