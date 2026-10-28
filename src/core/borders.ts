// 正: InterHub docs/plan/adr/0029-xlsx-cell-borders.md（§決定 2・6・9）
/**
 * styles.xml の罫線の読み取り：`<borders>` の各 `<border>`（各辺の線種と色）と、cellXfs の `borderId`。
 * ブックで 1 回だけ読み、セルの `s` 属性（cellXfs の位置）→ 罫線の番号 → 各辺、と引けるようにする（`BorderTable`）。
 * - どの辺も線を持たない `<border>`（Excel の `borderId="0"` など）は「罫線なし」にまとめ、セルに番号を付けない。
 * - 未知・`none`・無い `style` は線なし。`start`・`end` は `left`・`right`。`applyBorder` は見ない（E7）。
 * - 色は SpreadsheetML の `<color>`（cell-color.ts）をテーマで解いてから持つ。解けない色は null（自動）。
 * - 走査は線形（開きタグは `[^<>]*>`、閉じは `indexOf`）。解けない値は例外にせず罫線なしに倒す。
 */
import { type Rgba } from './color';
import { parseIndexedColors, resolveCellColor } from './cell-color';
import type { XlsxLimits } from './limits';
import type { Theme } from './theme';
import type { BorderEdge, BorderStyle, CellBorder } from './types';
import { NS, findElement, findOpenTag, forEachElement, parseAttrs } from './xml';

/** 罫線の表 */
export interface BorderTable {
  /** 罫線の番号 → 罫線（線のある罫線だけ。`SheetRow.borders` の値・`XlsxSheet.borders` の位置） */
  readonly borders: readonly CellBorder[];
  /** cellXfs の位置（セルの `s`）→ 罫線の番号。罫線なし・不正・上限の外は -1 */
  readonly xfBorder: Int32Array;
}

/** 罫線の無い表（styles.xml が無い・壊れているとき。共有する・変更しない） */
export const EMPTY_BORDER_TABLE: BorderTable = Object.freeze({
  borders: Object.freeze([] as CellBorder[]),
  xfBorder: new Int32Array(0),
});

/** 線種の名前（`none`・未知は含めない）。`Set` なので `__proto__` などの名前でも汚染されない */
const BORDER_STYLES: ReadonlySet<string> = new Set<BorderStyle>([
  'thin',
  'hair',
  'dotted',
  'dashed',
  'dashDot',
  'dashDotDot',
  'medium',
  'mediumDashed',
  'mediumDashDot',
  'mediumDashDotDot',
  'slantDashDot',
  'thick',
  'double',
]);

/** 真偽の属性（`1`・`true` が真） */
const boolAttr = (v: string | undefined): boolean => v === '1' || v === 'true';

/** 整数の属性（空・数でない・小数は null） */
function intAttr(v: string | undefined): number | null {
  if (v === undefined || v.trim() === '') return null;
  const n = Number(v);
  return Number.isInteger(n) ? n : null;
}

/**
 * `<border>` の中から辺 1 つ（要素名 name）を読む。要素が無い・線種が無い・未知・`none` なら undefined。
 * 色は辺の中の最初の `<color>`。無ければ null（自動）。
 */
function readEdge(
  body: string,
  name: string,
  theme: Theme,
  palette: readonly number[] | null,
): BorderEdge | undefined {
  let edge: BorderEdge | undefined;
  let seen = false;
  forEachElement(body, name, (open, inner) => {
    if (seen) return;
    seen = true;
    const style = parseAttrs(open).style;
    if (style === undefined || !BORDER_STYLES.has(style)) return;
    const colorTag = inner === null ? null : findOpenTag(inner, 'color');
    const color: Rgba | null =
      colorTag === null ? null : resolveCellColor(parseAttrs(colorTag), theme, palette);
    edge = { style: style as BorderStyle, color };
  });
  return edge;
}

/** `<border>` 1 件（開きタグと中身）→ 罫線。どの辺にも線が無ければ null */
function readBorder(
  open: string,
  body: string | null,
  theme: Theme,
  palette: readonly number[] | null,
): CellBorder | null {
  if (body === null) return null;
  const attrs = parseAttrs(open);
  const border: CellBorder = {
    diagonalUp: boolAttr(attrs.diagonalUp),
    diagonalDown: boolAttr(attrs.diagonalDown),
  };
  const left = readEdge(body, 'left', theme, palette) ?? readEdge(body, 'start', theme, palette);
  const right = readEdge(body, 'right', theme, palette) ?? readEdge(body, 'end', theme, palette);
  const top = readEdge(body, 'top', theme, palette);
  const bottom = readEdge(body, 'bottom', theme, palette);
  const diagonal = readEdge(body, 'diagonal', theme, palette);
  if (left) border.left = left;
  if (right) border.right = right;
  if (top) border.top = top;
  if (bottom) border.bottom = bottom;
  if (diagonal) border.diagonal = diagonal;
  return left || right || top || bottom || diagonal ? border : null;
}

/**
 * styles.xml から罫線の表を作る（ブックで 1 回）。
 * @param xml styles.xml（無ければ null＝罫線なし）
 * @param theme ブックのテーマ（色の `theme` を解く）
 * @param limits `maxBorders` を超えた `<border>` は読まない（それを参照するセルは罫線なし）
 */
export function parseBorders(xml: string | null, theme: Theme, limits: XlsxLimits): BorderTable {
  if (xml === null) return EMPTY_BORDER_TABLE;
  const block = findElement(xml, 'borders');
  const cellXfs = findElement(xml, 'cellXfs');
  if (block === null || block.body === null || cellXfs === null || cellXfs.body === null)
    return EMPTY_BORDER_TABLE;

  const palette = parseIndexedColors(xml);
  const borders: CellBorder[] = [];
  // <border> の位置（borderId）→ borders の位置（線の無い border は -1）
  const idToIndex: number[] = [];
  forEachElement(block.body, 'border', (open, body) => {
    if (idToIndex.length >= limits.maxBorders) return;
    const b = readBorder(open, body, theme, palette);
    if (b === null) {
      idToIndex.push(-1);
    } else {
      idToIndex.push(borders.length);
      borders.push(b);
    }
  });

  const xfBorder: number[] = [];
  for (const m of cellXfs.body.matchAll(new RegExp(String.raw`<${NS}xf\b[^<>]*>`, 'g'))) {
    const id = intAttr(parseAttrs(m[0]).borderId);
    xfBorder.push(id !== null && id >= 0 ? (idToIndex[id] ?? -1) : -1);
  }
  return { borders, xfBorder: Int32Array.from(xfBorder) };
}
