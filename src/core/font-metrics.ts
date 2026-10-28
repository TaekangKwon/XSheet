// 正: InterHub docs/plan/adr/0026-xlsx-image-overlay.md（§決定 2・3）
/**
 * フォントごとの寸法の実測表：最大桁幅 MDW（px）と、既定の行高（px）。
 * 出所: 2026-10-27・Windows GDI（TextRenderer）96dpi で数字 0〜9 の最大幅を実測。
 * Calibri 11＝7・游ゴシック 11＝8・맑은 고딕 11＝8・MS Pゴシック 11＝8 は、既知の Excel の値と一致する。
 * 既定の行高は Excel 16 の COM で確認した値（96dpi 換算）。
 */

/** サイズ（pt）と最大桁幅（px）の組（サイズの昇順） */
type MdwRow = readonly (readonly [size: number, mdw: number])[];

/** 表に無いフォントに使う行の名前 */
const FALLBACK = 'calibri';

const CALIBRI: MdwRow = [
  [6, 4],
  [8, 6],
  [9, 6],
  [10, 7],
  [11, 7],
  [12, 8],
  [14, 10],
  [16, 11],
  [18, 12],
  [20, 14],
  [24, 16],
];
const ARIAL: MdwRow = [
  [6, 4],
  [8, 6],
  [9, 7],
  [10, 8],
  [11, 8],
  [12, 9],
  [14, 11],
  [16, 12],
  [18, 13],
  [20, 15],
  [24, 18],
];
const TIMES: MdwRow = [
  [6, 4],
  [8, 6],
  [9, 6],
  [10, 7],
  [11, 7],
  [12, 8],
  [14, 9],
  [16, 11],
  [18, 12],
  [20, 14],
  [24, 16],
];
const MS_GOTHIC: MdwRow = [
  [6, 4],
  [8, 6],
  [9, 6],
  [10, 7],
  [11, 8],
  [12, 8],
  [14, 10],
  [16, 11],
  [18, 12],
  [20, 14],
  [24, 16],
];
const MEIRYO: MdwRow = [
  [6, 5],
  [8, 7],
  [9, 7],
  [10, 9],
  [11, 9],
  [12, 10],
  [14, 12],
  [16, 14],
  [18, 15],
  [20, 17],
  [24, 20],
];
const MALGUN: MdwRow = [
  [6, 4],
  [8, 6],
  [9, 7],
  [10, 8],
  [11, 8],
  [12, 9],
  [14, 10],
  [16, 12],
  [18, 13],
  [20, 15],
  [24, 18],
];
const GULIM: MdwRow = [
  [6, 5],
  [8, 6],
  [9, 6],
  [10, 8],
  [11, 8],
  [12, 9],
  [14, 10],
  [16, 11],
  [18, 13],
  [20, 16],
  [24, 18],
];
const DOTUM: MdwRow = [
  [6, 5],
  [8, 6],
  [9, 6],
  [10, 9],
  [11, 9],
  [12, 8],
  [14, 10],
  [16, 12],
  [18, 13],
  [20, 16],
  [24, 19],
];
const SEGOE: MdwRow = [
  [6, 4],
  [8, 6],
  [9, 6],
  [10, 8],
  [11, 8],
  [12, 9],
  [14, 10],
  [16, 12],
  [18, 13],
  [20, 15],
  [24, 17],
];

/** 正規化したフォント名 → 実測表（Yu Gothic は Arial と同じ値） */
const MDW_TABLE: Readonly<Record<string, MdwRow>> = {
  calibri: CALIBRI,
  arial: ARIAL,
  timesnewroman: TIMES,
  yugothic: ARIAL,
  mspgothic: MS_GOTHIC,
  msgothic: MS_GOTHIC,
  meiryo: MEIRYO,
  meiryoui: MEIRYO,
  malgungothic: MALGUN,
  gulim: GULIM,
  dotum: DOTUM,
  segoeui: SEGOE,
};

/** 現地語名・全角表記 → 表の名前（正規化したあとの名前で引く） */
const ALIASES: Readonly<Record<string, string>> = {
  游ゴシック: 'yugothic',
  mspゴシック: 'mspgothic',
  msゴシック: 'msgothic',
  メイリオ: 'meiryo',
  メイリオui: 'meiryoui',
  맑은고딕: 'malgungothic',
  굴림: 'gulim',
  돋움: 'dotum',
};

/** 既定の行高（px・96dpi）。キーは「正規化した名前:サイズ」 */
const DEFAULT_ROW_PX: Readonly<Record<string, number>> = {
  'calibri:11': 20,
  'yugothic:11': 25,
  'malgungothic:11': 22,
  'mspgothic:11': 18,
  'arial:10': 17,
};

/**
 * フォント名を表の名前にそろえる。NFKC 正規化（全角英字を半角に）・小文字・空白除去のあと、現地語名の別名を引く。
 * 表に無い名前は正規化しただけの名前を返す。
 */
export function canonicalFontName(name: string): string {
  const n = name.normalize('NFKC').toLowerCase().replace(/\s+/g, '');
  return ALIASES[n] ?? n;
}

/** サイズの上限（pt）。Excel のフォントサイズの上限 */
export const MAX_FONT_SIZE = 409;

/**
 * フォントの名前・サイズ（pt）→ 最大桁幅 MDW（px・96dpi）。
 * 実測表のサイズの間は、近い 2 点から線形補間して四捨五入する。表の範囲外は、端の点に比例させる。
 * 表に無いフォントは Calibri の行を使う。サイズが不正（数でない・0 以下）なら null。
 */
export function lookupMdw(name: string, size: number): number | null {
  if (!Number.isFinite(size) || size <= 0) return null;
  const s = Math.min(size, MAX_FONT_SIZE);
  const row = MDW_TABLE[canonicalFontName(name)] ?? MDW_TABLE[FALLBACK]!;
  const first = row[0]!;
  const last = row[row.length - 1]!;
  if (s <= first[0]) return Math.max(1, Math.round((s * first[1]) / first[0]));
  if (s >= last[0]) return Math.max(1, Math.round((s * last[1]) / last[0]));
  for (let i = 1; i < row.length; i++) {
    const [s1, m1] = row[i]!;
    if (s > s1) continue;
    const [s0, m0] = row[i - 1]!;
    return Math.max(1, Math.round(m0 + ((m1 - m0) * (s - s0)) / (s1 - s0)));
  }
  return last[1];
}

/**
 * フォントの既定の行高（px・96dpi）。実測した（名前・サイズ）の組だけを持ち、無ければ null。
 * Calibri 11＝20、游ゴシック 11＝25、맑은 고딕 11＝22、MS Pゴシック 11＝18、Arial 10＝17。
 */
export function fontDefaultRowPx(name: string, size: number): number | null {
  return DEFAULT_ROW_PX[`${canonicalFontName(name)}:${size}`] ?? null;
}
