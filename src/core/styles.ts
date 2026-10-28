// 正: InterHub docs/plan/adr/0025-xlsx-own-reader.md
/**
 * styles.xml の読み取り：セルの書式（cellXfs）の番号 → 数値の表示形式の番号、独自の表示形式（numFmts）。
 * 書式のうち塗り・文字の装飾は読まない（塗り・太字・色などは cell-format.ts、罫線は borders.ts、色の解釈は cell-color.ts）。ただし寸法の換算に要る、先頭の font（名前・サイズ）から
 * 最大桁幅 MDW と、フォントの既定の行高を出す（実測表は font-metrics.ts）。
 */
import { MAX_FONT_SIZE, fontDefaultRowPx, lookupMdw } from './font-metrics';
import { DEFAULT_MDW } from './types';
import { NS, findElement, parseAttrs } from './xml';

/** 読んだ styles.xml */
export interface Styles {
  /** セルの `s` 属性（cellXfs の位置）→ numFmtId */
  xfNumFmtIds: number[];
  /** 独自の表示形式（numFmtId → 書式コード）。Map なので `__proto__` などを入れても汚染されない */
  customFormats: Map<number, string>;
  /** 最大桁幅 MDW（px）。先頭の font（`<fonts>` の最初の `<font>`）から決める。読めなければ 7 */
  mdw: number;
  /** 先頭の font の既定の行高（px・96dpi）。実測表に無ければ null */
  fontRowPx: number | null;
}

/**
 * フォントの名前・サイズ（pt）→ 最大桁幅 MDW（px）。
 * 実測表（font-metrics.ts）から、サイズの間は線形補間・範囲外は比例で求め、表に無いフォントは Calibri の行を使う。
 * Calibri 11＝7px、游ゴシック 11＝8px、맑은 고딕 11＝8px。サイズが不正なら 7px。
 */
export function maxDigitWidth(name: string, size: number): number {
  return lookupMdw(name, size) ?? DEFAULT_MDW;
}

/** styles.xml を読む（無ければ空）。書式コードは実体参照を戻した文字列 */
export function parseStyles(xml: string | null): Styles {
  const styles: Styles = {
    xfNumFmtIds: [],
    customFormats: new Map(),
    mdw: DEFAULT_MDW,
    fontRowPx: null,
  };
  if (xml === null) return styles;

  for (const m of xml.matchAll(new RegExp(String.raw`<${NS}numFmt\b[^<>]*>`, 'g'))) {
    const a = parseAttrs(m[0]);
    const id = Number(a.numFmtId);
    if (Number.isInteger(id) && a.formatCode !== undefined)
      styles.customFormats.set(id, a.formatCode);
  }

  // 先頭の font（<fonts> の最初の <font>）の名前・サイズ → MDW・既定の行高（線形に切り出す）
  const fonts = findElement(xml, 'fonts');
  const font = fonts?.body ? findElement(fonts.body, 'font') : null;
  if (font?.body) {
    const sz = new RegExp(String.raw`<${NS}sz\b[^<>]*>`).exec(font.body);
    const nm = new RegExp(String.raw`<${NS}name\b[^<>]*>`).exec(font.body);
    if (sz) {
      const raw = Number(parseAttrs(sz[0]).val);
      const size = Number.isFinite(raw) ? Math.min(raw, MAX_FONT_SIZE) : raw;
      const name = nm ? (parseAttrs(nm[0]).val ?? '') : '';
      styles.mdw = maxDigitWidth(name, size);
      styles.fontRowPx = fontDefaultRowPx(name, size);
    }
  }

  // cellXfs の中の <xf> だけを数える（cellStyleXfs の <xf> と混ぜない）
  const block = findElement(xml, 'cellXfs');
  if (block?.body) {
    for (const m of block.body.matchAll(new RegExp(String.raw`<${NS}xf\b[^<>]*>`, 'g'))) {
      const id = Number(parseAttrs(m[0]).numFmtId);
      styles.xfNumFmtIds.push(Number.isInteger(id) ? id : 0);
    }
  }
  return styles;
}
