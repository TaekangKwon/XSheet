// 正: docs/adr/0008-align-wrap-spill.md §1
/**
 * cellXfs の `<xf>` の子 `<alignment>`（横・縦の配置と折り返し）の読み取り（M7-12・Moobius の依頼 R9）。
 * - `cell-format.ts`（InterHub の写し）は、xf ごとの書式に {@link readXfAlignments} の結果を足し、まとめるキーに {@link alignKey} を足すだけ。
 * - `applyAlignment` は見ない（塗り・文字の装飾と同じ。F5）。cellStyleXfs は読まない。
 * - 走査は線形：`<xf>`・`</xf>`・`<alignment>` の開きタグを 1 本の正規表現で順に拾う。閉じの無い xf は次の xf で終わったとみなす。
 * DOM・Node 固有の API を使わない。
 */
import type { CellFormat } from './types';
import { NS, parseAttrs } from './xml';

/** 書式のうち配置の項目 */
export type CellAlign = Pick<CellFormat, 'hAlign' | 'vAlign' | 'wrap'>;

/** horizontal の値 → 持つ値（`general`・不明は持たない） */
const H_ALIGN: ReadonlyMap<string, NonNullable<CellFormat['hAlign']>> = new Map([
  ['left', 'left'],
  ['center', 'center'],
  ['right', 'right'],
  ['distributed', 'left'],
  ['fill', 'left'],
  ['justify', 'left'],
  ['centerContinuous', 'center'],
] as const);

/** vertical の値 → 持つ値（`bottom`・不明は持たない） */
const V_ALIGN: ReadonlyMap<string, NonNullable<CellFormat['vAlign']>> = new Map([
  ['top', 'top'],
  ['center', 'center'],
  ['justify', 'top'],
  ['distributed', 'top'],
] as const);

/** `<alignment>` の開きタグ → 配置（何も無ければ undefined） */
function readAlignment(tag: string): CellAlign | undefined {
  const a = parseAttrs(tag);
  const out: CellAlign = {};
  const h = a.horizontal === undefined ? undefined : H_ALIGN.get(a.horizontal);
  if (h !== undefined) out.hAlign = h;
  const v = a.vertical === undefined ? undefined : V_ALIGN.get(a.vertical);
  if (v !== undefined) out.vAlign = v;
  if (a.wrapText === '1' || a.wrapText === 'true') out.wrap = true;
  return out.hAlign === undefined && out.vAlign === undefined && out.wrap === undefined
    ? undefined
    : out;
}

/**
 * cellXfs の中身 → xf の位置ごとの配置（配置の無い xf は undefined）。
 * 位置は `<xf` の開きタグの出現順（`cell-format.ts` の数え方と同じ）。
 * @param body `<cellXfs>` の中身
 */
export function readXfAlignments(body: string): (CellAlign | undefined)[] {
  const out: (CellAlign | undefined)[] = [];
  /** 今の xf の位置（開いていなければ -1） */
  let open = -1;
  const re = new RegExp(String.raw`<(/?)${NS}(xf|alignment)(?=[\s>/])[^<>]*>`, 'g');
  for (const m of body.matchAll(re)) {
    const closing = m[1] === '/';
    if (m[2] === 'xf') {
      if (closing) {
        open = -1;
        continue;
      }
      out.push(undefined);
      open = m[0].endsWith('/>') ? -1 : out.length - 1;
      continue;
    }
    if (!closing && open >= 0 && out[open] === undefined) out[open] = readAlignment(m[0]);
  }
  return out;
}

/** 書式をまとめるキーの配置の部分（数値だけ：横 0〜3・縦 0〜2・折り返し 0〜1） */
export function alignKey(f: CellAlign): string {
  const h = f.hAlign === 'left' ? 1 : f.hAlign === 'center' ? 2 : f.hAlign === 'right' ? 3 : 0;
  const v = f.vAlign === 'top' ? 1 : f.vAlign === 'center' ? 2 : 0;
  return `${h}${v}${f.wrap ? 1 : 0}`;
}
