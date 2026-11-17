// 正: docs/adr/0009-rich-text-runs.md §1
/**
 * リッチテキストの区切り（`<si>`・`<is>` の `<r>`）と、その書式（`<rPr>`）の読み取り（M7-16）。
 * - `<rPr>` の中身は `<font>` と同じ読み方（`cell-format.ts` が渡す読み手）。
 * - 区切りの文字を連結したものが平文と違えば、区切りを持たない（平文で描く）。
 * - 走査は線形（`forEachElement`・`findElement`。閉じは indexOf）。DOM 非依存。
 */
import type { RunFont, TextRun } from './types';
import { decodeText, findElement, forEachElement } from './xml';

/** `<rPr>` の中身 → 区切りの書式 */
export type RunFontReader = (rPrBody: string | null) => RunFont;

/** 区切りの書式を持ちうる断片か（`<rPr` が無ければ読まない） */
export function mayHaveRuns(inner: string): boolean {
  return inner.indexOf('rPr') >= 0;
}

/**
 * `<si>`・`<is>` の中身から区切りを読む。
 * @param inner 開きタグと閉じタグの間の文字列
 * @param text セルの文字（平文。`MAX_CELL_TEXT` で切った後）。区切りも同じ長さで切る
 * @param readFont `<rPr>` の読み手
 * @returns `<rPr>` のある区切りが無い・連結が平文と合わないときは undefined
 */
export function readRuns(
  inner: string,
  text: string,
  readFont: RunFontReader,
): TextRun[] | undefined {
  if (!mayHaveRuns(inner)) return undefined;
  const runs: TextRun[] = [];
  let styled = false;
  let joined = '';
  forEachElement(inner, 'r', (_open, body) => {
    if (body === null) return;
    let t = '';
    forEachElement(body, 't', (_o, tb) => {
      if (tb) t += decodeText(tb);
    });
    const rPr = findElement(body, 'rPr');
    joined += t;
    if (t === '') return;
    const run: TextRun = { text: t };
    if (rPr !== null) {
      run.font = readFont(rPr.body);
      styled = true;
    }
    runs.push(run);
  });
  if (!styled || !joined.startsWith(text)) return undefined; // 長さの上限で切った平文は、連結の先頭と同じ
  return joined.length === text.length ? runs : cutRuns(runs, text.length);
}

/** 区切りを先頭から length 文字で切る（UTF-16 の長さ） */
function cutRuns(runs: readonly TextRun[], length: number): TextRun[] {
  const out: TextRun[] = [];
  let pos = 0;
  for (const run of runs) {
    if (pos >= length) break;
    const rest = length - pos;
    out.push(run.text.length <= rest ? run : { ...run, text: run.text.slice(0, rest) });
    pos += run.text.length;
  }
  return out;
}
