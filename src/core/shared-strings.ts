// 正: InterHub docs/plan/adr/0025-xlsx-own-reader.md
/**
 * sharedStrings.xml の読み取り。
 * - ふりがな（`rPh`）は除く。リッチテキスト（`r` の連なり）は連結して平文にする。
 * - 文字の `&amp;` などの実体参照と、Excel の `_xHHHH_` を戻す。
 * - 1 件ずつの展開は引かれたときに行う（索引だけ先に作る）。
 */
import { NS, decodeText, forEachElement } from './xml';

/** 共有文字列の表 */
export interface SharedStrings {
  readonly count: number;
  /** i 番目（0 始まり）の文字列。範囲外は空文字 */
  get(i: number): string;
}

// ふりがな <rPh>…</rPh> の開きタグ。閉じは indexOf で探す（`[\s\S]*?` だと、閉じの無い <rPh を並べた細工で二乗時間になる）
const RPH_OPEN = new RegExp(String.raw`<(${NS})rPh(?=[\s>/])[^<>]*>`, 'g');

/** ふりがな（`<rPh>…</rPh>`）を取り除く。閉じの無い `<rPh>` に当たったら、以後に閉じは無いので残りはそのまま残す */
function stripPhonetic(s: string): string {
  RPH_OPEN.lastIndex = 0;
  let out = '';
  let pos = 0;
  for (;;) {
    const m = RPH_OPEN.exec(s);
    if (m === null) break;
    let end = RPH_OPEN.lastIndex;
    if (!m[0].endsWith('/>')) {
      const closeTag = `</${m[1] ?? ''}rPh>`;
      const close = s.indexOf(closeTag, end);
      if (close < 0) break;
      end = close + closeTag.length;
    }
    out += s.slice(pos, m.index);
    pos = end;
    RPH_OPEN.lastIndex = end;
  }
  return pos === 0 ? s : out + s.slice(pos);
}

/**
 * `<si>`・`<is>` の中身（XML の断片）を平文にする。ふりがなを除き、`<t>` を順に連結する。
 * @param inner 開きタグと閉じタグの間の文字列
 */
export function richTextToPlain(inner: string): string {
  const body = inner.indexOf('rPh') >= 0 ? stripPhonetic(inner) : inner;
  let out = '';
  forEachElement(body, 't', (_open, text) => {
    if (text) out += decodeText(text);
  });
  return out;
}

/** sharedStrings.xml を読む。`<si>` の位置の索引だけを作り、文字への展開は get で行う */
export function parseSharedStrings(xml: string | null): SharedStrings {
  const inners: string[] = [];
  if (xml !== null) {
    forEachElement(xml, 'si', (_open, body) => inners.push(body ?? ''));
  }
  return {
    count: inners.length,
    get(i) {
      const inner = inners[i];
      return inner === undefined ? '' : richTextToPlain(inner);
    },
  };
}
