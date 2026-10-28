// 正: InterHub docs/plan/adr/0025-xlsx-own-reader.md
/**
 * XML を正規表現で走査する小さな道具と、セルの番地の計算。
 * DOMParser に依存せず Node でも動かすため、タグは名前空間の接頭辞を許して照合する
 * （`<c>` も `<x:c>` も読む。実験 05 §6）。
 */
import { MAX_COLS, MAX_ROWS } from './limits';

/** 要素名の前に付いてよい名前空間の接頭辞（正規表現の断片。例 `x:`） */
export const NS = String.raw`(?:[A-Za-z_][\w.-]*:)?`;

/** コード単位 1 つの文字参照を戻す。範囲外・不正な値は U+FFFD */
function codePointToString(n: number): string {
  if (!Number.isFinite(n) || n < 0 || n > 0x10ffff || (n >= 0xd800 && n <= 0xdfff)) return '�';
  return String.fromCodePoint(n);
}

const NAMED: Readonly<Record<string, string>> = Object.freeze({
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
});

/** XML の実体参照（`&amp;` `&#x41;` `&#65;`）を戻す。属性値と、`_xHHHH_` を使わない値に使う */
export function decodeEntities(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_m, e: string) => {
    if (e[0] === '#')
      return codePointToString(
        e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10),
      );
    return NAMED[e] ?? '';
  });
}

/**
 * 文字列セルの本文を戻す。実体参照のあと、Excel の `_xHHHH_`（制御文字などの退避）を戻す。
 * `_x005F_` は `_` を表す（本物の `_x0041_` という文字列を守るため）。
 */
export function decodeText(s: string): string {
  const e = decodeEntities(s);
  if (e.indexOf('_x') < 0) return e;
  return e.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));
}

/** `findElement` の結果 */
export interface ElementSpan {
  /** 開きタグ全体（属性を含む。`parseAttrs` に渡せる） */
  open: string;
  /** 中身。自己閉じ（`<cols/>`）・閉じタグの無いときは null */
  body: string | null;
  /** 要素の終わりの次の位置（次の検索の開始位置）。閉じタグの無いときは xml の長さ（検索を打ち切る） */
  end: number;
}

/**
 * `from` 以降で最初の `<name ...>…</name>`（名前空間の接頭辞可）を、線形時間で切り出す。
 * 開きタグを探し、閉じタグは `indexOf` で探す（`[\s\S]*?` の正規表現は、閉じが無いと入れ子の走査で二乗になる）。
 */
export function findElement(xml: string, name: string, from = 0): ElementSpan | null {
  const re = new RegExp(String.raw`<(${NS})${name}\b[^<>]*>`, 'g');
  re.lastIndex = from;
  const m = re.exec(xml);
  if (!m) return null;
  const openEnd = m.index + m[0].length;
  if (m[0].endsWith('/>')) return { open: m[0], body: null, end: openEnd };
  const closeTag = `</${m[1] ?? ''}${name}>`;
  const close = xml.indexOf(closeTag, openEnd);
  if (close < 0) return { open: m[0], body: null, end: xml.length };
  return { open: m[0], body: xml.slice(openEnd, close), end: close + closeTag.length };
}

/** forEachElement の開きタグの正規表現（名前ごとに 1 つ。lastIndex は呼び出しのたびに設定し直すので、入れ子の呼び出しで共有しても壊れない） */
const openTagCache = new Map<string, RegExp>();

/**
 * `<name …>…</name>`（名前空間の接頭辞可）を出現順に callback へ渡す。同じ名前の入れ子は扱わない（`c`・`si`・`t` など入れ子にならない要素用）。
 * 開きタグは `[^<>]*>`、閉じは `indexOf` で探すので、`>` や閉じの無い要素を大量に並べても時間は線形。
 * 自己閉じ（`<c/>`）は body を null で渡す。閉じタグの無い要素に当たったら、以後に閉じが無いので打ち切る。
 * 正規表現は名前ごとに 1 度だけ作る（セル 1 つごとに作らない）。callback の中で同じ `xml` を再び走査してよい。
 */
export function forEachElement(
  xml: string,
  name: string,
  cb: (open: string, body: string | null) => void,
): void {
  let re = openTagCache.get(name);
  if (re === undefined) {
    re = new RegExp(String.raw`<(${NS})${name}(?=[\s>/])[^<>]*>`, 'g');
    openTagCache.set(name, re);
  }
  for (let pos = 0; ;) {
    re.lastIndex = pos;
    const m = re.exec(xml);
    if (m === null) return;
    pos = m.index + m[0].length;
    if (m[0].endsWith('/>')) {
      cb(m[0], null);
      continue;
    }
    const closeTag = `</${m[1] ?? ''}${name}>`;
    const close = xml.indexOf(closeTag, pos);
    if (close < 0) return;
    cb(m[0], xml.slice(pos, close));
    pos = close + closeTag.length;
  }
}

/** findOpenTag の正規表現（名前ごとに 1 つ。lastIndex は呼び出しのたびに設定し直す） */
const findOpenTagCache = new Map<string, RegExp>();

/**
 * `from` 以降で最初の `<name …>` の開きタグだけを返す（閉じは探さない。名前空間の接頭辞可。無ければ null）。
 * 開きタグは `[^<>]*>` で打ち切るので、`>` の無いタグを大量に並べても時間は線形。正規表現は名前ごとに 1 度だけ作る。
 */
export function findOpenTag(xml: string, name: string, from = 0): string | null {
  let re = findOpenTagCache.get(name);
  if (re === undefined) {
    re = new RegExp(String.raw`<${NS}${name}(?=[\s>/])[^<>]*>`, 'g');
    findOpenTagCache.set(name, re);
  }
  re.lastIndex = from;
  return re.exec(xml)?.[0] ?? null;
}

/** 属性の一覧。キーに `__proto__` などが来てもプロトタイプを汚さないよう、プロトタイプの無いオブジェクトで返す */
export type Attrs = Record<string, string | undefined>;

/**
 * タグ（`<c r="A1" s="2">` の属性の部分）から属性を取り出す。値は実体参照を戻す。
 * 名前の先頭は後読みで「前が名前の文字でない位置」に限る（長い名前の途中から照合し直して二乗時間になるのを防ぐ。線形時間）。
 */
export function parseAttrs(tag: string): Attrs {
  const o: Attrs = Object.create(null);
  for (const m of tag.matchAll(/(?<![\w:.-])([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) {
    o[m[1]!] = decodeEntities(m[2] ?? m[3] ?? '');
  }
  return o;
}

/** 名前空間の接頭辞を除いた属性の値（`r:id` と `rel:id` を同じに扱う）。`id` を探すとき `sheetId` は拾わない */
export function attrByLocalName(attrs: Attrs, local: string): string | undefined {
  for (const k in attrs) {
    if (k === local || k.endsWith(':' + local)) return attrs[k];
  }
  return undefined;
}

/** `A`→1、`Z`→26、`AA`→27、`XFD`→16384 */
function colNumber(letters: string): number {
  let c = 0;
  for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
  return c;
}

/** 列番号（1 始まり）を列の記号にする。1→A、27→AA。1 未満・整数でなければ空文字 */
export function colName(c: number): string {
  if (!Number.isInteger(c) || c < 1) return '';
  let s = '';
  for (let n = c; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** 番地 `AB12` の位置（1 始まり）。範囲外・不正なら null */
export function parseCellRef(ref: string): { r: number; c: number } | null {
  const m = /^([A-Z]{1,3})(\d{1,7})$/.exec(ref);
  if (!m) return null;
  const c = colNumber(m[1]!);
  const r = parseInt(m[2]!, 10);
  if (c < 1 || c > MAX_COLS || r < 1 || r > MAX_ROWS) return null;
  return { r, c };
}

/** 列だけの番地（`AB`）。セルの r 属性（`AB12`）から列だけ欲しいとき用。範囲外・不正なら 0 */
export function colOfRef(ref: string): number {
  const m = /^([A-Z]{1,3})\d*$/.exec(ref);
  if (!m) return 0;
  const c = colNumber(m[1]!);
  return c >= 1 && c <= MAX_COLS ? c : 0;
}

/** 開きタグの名前の直後の文字として許す文字（`<c` の後が `o`（`<cols`）などでないことの確認用） */
export function isTagDelimiter(ch: string | undefined): boolean {
  return ch === ' ' || ch === '>' || ch === '/' || ch === '\t' || ch === '\n' || ch === '\r';
}

/** Strict 形式（ISO 29500 Strict）の名前空間の目印。1 次では読まない（unsupported） */
export const STRICT_NS = 'purl.oclc.org/ooxml';

/** ルート要素の宣言（先頭の 2,000 文字）に Strict の名前空間があるか */
export function looksStrict(xml: string): boolean {
  return xml.slice(0, 2000).includes(STRICT_NS);
}
