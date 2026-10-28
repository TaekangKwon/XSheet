// 正: InterHub docs/plan/adr/0027-xlsx-shape-text-and-connectors.md（§決定 1・6）
/**
 * 図形（`xdr:sp`）の文字（`xdr:txBody`）の読み取り：段落（`a:p`）ごとに `a:r`・`a:br`・`a:fld` を出現順に並べ、書式（大きさ・太字・斜体・下線・取り消し線・揃え・箇条書き）と
 * 本文の枠（縦位置・余白・折り返し・縮小・縦書き）を読む。lib は pt・EMU・列挙値のまま持ち、px への換算は表示側。
 * 強調表示・フォント名・上付き／下付き・行間・段落の間隔・字下げ・`a:lstStyle`・ハイパーリンクは読まない。
 * 文字の色は `a:rPr` の塗りと図形の `fontRef` から解いて持たせる（ISSUE-95。描くのは ISSUE-96）。
 * DOM に依存せず、開きタグは `[^<>]*>`、閉じは `indexOf` で探して線形時間に走査する（`a:p`・`a:r`・`a:t` は入れ子にならない）。
 * 非信頼の入力として、文字数・段落数・ランの数に上限を掛ける（超えた分は捨てて `truncated` で知らせる）。文字はログに出さない（console を使わない）。
 */
import type {
  DrawingText,
  DrawingTextParagraph,
  DrawingTextRun,
  TextAlign,
  TextAnchor,
  TextVert,
} from './types';
import { NS, decodeEntities, findElement, findOpenTag, forEachElement, parseAttrs } from './xml';
import type { Rgba } from './color';
import { readFontColor, readRunColor } from './drawing-color';
import { boolAttr, intAttr } from './drawing-shape';
import { DEFAULT_THEME, type Theme } from './theme';

/** 文字の大きさの既定（pt。`sz` が無いとき） */
const DEFAULT_SIZE_PT = 11;
/** 文字の大きさの範囲（pt） */
const MIN_SIZE_PT = 1;
const MAX_SIZE_PT = 400;
/** 内側の余白の既定（EMU。左・上・右・下） */
const DEFAULT_INSETS = { l: 91440, t: 45720, r: 91440, b: 45720 } as const;
/** 内側の余白の上限（EMU。2^31-1） */
const MAX_INSET = 2_147_483_647;

const ALIGNS: ReadonlySet<string> = new Set<TextAlign>(['l', 'ctr', 'r', 'just', 'dist']);
const VERTS: ReadonlySet<string> = new Set<TextVert>([
  'horz',
  'vert',
  'vert270',
  'wordArtVert',
  'eaVert',
  'mongolianVert',
  'wordArtVertRtl',
]);

/** 段落の直下の要素（ここに当たったら中を飛ばして次へ進む。`a:fld` は中に `a:pPr`・`a:rPr` を持つ） */
const PARA_TOKEN_RE = new RegExp(
  String.raw`<(${NS})(pPr|r|br|fld|endParaRPr)(?=[\s>/])[^<>]*>`,
  'g',
);

/** `a:t` の開きタグ */
const T_RE = new RegExp(String.raw`<(${NS})t(?=[\s>/])[^<>]*>`, 'g');

/** タブ・改行以外の制御文字（C0・DEL・C1） */
// eslint-disable-next-line no-control-regex -- 制御文字を除くことが目的
const CONTROL_RE = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F]/g;

/** 読み取りの上限（`limits.ts` の値。テストでは小さい値を渡す） */
export interface TextCaps {
  maxChars: number;
  maxParagraphs: number;
  maxRuns: number;
  /** 1 シートの合計の残り（文字数・ラン・段落を 1 つずつ数えた数。この図形の分がここを超えたら以後を捨てる） */
  maxUnits: number;
}

/** `readShapeText` の結果 */
export interface ShapeTextResult {
  /** 文字（ランの文字が 1 つも無い図形・`txBody` の無い図形は null） */
  text: DrawingText | null;
  /** 読み取った文字数（末尾に足した「…」を含む） */
  chars: number;
  /** 1 シートの合計に足す数＝文字数＋ラン（`a:br` を含む）の数＋段落の数 */
  units: number;
  /** 文字を切った・段落やランを捨てた（文字の無い図形では、上限で失った文字が無ければ立てない） */
  truncated: boolean;
}

/** `a:rPr`・`a:endParaRPr` の `sz`（1/100pt）を pt にする。無い・不正は 11。1〜400 に丸める */
function sizeOf(tag: string | null): number {
  const n = intAttr(tag === null ? undefined : parseAttrs(tag).sz);
  if (n === null) return DEFAULT_SIZE_PT;
  return Math.min(Math.max(n / 100, MIN_SIZE_PT), MAX_SIZE_PT);
}

/** 余白（EMU）。無い・不正は既定値。0〜2^31-1 に丸める */
function insetOf(s: string | undefined, dflt: number): number {
  const n = intAttr(s);
  return n === null ? dflt : Math.min(Math.max(n, 0), MAX_INSET);
}

/** 本文の末尾が高位サロゲートで切れていたら、その 1 文字を外す（文字を途中で切って壊さない） */
function trimLoneSurrogate(s: string): string {
  const c = s.length === 0 ? 0 : s.charCodeAt(s.length - 1);
  return c >= 0xd800 && c <= 0xdbff ? s.slice(0, -1) : s;
}

/** 文字の枠（`a:bodyPr`）を読む。無ければすべて既定値 */
function readBodyPr(
  txBody: string,
): Pick<DrawingText, 'anchor' | 'insets' | 'noWrap' | 'fontScale' | 'vert'> {
  const el = findElement(txBody, 'bodyPr');
  const a = el === null ? parseAttrs('') : parseAttrs(el.open);
  const anchor: TextAnchor = a.anchor === 'ctr' || a.anchor === 'b' ? a.anchor : 't';
  const vert = (a.vert !== undefined && VERTS.has(a.vert) ? a.vert : 'horz') as TextVert;
  let fontScale = 1;
  const fit = el?.body == null ? null : findOpenTag(el.body, 'normAutofit');
  if (fit !== null) {
    const f = intAttr(parseAttrs(fit).fontScale);
    if (f !== null) fontScale = Math.min(Math.max(f / 100000, 0.01), 1);
  }
  return {
    anchor,
    insets: {
      l: insetOf(a.lIns, DEFAULT_INSETS.l),
      t: insetOf(a.tIns, DEFAULT_INSETS.t),
      r: insetOf(a.rIns, DEFAULT_INSETS.r),
      b: insetOf(a.bIns, DEFAULT_INSETS.b),
    },
    noWrap: a.wrap === 'none',
    fontScale,
    vert,
  };
}

/** 段落の箇条書きの状態（`buAutoNum` の番号の続きを数える） */
interface NumState {
  type: string;
  start: number;
  next: number;
}

/** 段落の書式（`a:pPr`）：揃えと、箇条書きの文字（番号の続きは `num` を更新して数える） */
function readParaProps(
  open: string,
  body: string | null,
  num: { cur: NumState | null },
): { align: TextAlign; bullet: string | null } {
  const algn = parseAttrs(open).algn;
  const align = (algn !== undefined && ALIGNS.has(algn) ? algn : 'l') as TextAlign;
  let bullet: string | null = null;
  const none = body !== null && findOpenTag(body, 'buNone') !== null;
  const auto = body === null || none ? null : findOpenTag(body, 'buAutoNum');
  const chr = body === null || none || auto !== null ? null : findOpenTag(body, 'buChar');
  if (auto !== null) {
    const a = parseAttrs(auto);
    const type = a.type ?? 'arabicPeriod';
    const startAt = intAttr(a.startAt);
    const cur = num.cur;
    // 同じ種類が続く段落は番号を続ける。startAt が違う指定なら、そこから数え直す
    if (cur !== null && cur.type === type && (startAt === null || startAt === cur.start)) {
      bullet = `${cur.next}.`;
      cur.next++;
    } else {
      const start = startAt ?? 1;
      bullet = `${start}.`;
      num.cur = { type, start, next: start + 1 };
    }
    return { align, bullet };
  }
  num.cur = null;
  if (chr !== null) {
    // 箇条書きの記号は 1 文字（サロゲートペアを割らない）。制御文字は除く
    const ch = Array.from((parseAttrs(chr).char ?? '').replace(CONTROL_RE, ''))[0];
    bullet = ch ?? null;
  }
  return { align, bullet };
}

/** ラン・フィールド（`a:r`・`a:fld`）の中身から `a:t` の本文（実体参照・制御文字の処理前）を取り出す。無ければ空 */
function rawText(body: string | null): string {
  if (body === null) return '';
  T_RE.lastIndex = 0;
  const m = T_RE.exec(body);
  if (m === null || m[0].endsWith('/>')) return '';
  const start = m.index + m[0].length;
  const close = body.indexOf(`</${m[1] ?? ''}t>`, start);
  return close < 0 ? '' : body.slice(start, close);
}

/** 本文に空でない `a:t` があるか（文字の無い図形で、上限のために文字を失ったかの判定。線形時間） */
function hasNonEmptyText(txBody: string): boolean {
  T_RE.lastIndex = 0;
  for (let m = T_RE.exec(txBody); m !== null; m = T_RE.exec(txBody)) {
    if (m[0].endsWith('/>')) continue;
    const next = txBody[T_RE.lastIndex];
    if (next !== undefined && next !== '<') {
      T_RE.lastIndex = 0;
      return true;
    }
  }
  return false;
}

/** 書式（`a:rPr`）を読んでランにする。`rPrBody` は `a:rPr` の中身（色の塗り）、`fallback` は図形の文字の既定の色（`fontRef`） */
function makeRun(
  rPr: string | null,
  rPrBody: string | null,
  text: string,
  br: boolean,
  fallback: Rgba,
  theme: Theme,
): DrawingTextRun {
  const a = rPr === null ? parseAttrs('') : parseAttrs(rPr);
  return {
    br,
    text,
    sizePt: sizeOf(rPr),
    bold: boolAttr(a.b),
    italic: boolAttr(a.i),
    underline: a.u !== undefined && a.u !== 'none',
    strike: a.strike !== undefined && a.strike !== 'noStrike',
    color: readRunColor(rPrBody, fallback, theme),
  };
}

/**
 * 図形（`xdr:sp`）の中身から文字（`xdr:txBody`）を読む。
 * - `a:t` は `decodeEntities` で戻し、タブと改行以外の制御文字を除く。
 * - 上限（文字数 `maxChars`・段落 `maxParagraphs`・ラン `maxRuns`・1 シートの残り `maxUnits`）を超えたら、文字は切って末尾に「…」を付け、段落とランは以降を捨てて `truncated` にする。
 * - 空のラン（文字も `a:br` も無いもの）は一覧に入れず、ランの数にも数えない。
 * - ランの文字（長さ 1 以上）が 1 つも無い図形は `text` を null にする（文字の無い図形）。このとき `truncated` は、本文に空でない `a:t` があって上限で失ったときだけ立てる。
 * - ランの色（ISSUE-95）は `a:rPr` の塗り → `xdr:style` の `a:fontRef` の子の色 → テーマの `tx1`（`drawing-color.ts`）。
 * @param spBody `xdr:sp` の中身
 * @param theme ブックのテーマ（色の解決に使う。省くと既定の 12 色）
 */
export function readShapeText(
  spBody: string,
  caps: TextCaps,
  theme: Theme = DEFAULT_THEME,
): ShapeTextResult {
  const tx = findElement(spBody, 'txBody');
  if (tx?.body == null) return { text: null, chars: 0, units: 0, truncated: false };
  const txBody = tx.body;
  const fontColor = readFontColor(spBody, theme);

  const paragraphs: DrawingTextParagraph[] = [];
  const num: { cur: NumState | null } = { cur: null };
  let chars = 0;
  let runs = 0;
  let units = 0;
  let hasText = false;
  let truncated = false;
  let done = false;

  forEachElement(txBody, 'p', (_open, pBody) => {
    if (done) return;
    if (paragraphs.length >= caps.maxParagraphs || units + 1 > caps.maxUnits) {
      truncated = true;
      done = true;
      return;
    }
    const para: DrawingTextParagraph = {
      runs: [],
      align: 'l',
      bullet: null,
      endSizePt: DEFAULT_SIZE_PT,
    };
    paragraphs.push(para);
    units++;
    if (pBody === null) {
      num.cur = null;
      return;
    }
    let sawProps = false;
    for (let pos = 0; !done;) {
      PARA_TOKEN_RE.lastIndex = pos;
      const m = PARA_TOKEN_RE.exec(pBody);
      if (m === null) break;
      const open = m[0];
      const name = m[2]!;
      pos = m.index + open.length;
      let body: string | null = null;
      if (!open.endsWith('/>')) {
        const closeTag = `</${m[1] ?? ''}${name}>`;
        const close = pBody.indexOf(closeTag, pos);
        if (close < 0) break;
        body = pBody.slice(pos, close);
        pos = close + closeTag.length;
      }
      if (name === 'pPr') {
        if (!sawProps) {
          sawProps = true;
          const p = readParaProps(open, body, num);
          para.align = p.align;
          para.bullet = p.bullet;
        }
        continue;
      }
      if (name === 'endParaRPr') {
        para.endSizePt = sizeOf(open);
        continue;
      }
      // r・br・fld（空のラン＝文字も br も無いものは、一覧にも数にも入れない）
      const rEl = body === null ? null : findElement(body, 'rPr');
      const rPr = rEl?.open ?? null;
      const rPrBody = rEl?.body ?? null;
      let text = '';
      if (name !== 'br') {
        text = decodeEntities(rawText(body)).replace(CONTROL_RE, '');
        if (text.length === 0) continue;
      }
      if (runs >= caps.maxRuns || units + 1 > caps.maxUnits) {
        truncated = true;
        done = true;
        break;
      }
      if (name === 'br') {
        runs++;
        units++;
        para.runs.push(makeRun(rPr, rPrBody, '', true, fontColor, theme));
        continue;
      }
      const room = Math.min(caps.maxChars - chars, caps.maxUnits - units - 1);
      if (text.length > room) {
        truncated = true;
        done = true;
        text = trimLoneSurrogate(text.slice(0, Math.max(room, 0)));
        // 切った先に何も残らないとき（上限が 0・直前でちょうど切れた）は、「…」だけのランを作らない
        if (chars + text.length === 0) break;
        text += '…';
      }
      chars += text.length;
      units += 1 + text.length;
      hasText = true;
      runs++;
      para.runs.push(makeRun(rPr, rPrBody, text, false, fontColor, theme));
    }
    if (!sawProps) num.cur = null;
  });

  // 文字の無い図形は、上限のために失った文字があるときだけ truncated（段落・ランの上限を超えたのが空の要素だけなら立てない）
  if (!hasText)
    return { text: null, chars: 0, units: 0, truncated: truncated && hasNonEmptyText(txBody) };
  return { text: { paragraphs, ...readBodyPr(txBody) }, chars, units, truncated };
}
