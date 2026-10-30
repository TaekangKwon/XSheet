// 正: InterHub docs/plan/adr/0031-xlsx-charts.md（§決定 1・2・9）
/**
 * グラフ（`xl/charts/chartN.xml` の `c:chartSpace`）の読み取り（画面に依存しない。Node の vitest でも動く）。
 * G1（縦棒・横棒・折れ線・円・ドーナツ）だけを `ChartModel` にし、G1 の外の種類・構造は `unsupported`（破線枠）にする。
 * 値は `c:numCache`／`c:strCache`（無ければ `c:numLit`／`c:strLit`）の保存値だけを使い、数式（`c:f`）は計算しない。
 * 色は `color.ts`・`drawing-color.ts`（テーマ・変換・代表色）で `Rgba` に解く。明示が無い所は持たず、既定の色・範囲・配置は描画側の仕事（ISSUE-102）。
 * DOMParser に依存せず、開きタグ・閉じタグを 1 度ずつ数える走査（`childElements`）で線形時間に読む
 * （同じ名前の入れ子・`c:spPr` が各所にあっても、直下の子だけを引くので取り違えない）。
 * ファイルの文字列は系列名・項目・タイトルの文字と書式コードにだけ使い、型の他の値は数値と union の文字列に解く。解けない値は例外にせず既定に倒す。
 * 失敗（G1 外・上限超え・壊れ・描ける系列なし）は例外にせず理由つきの `ok: false`。例外のメッセージにファイル由来の文字列を入れない。
 */
import { fillElementColor } from './drawing-color';
import { intAttr } from './drawing-shape';
import { formatNumberByCode } from './format';
import { MAX_FORMAT_CODE_LENGTH, type XlsxLimits } from './limits';
import type { Theme } from './theme';
import type {
  ChartAxis,
  ChartBox,
  ChartFailReason,
  ChartLegend,
  ChartLine,
  ChartMarker,
  ChartMarkerSymbol,
  ChartModel,
  ChartPaint,
  ChartResult,
  ChartSeries,
  ChartTextStyle,
  ChartTickMark,
  ChartTitle,
  LineDash,
  ManualLayout,
} from './types';
import { NS, decodeText, findOpenTag, forEachElement, parseAttrs } from './xml';

/** 線の太さの上限（EMU。`drawing-shape.ts` と同じ。ECMA-376 の `ST_LineWidth` の最大値） */
const MAX_LINE_WIDTH = 20_116_800;

const DASHES: ReadonlySet<string> = new Set<LineDash>([
  'solid',
  'dot',
  'dash',
  'lgDash',
  'dashDot',
  'lgDashDot',
  'lgDashDotDot',
  'sysDash',
  'sysDot',
  'sysDashDot',
  'sysDashDotDot',
]);

/** 種類の要素の名前 → G1 の種類。`/Chart$/` に当たる要素でこの表に無いものは G1 外（面・散布図・3D など） */
const GROUP_TYPES: ReadonlyMap<string, ChartModel['type']> = new Map([
  ['barChart', 'bar'],
  ['lineChart', 'line'],
  ['pieChart', 'pie'],
  ['doughnutChart', 'doughnut'],
]);

/** 塗りの要素名（`c:spPr`・`a:ln`・`a:defRPr` の直下） */
const FILL_NAMES: ReadonlySet<string> = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'pattFill',
  'blipFill',
]);

const MARKER_SYMBOLS: readonly ChartMarkerSymbol[] = [
  'auto',
  'circle',
  'dash',
  'diamond',
  'dot',
  'none',
  'picture',
  'plus',
  'square',
  'star',
  'triangle',
  'x',
];

/** 数値の文字（符号・小数・指数。`NaN`・`Infinity`・空文字・16 進は数値にしない） */
const NUM_RE = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

/** 開き・閉じタグ 1 つ（名前空間の接頭辞可。`[^<>]*>` で打ち切る）。`childElements` の走査用 */
const TAG_RE = new RegExp(String.raw`<(/?)${NS}([A-Za-z_][\w.-]*)(?=[\s>/])[^<>]*>`, 'g');

/** 失敗の理由を運ぶ内部の例外（`parseChart` が捕まえて `ok: false` にする。メッセージは固定でファイル由来の文字列を入れない） */
class ChartAbort extends Error {
  readonly reason: ChartFailReason;
  constructor(reason: ChartFailReason) {
    super('chart');
    this.name = 'ChartAbort';
    this.reason = reason;
  }
}

/** 直下の子要素 1 つ（名前は接頭辞を除いた局所名。自己閉じ・閉じの無いものの `body` は null） */
interface Child {
  name: string;
  open: string;
  body: string | null;
}

/**
 * `xml` の直下（深さ 0）の子要素を出現順に並べる。開き・閉じタグを正規表現で 1 度ずつ数えるだけなので、時間は長さに比例する
 * （閉じタグの名前は照合せず、深さだけを数える。閉じの無い子は `body` が null）。
 */
function childElements(xml: string): Child[] {
  const out: Child[] = [];
  TAG_RE.lastIndex = 0;
  let depth = 0;
  let name = '';
  let open = '';
  let bodyStart = 0;
  for (let m = TAG_RE.exec(xml); m !== null; m = TAG_RE.exec(xml)) {
    if (m[1] === '/') {
      if (depth === 0) continue;
      depth--;
      if (depth === 0) out.push({ name, open, body: xml.slice(bodyStart, m.index) });
    } else if (m[0].endsWith('/>')) {
      if (depth === 0) out.push({ name: m[2]!, open: m[0], body: null });
    } else {
      if (depth === 0) {
        name = m[2]!;
        open = m[0];
        bodyStart = m.index + m[0].length;
      }
      depth++;
    }
  }
  TAG_RE.lastIndex = 0;
  if (depth > 0) out.push({ name, open, body: null });
  return out;
}

/** 子の一覧から名前で最初のものを引く */
const find = (kids: readonly Child[], name: string): Child | undefined =>
  kids.find((k) => k.name === name);

/** 子の一覧（`body` が無ければ空） */
const kidsOf = (el: Child | undefined): Child[] =>
  el === undefined || el.body === null ? [] : childElements(el.body);

/** 要素の `val` 属性 */
const valOf = (el: Child | undefined): string | undefined =>
  el === undefined ? undefined : parseAttrs(el.open).val;

/** 数値の文字（前後の空白可）を有限の数にする。数値でなければ null */
function numberOf(s: string | undefined): number | null {
  if (s === undefined) return null;
  const t = s.trim();
  if (!NUM_RE.test(t)) return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

/** 要素の `val` を数にする。無い・数値でなければ undefined */
function numVal(el: Child | undefined): number | undefined {
  return numberOf(valOf(el)) ?? undefined;
}

/** 要素の `val` を有限の正の数にする（目盛の間隔用）。無い・数値でない・0 以下なら undefined */
function positiveVal(el: Child | undefined): number | undefined {
  const n = numVal(el);
  return n !== undefined && n > 0 ? n : undefined;
}

/** 要素の `val` を整数にし、範囲に丸める。無い・整数でなければ undefined */
function intVal(el: Child | undefined, min: number, max: number): number | undefined {
  const n = intAttr(valOf(el));
  return n === null ? undefined : Math.min(Math.max(n, min), max);
}

/** 要素の `val` を 0 以上の整数（点・項目の位置）にする。無い・整数でない・負なら undefined */
function indexVal(el: Child | undefined): number | undefined {
  const n = intAttr(valOf(el));
  return n !== null && n >= 0 ? n : undefined;
}

/** 要素の真偽（`c:delete`・`c:overlay` など）。要素が無ければ `dflt`、`val` が無ければ真、`1`・`true` が真、`0`・`false` が偽、それ以外は `dflt` */
function boolVal(el: Child | undefined, dflt: boolean): boolean {
  if (el === undefined) return dflt;
  const v = valOf(el);
  if (v === undefined || v === '1' || v === 'true') return true;
  if (v === '0' || v === 'false') return false;
  return dflt;
}

/** 要素の `val` を、許す語のどれかなら返す（それ以外・無いときは undefined） */
function enumVal<T extends string>(el: Child | undefined, allowed: readonly T[]): T | undefined {
  const v = valOf(el);
  return v !== undefined && (allowed as readonly string[]).includes(v) ? (v as T) : undefined;
}

/** 1 グラフの読み取りに持ち回るもの */
interface Ctx {
  theme: Theme;
  limits: XlsxLimits;
  date1904: boolean;
  /** これまでに数えた系列名・項目・タイトルの文字数 */
  chars: number;
}

/** 文字数を数える。上限（`maxChartTextChars`）を超えたら `too-large` */
function charge(ctx: Ctx, s: string): string {
  ctx.chars += s.length;
  if (ctx.chars > ctx.limits.maxChartTextChars) throw new ChartAbort('too-large');
  return s;
}

// ---- 色・線・面 ----

/** 子の一覧の最初の塗りを `ChartPaint` にする。塗りが無い・色が解けない・`a:grpFill` などは undefined */
function paintOf(kids: readonly Child[], theme: Theme): ChartPaint | undefined {
  const el = kids.find((k) => FILL_NAMES.has(k.name));
  if (el === undefined) return undefined;
  if (el.name === 'noFill') return { kind: 'none' };
  const color = fillElementColor(el, theme);
  return color === null ? undefined : { kind: 'solid', color };
}

/** `a:ln` を線にする（太さ・塗り・線種。明示のあったものだけ持つ） */
function lineOf(ln: Child, theme: Theme): ChartLine {
  const out: ChartLine = {};
  const w = intAttr(parseAttrs(ln.open).w);
  if (w !== null) out.widthEmu = Math.min(Math.max(w, 0), MAX_LINE_WIDTH);
  const kids = kidsOf(ln);
  const paint = paintOf(kids, theme);
  if (paint !== undefined) out.paint = paint;
  const dash = find(kids, 'prstDash');
  if (dash !== undefined) {
    const v = valOf(dash);
    out.dash = (v !== undefined && DASHES.has(v) ? v : 'solid') as LineDash;
  }
  return out;
}

/** `c:spPr` を面（塗りと枠線）にする。`spPr` が無ければ空 */
function boxOf(spPr: Child | undefined, theme: Theme): ChartBox {
  const out: ChartBox = {};
  const kids = kidsOf(spPr);
  const fill = paintOf(kids, theme);
  if (fill !== undefined) out.fill = fill;
  const ln = find(kids, 'ln');
  if (ln !== undefined) out.line = lineOf(ln, theme);
  return out;
}

// ---- 文字 ----

/** `a:defRPr`・`a:rPr` の文字の大きさ・太字・色（明示のあったものだけ） */
function runStyleOf(el: Child | undefined, theme: Theme): ChartTextStyle {
  const out: ChartTextStyle = {};
  if (el === undefined) return out;
  const a = parseAttrs(el.open);
  const sz = intAttr(a.sz);
  if (sz !== null) out.sizePt = Math.min(Math.max(sz / 100, 1), 400);
  if (a.b === '1' || a.b === 'true') out.bold = true;
  else if (a.b === '0' || a.b === 'false') out.bold = false;
  const fill = kidsOf(el).find((k) => FILL_NAMES.has(k.name));
  if (fill !== undefined) {
    if (fill.name === 'noFill') out.color = { rgb: 0, alpha: 0 };
    else {
      const c = fillElementColor(fill, theme);
      if (c !== null) out.color = c;
    }
  }
  return out;
}

/** 文字の本体（`c:rich`・`c:txPr`）：段落ごとの文字（`a:br` でも分ける）と、文字の見た目（先頭の `a:defRPr`・先頭のランの `a:rPr`・`a:bodyPr rot`） */
function readTextBody(ctx: Ctx, body: string | null): { lines: string[]; style: ChartTextStyle } {
  const kids = body === null ? [] : childElements(body);
  const lines: string[] = [];
  let defStyle: ChartTextStyle | null = null;
  let runStyle: ChartTextStyle | null = null;
  for (const p of kids) {
    if (p.name !== 'p') continue;
    const pk = kidsOf(p);
    if (defStyle === null) {
      const def = find(kidsOf(find(pk, 'pPr')), 'defRPr');
      if (def !== undefined) defStyle = runStyleOf(def, ctx.theme);
    }
    let cur = '';
    for (const c of pk) {
      if (c.name === 'br') {
        lines.push(charge(ctx, cur));
        cur = '';
      } else if (c.name === 'r' || c.name === 'fld') {
        const rk = kidsOf(c);
        if (runStyle === null) {
          const rPr = find(rk, 'rPr');
          if (rPr !== undefined) runStyle = runStyleOf(rPr, ctx.theme);
        }
        const t = find(rk, 't');
        cur += t?.body == null ? '' : decodeText(t.body);
      }
    }
    lines.push(charge(ctx, cur));
  }
  const style: ChartTextStyle = { ...defStyle, ...runStyle };
  const rot = intAttr(parseAttrs(find(kids, 'bodyPr')?.open ?? '').rot);
  if (rot !== null && Math.abs(rot / 60000) <= 90) style.rotDeg = rot / 60000;
  return { lines, style };
}

/** `c:txPr` の文字の見た目（無ければ空） */
function readTxPr(ctx: Ctx, el: Child | undefined): ChartTextStyle {
  return el === undefined ? {} : readTextBody(ctx, el.body).style;
}

// ---- 保存値（c:numCache・c:strCache・c:numLit・c:strLit・c:multiLvlStrCache の c:lvl） ----

/** 保存値 1 かたまり（キャッシュ・リテラル）。点は疎（`c:pt` のある位置だけ） */
interface Cache {
  /** `c:ptCount`（無ければ最大の `c:pt idx` + 1） */
  count: number;
  pts: { idx: number; text: string }[];
  /** `c:formatCode`（無ければ undefined） */
  formatCode?: string;
}

/** `c:pt` の `c:v` の本文（実体参照は戻さない。無ければ空） */
function pointText(body: string | null): string {
  if (body === null) return '';
  const found: { text: string | null } = { text: null };
  forEachElement(body, 'v', (_open, b) => {
    found.text ??= b ?? '';
  });
  return found.text ?? '';
}

/**
 * 保存値のかたまりの本文（`c:numCache` の中身など）を読む。点の数が上限（`maxChartPoints`）を超えたら `too-large`（配列を作る前に断る）。
 * @param hint `c:ptCount` が本文に無いとき（多段の項目の `c:lvl`）に使う、外側の `c:ptCount`
 */
function readCache(ctx: Ctx, body: string, hint: number | null = null): Cache {
  const max = ctx.limits.maxChartPoints;
  const tag = findOpenTag(body, 'ptCount');
  const declared = (tag === null ? null : intAttr(parseAttrs(tag).val)) ?? hint;
  if (declared !== null && declared > max) throw new ChartAbort('too-large');
  const pts: { idx: number; text: string }[] = [];
  let maxIdx = -1;
  forEachElement(body, 'pt', (open, b) => {
    const idx = intAttr(parseAttrs(open).idx);
    if (idx === null || idx < 0 || (declared !== null && idx >= declared)) return;
    if (pts.length >= max) throw new ChartAbort('too-large');
    pts.push({ idx, text: pointText(b) });
    if (idx > maxIdx) maxIdx = idx;
  });
  const count = declared !== null && declared >= 0 ? declared : maxIdx + 1;
  if (count > max) throw new ChartAbort('too-large');
  const code: { text: string | null } = { text: null };
  forEachElement(body, 'formatCode', (_open, b) => {
    code.text ??= b === null ? '' : decodeText(b);
  });
  return code.text === null ? { count, pts } : { count, pts, formatCode: code.text };
}

/** 書式コードを、長すぎれば General にする（ssf に渡す前の ReDoS 対策。`format.ts` の扱いと同じ） */
const safeCode = (code: string): string =>
  code.length > MAX_FORMAT_CODE_LENGTH ? 'General' : code;

/** `c:val` の保存値（`c:numRef` の `c:numCache`、無ければ `c:numLit`）。保存値の無いものは null */
function readValueCache(ctx: Ctx, val: Child | undefined): Cache | null {
  for (const k of kidsOf(val)) {
    if (k.name === 'numLit' && k.body !== null) return readCache(ctx, k.body);
    if (k.name === 'numRef') {
      const cache = find(kidsOf(k), 'numCache');
      if (cache?.body != null) return readCache(ctx, cache.body);
    }
  }
  return null;
}

/**
 * 点の値を数にする。長さは最初から項目の数 `n` で作り、`n` 以上の位置の点は捨てる
 * （`c:ptCount` が大きく項目が少ない系列を 255 本並べても、大きな配列を作らない）。数値にならない値・無い点は null
 */
function numbersOf(cache: Cache, n: number): (number | null)[] {
  const out: (number | null)[] = new Array<number | null>(n).fill(null);
  for (const p of cache.pts) if (p.idx < n) out[p.idx] = numberOf(p.text);
  return out;
}

/**
 * 点を表示用の文字にする（長さ `count`・無い点は空）。数値の点は書式を ssf で当てる（数値でない文字はそのまま）。
 * 書式は `override`（項目軸の `c:numFmt`。`sourceLinked="0"` のとき）があればそれ、無ければ保存値の `c:formatCode`
 */
function textsOf(ctx: Ctx, cache: Cache, numeric: boolean, override?: string): string[] {
  const out: string[] = new Array<string>(cache.count).fill('');
  const code = safeCode(override ?? cache.formatCode ?? 'General');
  for (const p of cache.pts) {
    const n = numeric ? numberOf(p.text) : null;
    out[p.idx] = charge(
      ctx,
      n === null ? decodeText(p.text) : formatNumberByCode(code, n, ctx.date1904),
    );
  }
  return out;
}

/**
 * 項目（`c:cat`）の文字。`override` は項目軸の書式（数値の項目にだけ効く）。`c:strRef`・`c:strLit`・`c:numRef`（書式を当てる）・`c:numLit`・`c:multiLvlStrRef`（最も内側＝最初の `c:lvl` だけ。暫定）。
 * 保存値が無い・点が 0 個なら null
 */
function readCategories(ctx: Ctx, cat: Child | undefined, override?: string): string[] | null {
  for (const k of kidsOf(cat)) {
    let cache: Cache | null = null;
    let numeric = false;
    if (k.name === 'strLit' || k.name === 'numLit') {
      numeric = k.name === 'numLit';
      if (k.body !== null) cache = readCache(ctx, k.body);
    } else if (k.name === 'strRef' || k.name === 'numRef') {
      numeric = k.name === 'numRef';
      const c = find(kidsOf(k), numeric ? 'numCache' : 'strCache');
      if (c?.body != null) cache = readCache(ctx, c.body);
    } else if (k.name === 'multiLvlStrRef') {
      const c = find(kidsOf(k), 'multiLvlStrCache');
      if (c?.body != null) {
        const tag = findOpenTag(c.body, 'ptCount');
        const hint = tag === null ? null : intAttr(parseAttrs(tag).val);
        const lvl = find(childElements(c.body), 'lvl');
        if (lvl?.body != null) cache = readCache(ctx, lvl.body, hint);
      }
    }
    if (cache !== null && cache.count > 0) return textsOf(ctx, cache, numeric, override);
  }
  return null;
}

/** 系列名（`c:tx` の `c:strRef` の保存値、または `c:v`）。無ければ空 */
function readSeriesName(ctx: Ctx, tx: Child | undefined): string {
  for (const k of kidsOf(tx)) {
    if (k.name === 'v') return charge(ctx, k.body === null ? '' : decodeText(k.body));
    if (k.name === 'strRef') {
      const cache = find(kidsOf(k), 'strCache');
      if (cache?.body != null) {
        const first = readCache(ctx, cache.body).pts[0];
        return charge(ctx, first === undefined ? '' : decodeText(first.text));
      }
    }
  }
  return '';
}

// ---- タイトル・凡例・配置 ----

/** `c:layout` の `c:manualLayout`。明示の値が 1 つも無ければ undefined */
function readLayout(el: Child | undefined): ManualLayout | undefined {
  const ml = find(kidsOf(el), 'manualLayout');
  if (ml === undefined) return undefined;
  const kids = kidsOf(ml);
  const out: ManualLayout = {};
  const target = enumVal(find(kids, 'layoutTarget'), ['inner', 'outer'] as const);
  if (target !== undefined) out.target = target;
  const modes = ['edge', 'factor'] as const;
  const xMode = enumVal(find(kids, 'xMode'), modes);
  const yMode = enumVal(find(kids, 'yMode'), modes);
  const wMode = enumVal(find(kids, 'wMode'), modes);
  const hMode = enumVal(find(kids, 'hMode'), modes);
  if (xMode !== undefined) out.xMode = xMode;
  if (yMode !== undefined) out.yMode = yMode;
  if (wMode !== undefined) out.wMode = wMode;
  if (hMode !== undefined) out.hMode = hMode;
  const x = numVal(find(kids, 'x'));
  const y = numVal(find(kids, 'y'));
  const w = numVal(find(kids, 'w'));
  const h = numVal(find(kids, 'h'));
  if (x !== undefined) out.x = x;
  if (y !== undefined) out.y = y;
  if (w !== undefined) out.w = w;
  if (h !== undefined) out.h = h;
  return Object.keys(out).length === 0 ? undefined : out;
}

/** タイトル（`c:title`）。文字は `c:rich` の段落、無ければ `c:strRef` の保存値。文字が無ければ `lines` は空（自動のタイトル） */
function readTitle(ctx: Ctx, el: Child): ChartTitle {
  const kids = kidsOf(el);
  let lines: string[] = [];
  let style: ChartTextStyle = {};
  const txKids = kidsOf(find(kids, 'tx'));
  const rich = find(txKids, 'rich');
  if (rich !== undefined) {
    const r = readTextBody(ctx, rich.body);
    lines = r.lines.every((l) => l === '') ? [] : r.lines;
    style = r.style;
  } else {
    const cache = find(kidsOf(find(txKids, 'strRef')), 'strCache');
    const first = cache?.body == null ? undefined : readCache(ctx, cache.body).pts[0];
    if (first !== undefined && first.text !== '') lines = [charge(ctx, decodeText(first.text))];
  }
  const out: ChartTitle = {
    lines,
    text: { ...readTxPr(ctx, find(kids, 'txPr')), ...style },
    overlay: boolVal(find(kids, 'overlay'), false),
  };
  const layout = readLayout(find(kids, 'layout'));
  if (layout !== undefined) out.layout = layout;
  return out;
}

/** 凡例（`c:legend`） */
function readLegend(ctx: Ctx, el: Child): ChartLegend {
  const kids = kidsOf(el);
  const deleted: number[] = [];
  for (const e of kids) {
    if (e.name !== 'legendEntry') continue;
    const ek = kidsOf(e);
    const idx = indexVal(find(ek, 'idx'));
    if (idx !== undefined && boolVal(find(ek, 'delete'), false)) deleted.push(idx);
  }
  const out: ChartLegend = {
    pos: enumVal(find(kids, 'legendPos'), ['r', 'l', 't', 'b', 'tr'] as const) ?? 'r',
    overlay: boolVal(find(kids, 'overlay'), false),
    deleted,
    text: readTxPr(ctx, find(kids, 'txPr')),
  };
  const layout = readLayout(find(kids, 'layout'));
  if (layout !== undefined) out.layout = layout;
  return out;
}

// ---- 軸 ----

/** 目盛線（`c:majorGridlines`・`c:minorGridlines`）。要素があれば線あり（`c:spPr` の `a:ln` の明示だけ持つ） */
function readGridlines(el: Child | undefined, theme: Theme): ChartLine | undefined {
  if (el === undefined) return undefined;
  const ln = find(kidsOf(find(kidsOf(el), 'spPr')), 'ln');
  return ln === undefined ? {} : lineOf(ln, theme);
}

/** 軸 1 本（`c:catAx`・`c:dateAx`・`c:valAx`）を読む。`kids` はその直下の子 */
function readAxis(ctx: Ctx, kind: ChartAxis['kind'], kids: readonly Child[]): ChartAxis {
  const scaling = kidsOf(find(kids, 'scaling'));
  const out: ChartAxis = {
    kind,
    deleted: boolVal(find(kids, 'delete'), false),
    orientation: enumVal(find(scaling, 'orientation'), ['minMax', 'maxMin'] as const) ?? 'minMax',
    text: readTxPr(ctx, find(kids, 'txPr')),
  };
  const pos = enumVal(find(kids, 'axPos'), ['b', 'l', 'r', 't'] as const);
  if (pos !== undefined) out.pos = pos;
  const min = numVal(find(scaling, 'min'));
  const max = numVal(find(scaling, 'max'));
  if (min !== undefined) out.min = min;
  if (max !== undefined) out.max = max;
  const majorUnit = positiveVal(find(kids, 'majorUnit'));
  const minorUnit = positiveVal(find(kids, 'minorUnit'));
  if (majorUnit !== undefined) out.majorUnit = majorUnit;
  if (minorUnit !== undefined) out.minorUnit = minorUnit;
  const crosses = enumVal(find(kids, 'crosses'), ['autoZero', 'min', 'max'] as const);
  if (crosses !== undefined) out.crosses = crosses;
  const crossesAt = numVal(find(kids, 'crossesAt'));
  if (crossesAt !== undefined) out.crossesAt = crossesAt;
  const crossBetween = enumVal(find(kids, 'crossBetween'), ['between', 'midCat'] as const);
  if (crossBetween !== undefined) out.crossBetween = crossBetween;
  const major = readGridlines(find(kids, 'majorGridlines'), ctx.theme);
  const minor = readGridlines(find(kids, 'minorGridlines'), ctx.theme);
  if (major !== undefined) out.majorGridlines = major;
  if (minor !== undefined) out.minorGridlines = minor;
  const marks = ['none', 'in', 'out', 'cross'] as const satisfies readonly ChartTickMark[];
  const majorTick = enumVal(find(kids, 'majorTickMark'), marks);
  const minorTick = enumVal(find(kids, 'minorTickMark'), marks);
  if (majorTick !== undefined) out.majorTickMark = majorTick;
  if (minorTick !== undefined) out.minorTickMark = minorTick;
  const lblPos = enumVal(find(kids, 'tickLblPos'), ['nextTo', 'low', 'high', 'none'] as const);
  if (lblPos !== undefined) out.tickLblPos = lblPos;
  const skip = intVal(find(kids, 'tickLblSkip'), 1, Number.MAX_SAFE_INTEGER);
  if (skip !== undefined) out.tickLblSkip = skip;
  const numFmt = find(kids, 'numFmt');
  if (numFmt !== undefined) {
    const a = parseAttrs(numFmt.open);
    if (a.formatCode !== undefined)
      out.numFmt = {
        code: safeCode(a.formatCode),
        sourceLinked: a.sourceLinked === '1' || a.sourceLinked === 'true',
      };
  }
  const title = find(kids, 'title');
  if (title !== undefined) out.title = readTitle(ctx, title);
  const ln = find(kidsOf(find(kids, 'spPr')), 'ln');
  if (ln !== undefined) out.line = lineOf(ln, ctx.theme);
  return out;
}

// ---- 系列 ----

/** 系列のマーカー（`c:marker`） */
function readMarker(el: Child, theme: Theme): ChartMarker {
  const kids = kidsOf(el);
  const out: ChartMarker = {};
  const symbol = enumVal(find(kids, 'symbol'), MARKER_SYMBOLS);
  if (symbol !== undefined) out.symbol = symbol;
  const size = intVal(find(kids, 'size'), 2, 72);
  if (size !== undefined) out.size = size;
  const box = boxOf(find(kids, 'spPr'), theme);
  if (box.fill !== undefined) out.fill = box.fill;
  if (box.line !== undefined) out.line = box.line;
  return out;
}

/** 点ごとの塗りの上書き（`c:dPt` の `c:spPr` の塗り）。点の位置は 0 以上で項目の数 `n` 未満、数は `maxChartPoints` まで */
function readPoints(
  kids: readonly Child[],
  ctx: Ctx,
  n: number,
): Map<number, ChartPaint> | undefined {
  const out = new Map<number, ChartPaint>();
  for (const d of kids) {
    if (d.name !== 'dPt') continue;
    const dk = kidsOf(d);
    const idx = indexVal(find(dk, 'idx'));
    const fill = boxOf(find(dk, 'spPr'), ctx.theme).fill;
    if (idx === undefined || idx >= n || fill === undefined) continue;
    if (out.size >= ctx.limits.maxChartPoints) break;
    out.set(idx, fill);
  }
  return out.size === 0 ? undefined : out;
}

/** 系列の読み取りの途中の姿（値を読んだ後・項目を決める前） */
interface RawSeries {
  /** `c:idx`（無い・負・整数でなければ出現の位置） */
  idx: number;
  order: number;
  position: number;
  kids: Child[];
  cache: Cache;
}

/** 系列の見た目（名前・色・マーカー・点）を読んで `ChartSeries` にする */
function buildSeries(ctx: Ctx, raw: RawSeries, n: number): ChartSeries {
  const kids = raw.kids;
  const out: ChartSeries = {
    name: readSeriesName(ctx, find(kids, 'tx')),
    idx: raw.idx,
    values: numbersOf(raw.cache, n),
  };
  const box = boxOf(find(kids, 'spPr'), ctx.theme);
  if (box.fill !== undefined) out.fill = box.fill;
  if (box.line !== undefined) out.line = box.line;
  const marker = find(kids, 'marker');
  if (marker !== undefined) out.marker = readMarker(marker, ctx.theme);
  const smooth = find(kids, 'smooth');
  if (smooth !== undefined) out.smooth = boolVal(smooth, true);
  const points = readPoints(kids, ctx, n);
  if (points !== undefined) out.points = points;
  if (raw.cache.formatCode !== undefined) out.formatCode = safeCode(raw.cache.formatCode);
  return out;
}

// ---- 全体 ----

/**
 * `c:chartSpace` の `c:style`（直下、または `mc:AlternateContent` の `mc:Fallback` の中。Choice の `c14:style` は {@link readC14Style} が別に読む）。
 * 1〜48 はそのまま、101〜148（`c14:style` の形）は 100 を引く。範囲外は丸めず undefined
 */
function readStyle(spaceKids: readonly Child[]): number | undefined {
  let el = find(spaceKids, 'style');
  if (el === undefined) {
    const ac = find(spaceKids, 'AlternateContent');
    el = find(kidsOf(find(kidsOf(ac), 'Fallback')), 'style');
  }
  const v = intAttr(valOf(el));
  if (v === null) return undefined;
  if (v >= 1 && v <= 48) return v;
  return v >= 101 && v <= 148 ? v - 100 : undefined;
}

/**
 * `c14:style`：`c:chartSpace` 直下の `mc:AlternateContent` の `mc:Choice` の中の `c14:style` の `val`（整数）。
 * 要素が無い・整数でなければ undefined（値は描画に使わない。「持つか」だけが既定の表を分ける。ISSUE-109）
 */
function readC14Style(spaceKids: readonly Child[]): number | undefined {
  for (const ac of spaceKids) {
    if (ac.name !== 'AlternateContent') continue;
    for (const choice of kidsOf(ac)) {
      if (choice.name !== 'Choice') continue;
      const v = intAttr(valOf(find(kidsOf(choice), 'style')));
      if (v !== null) return v;
    }
  }
  return undefined;
}

/** G1 の種類の要素のうち、同じ棒・折れ線の並べ方の語から `grouping` を決める */
function groupingOf(type: ChartModel['type'], el: Child | undefined): ChartModel['grouping'] {
  if (type === 'pie' || type === 'doughnut') return 'standard';
  if (type === 'bar')
    return enumVal(el, ['clustered', 'stacked', 'percentStacked'] as const) ?? 'clustered';
  return enumVal(el, ['standard', 'stacked', 'percentStacked'] as const) ?? 'standard';
}

/** `c:dispBlanksAs`：要素が無ければ `gap`、`val` の無い要素は ECMA の既定 `zero`、未知の値は `gap` */
function blanksOf(el: Child | undefined): ChartModel['dispBlanksAs'] {
  if (el === undefined) return 'gap';
  if (valOf(el) === undefined) return 'zero';
  return enumVal(el, ['gap', 'zero', 'span'] as const) ?? 'gap';
}

/** 軸の要素名 */
const AXIS_NAMES: ReadonlySet<string> = new Set(['catAx', 'dateAx', 'valAx', 'serAx']);

/** 本体。失敗は `ChartAbort` を投げる（`parseChart` が結果にする） */
function parse(xml: string, theme: Theme, limits: XlsxLimits, date1904: boolean): ChartModel {
  const ctx: Ctx = { theme, limits, date1904, chars: 0 };

  // 骨組み：c:chartSpace > c:chart > c:plotArea。どれかが読めない（閉じが無い・無い）なら壊れている
  const space = find(childElements(xml), 'chartSpace');
  if (space === undefined || space.body === null) throw new ChartAbort('broken');
  const spaceKids = childElements(space.body);
  const chart = find(spaceKids, 'chart');
  if (chart === undefined || chart.body === null) throw new ChartAbort('broken');
  const chartKids = childElements(chart.body);
  const plot = find(chartKids, 'plotArea');
  if (plot === undefined || plot.body === null) throw new ChartAbort('broken');
  const plotKids = childElements(plot.body);

  // G1 の判定（§決定 1・P2）：種類の要素がちょうど 1 つで G1 の 4 種、数値軸が 1 本以下、対数軸と 3D 表示が無い
  if (find(chartKids, 'view3D') !== undefined) throw new ChartAbort('unsupported');
  const groups = plotKids.filter((k) => k.name.endsWith('Chart'));
  const group = groups.length === 1 ? groups[0]! : null;
  const type = group === null ? undefined : GROUP_TYPES.get(group.name);
  if (group === null || type === undefined) throw new ChartAbort('unsupported');
  const axes = plotKids
    .filter((k) => AXIS_NAMES.has(k.name))
    .map((el) => ({ el, kids: kidsOf(el) }));
  if (axes.filter((a) => a.el.name === 'valAx').length > 1) throw new ChartAbort('unsupported');
  for (const a of axes) {
    if (find(kidsOf(find(a.kids, 'scaling')), 'logBase') !== undefined)
      throw new ChartAbort('unsupported');
  }

  // 系列：c:order の順（無ければ出現順）。保存値の無い系列は除く（P3）。上限は配列を作る前に確かめる
  const groupKids = kidsOf(group);
  const serEls = groupKids.filter((k) => k.name === 'ser');
  if (serEls.length > limits.maxChartSeries) throw new ChartAbort('too-large');
  const raws: RawSeries[] = [];
  serEls.forEach((el, position) => {
    const kids = kidsOf(el);
    const cache = readValueCache(ctx, find(kids, 'val'));
    if (cache === null || cache.count === 0) return;
    raws.push({
      idx: indexVal(find(kids, 'idx')) ?? position,
      order: intAttr(valOf(find(kids, 'order'))) ?? position,
      position,
      kids,
      cache,
    });
  });
  raws.sort((a, b) => a.order - b.order || a.position - b.position);
  if (raws.length === 0) throw new ChartAbort('no-data');

  // 項目：最初の系列の c:cat。無ければ 1, 2, 3…。長さは項目の数（無ければ最長の系列）に揃える
  // 項目軸の書式が `sourceLinked="0"`（明示）なら、数値の項目はその書式で作る（保存値の書式より優先）
  const catFmt = (
    type === 'bar' || type === 'line'
      ? axes.find((a) => a.el.name === 'catAx' || a.el.name === 'dateAx')
      : undefined
  )?.kids;
  const fmt = find(catFmt ?? [], 'numFmt');
  const fmtAttrs = fmt === undefined ? undefined : parseAttrs(fmt.open);
  const override =
    fmtAttrs?.formatCode !== undefined &&
    (fmtAttrs.sourceLinked === '0' || fmtAttrs.sourceLinked === 'false')
      ? fmtAttrs.formatCode
      : undefined;
  const cats = readCategories(ctx, find(raws[0]!.kids, 'cat'), override);
  const n = cats !== null ? cats.length : Math.max(...raws.map((r) => r.cache.count));
  if (n * raws.length > limits.maxChartPoints) throw new ChartAbort('too-large');
  const categories = cats ?? Array.from({ length: n }, (_, i) => String(i + 1));
  const series = raws.map((r) => buildSeries(ctx, r, n));

  const out: ChartModel = {
    type,
    grouping: groupingOf(type, find(groupKids, 'grouping')),
    varyColors: boolVal(find(groupKids, 'varyColors'), true),
    categories,
    series,
    autoTitleDeleted: boolVal(find(chartKids, 'autoTitleDeleted'), false),
    chartArea: boxOf(find(spaceKids, 'spPr'), theme),
    plotArea: boxOf(find(plotKids, 'spPr'), theme),
    text: readTxPr(ctx, find(spaceKids, 'txPr')),
    dispBlanksAs: blanksOf(find(chartKids, 'dispBlanksAs')),
  };
  if (type === 'bar') {
    out.barDir = enumVal(find(groupKids, 'barDir'), ['col', 'bar'] as const) ?? 'col';
    const gapWidth = intVal(find(groupKids, 'gapWidth'), 0, 500);
    const overlap = intVal(find(groupKids, 'overlap'), -100, 100);
    if (gapWidth !== undefined) out.gapWidth = gapWidth;
    if (overlap !== undefined) out.overlap = overlap;
  }
  if (type === 'pie' || type === 'doughnut') {
    const angle = intVal(find(groupKids, 'firstSliceAng'), 0, 360);
    if (angle !== undefined) out.firstSliceAng = angle;
  }
  if (type === 'doughnut') {
    const hole = intVal(find(groupKids, 'holeSize'), 10, 90);
    if (hole !== undefined) out.holeSize = hole;
  }
  if (type === 'line') {
    const marker = find(groupKids, 'marker');
    if (marker !== undefined) out.lineMarkers = boolVal(marker, true);
  }
  if (type === 'bar' || type === 'line') {
    const cat = axes.find((a) => a.el.name === 'catAx' || a.el.name === 'dateAx');
    const val = axes.find((a) => a.el.name === 'valAx');
    if (cat !== undefined)
      out.catAxis = readAxis(ctx, cat.el.name === 'dateAx' ? 'date' : 'cat', cat.kids);
    if (val !== undefined) out.valAxis = readAxis(ctx, 'val', val.kids);
  }
  const title = find(chartKids, 'title');
  if (title !== undefined) out.title = readTitle(ctx, title);
  const legend = find(chartKids, 'legend');
  if (legend !== undefined) out.legend = readLegend(ctx, legend);
  const plotLayout = readLayout(find(plotKids, 'layout'));
  if (plotLayout !== undefined) out.plotLayout = plotLayout;
  const style = readStyle(spaceKids);
  if (style !== undefined) out.style = style;
  const c14Style = readC14Style(spaceKids);
  if (c14Style !== undefined) out.c14Style = c14Style;
  // c:roundedCorners：要素が無ければ持たない（既定は描画側）。val が無い・未知の値は真（CT_Boolean の既定）
  const rc = find(spaceKids, 'roundedCorners');
  if (rc !== undefined) out.roundedCorners = boolVal(rc, true);
  return out;
}

/**
 * グラフの部品（`chartN.xml` の全文）を `ChartModel` にする。
 * G1 の外・上限超え・壊れ・描ける系列なしは、例外にせず理由つきの `ok: false`（表示側はどの理由でも破線枠）。
 * @param theme ブックのテーマの 12 色（色の解決用）
 * @param limits `maxChartSeries`・`maxChartPoints`・`maxChartTextChars` を使う
 * @param date1904 ブックが 1904 年起点か（日付の項目の書式化用）
 */
export function parseChart(
  xml: string,
  theme: Theme,
  limits: XlsxLimits,
  date1904: boolean,
): ChartResult {
  try {
    return { ok: true, chart: parse(xml, theme, limits, date1904) };
  } catch (e) {
    return { ok: false, reason: e instanceof ChartAbort ? e.reason : 'broken' };
  }
}
