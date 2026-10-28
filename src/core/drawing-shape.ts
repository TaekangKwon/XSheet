// 正: InterHub docs/plan/adr/0027-xlsx-shape-text-and-connectors.md（§決定 1・3）
/**
 * 図形（`xdr:sp`）・線（`xdr:cxnSp`）の見た目の読み取り：回転・反転、枠線と塗りの有無、線の形・調整値・太さ・線種・矢じり。
 * 色の解釈は `drawing-color.ts`、形の定義は `drawing-geom.ts`（ここは塗り・参照の切り出しまで）。DOM に依存せず、`findElement`・`forEachElement`・`findOpenTag` だけで線形時間に走査する。
 * 図形の名前・文字はログに出さない（console を使わない）。
 */
import type {
  ArrowSize,
  ArrowType,
  DrawingLine,
  LineDash,
  LineEnd,
  LineGeom,
  ShapeOutline,
} from './types';
import { NS, findElement, findOpenTag, forEachElement, parseAttrs, type Attrs } from './xml';

/** 線として扱う preset（`xdr:sp` の `prstGeom prst` がこれなら種類 `line`）。`xdr:cxnSp` はこれ以外の形でも線 */
export const LINE_PRESETS: ReadonlySet<string> = new Set<LineGeom>([
  'line',
  'straightConnector1',
  'bentConnector2',
  'bentConnector3',
  'bentConnector4',
  'bentConnector5',
  'curvedConnector2',
  'curvedConnector3',
  'curvedConnector4',
  'curvedConnector5',
]);

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

const ARROWS: ReadonlySet<string> = new Set<ArrowType>([
  'triangle',
  'stealth',
  'diamond',
  'oval',
  'arrow',
]);

/** 回転の 1 周（1/60000 度） */
const FULL_TURN = 21_600_000;

/** 調整値・自由図形の整数の引数の上限（±2^31-1。値を有限に保つためだけ。Excel の U 字線は端の揃った形で adj1=1,800,000 を書く。異常値の px 対策は表示側） */
export const ADJ_LIMIT = 2_147_483_647;

/** 線の太さの上限（EMU。ECMA-376 の `ST_LineWidth` の最大値 1,584pt） */
const MAX_LINE_WIDTH = 20_116_800;

/** `lnRef idx` 1・2・3 の太さ（EMU。Office 2013 以降の既定テーマの値。テーマは読まない） */
const LN_REF_WIDTH: Readonly<Record<number, number>> = { 1: 6350, 2: 12700, 3: 19050 };

/** 塗りの要素名（`spPr` 直下・`a:ln` 直下） */
const FILL_RE = new RegExp(
  String.raw`<${NS}(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)(?=[\s>/])`,
);

/** 開き・閉じタグ 1 つ（名前空間の接頭辞可。`[^<>]*>` で打ち切る）。`spPr` 直下の `extLst` を探す深さの計算用 */
const TAG_RE = new RegExp(String.raw`<(/?)${NS}([A-Za-z_][\w.-]*)(?=[\s>/])[^<>]*>`, 'g');

/** 整数の文字（前後の空白可）を数にする。整数でない・桁が多すぎるなら null */
export function intAttr(s: string | undefined): number | null {
  if (s === undefined) return null;
  const t = s.trim();
  return /^-?\d{1,15}$/.test(t) ? Number(t) : null;
}

/** XML の真偽（`1`・`true`）。無ければ false */
export function boolAttr(s: string | undefined): boolean {
  return s === '1' || s === 'true';
}

/** `a:ln` の開きタグと中身（自己閉じ・閉じの無いときの body は null） */
interface LnSpan {
  open: string;
  body: string | null;
}

/** 塗りの要素 1 つ（`a:solidFill`・`a:gradFill` など）。色の読み取り（`drawing-color.ts`）が使う */
export interface FillEl {
  /** 要素名（接頭辞を除く。`noFill`・`solidFill`・`gradFill`・`blipFill`・`pattFill`・`grpFill`） */
  name: string;
  /** 開きタグ全体（`a:pattFill` の `prst` を読む） */
  open: string;
  /** 中身（自己閉じ・閉じの無いときは null） */
  body: string | null;
}

/** `xdr:style` の参照（`a:fillRef`・`a:lnRef`）：`idx` と、子の色の要素を含む中身 */
export interface StyleRef {
  /** `idx`（無い・不正なら null） */
  idx: number | null;
  /** 中身（子の色の要素。自己閉じ・閉じの無いときは null） */
  body: string | null;
}

/** `readShapeProps` の結果 */
export interface ShapeProps {
  /** `a:prstGeom prst`（無ければ null） */
  prst: string | null;
  /** 回転（1/60000 度。0〜21,599,999 に正規化） */
  rot: number;
  flipH: boolean;
  flipV: boolean;
  /** 枠線（線）が見える */
  hasOutline: boolean;
  /** 塗りがある */
  hasFill: boolean;
  /** `a:ln`（無ければ null） */
  ln: LnSpan | null;
  /** `lnRef idx`（`xdr:style` が無い・値が不正なら null） */
  lnRefIdx: number | null;
  /** `spPr` 直下（`a:ln` を除く）で最初に出る塗りの要素（無ければ null。色の読み取り用。有無の規則は `hasFill`） */
  fillEl: FillEl | null;
  /** `a:ln` の中で最初に出る塗りの要素（無ければ null。色の読み取り用） */
  lnFillEl: FillEl | null;
  /** `xdr:style` の `a:fillRef`（無ければ null） */
  fillRef: StyleRef | null;
  /** `xdr:style` の `a:lnRef`（無ければ null） */
  lnRef: StyleRef | null;
  /** `a:prstGeom` の中身（`a:avLst` を含む。自己閉じ・無いときは null） */
  geomBody: string | null;
  /** `a:custGeom`（`a:prstGeom` が無いときだけ探す。無ければ null。自己閉じ・閉じの無いときの body は null） */
  custGeom: { body: string | null } | null;
}

/**
 * `spPr` 直下（深さ 0）で最初に出る `extLst` の位置。無ければ -1。
 * Office 拡張（`a14:hiddenFill` などに本物でない塗り・線が入る）はここより後ろを読まないための位置。
 * `a:ln`・`a:blip` の中の `extLst` では切らない（`a:ln` が途中で切れるため）。深さは開き・閉じタグの数で数え、線形時間。
 */
function topLevelExtLst(xml: string): number {
  let depth = 0;
  TAG_RE.lastIndex = 0;
  for (let m = TAG_RE.exec(xml); m !== null; m = TAG_RE.exec(xml)) {
    if (m[1] === '/') {
      if (depth > 0) depth--;
      continue;
    }
    if (depth === 0 && m[2] === 'extLst') return m.index;
    if (!m[0].endsWith('/>')) depth++;
  }
  return -1;
}

/** 回転を 0〜21,599,999 に正規化する。整数でなければ 0 */
function normalizeRot(s: string | undefined): number {
  const n = intAttr(s);
  if (n === null) return 0;
  return ((n % FULL_TURN) + FULL_TURN) % FULL_TURN;
}

/** 開きタグの `idx`（`lnRef`・`fillRef`）を整数で返す。タグが無い・不正なら null */
function refIdx(spBody: string, name: 'lnRef' | 'fillRef'): number | null {
  const tag = findOpenTag(spBody, name);
  return tag === null ? null : intAttr(parseAttrs(tag).idx);
}

/** 塗りの要素の最初の名前（`noFill` など）。無ければ null */
function firstFill(xml: string): string | null {
  return FILL_RE.exec(xml)?.[1] ?? null;
}

/** 塗りの要素の最初の 1 つを開きタグと中身ごと切り出す。無ければ null */
function firstFillEl(xml: string | null): FillEl | null {
  if (xml === null) return null;
  const m = FILL_RE.exec(xml);
  if (m === null) return null;
  const name = m[1]!;
  const el = findElement(xml, name, m.index);
  return el === null ? null : { name, open: el.open, body: el.body };
}

/** `xdr:style` の参照（`fillRef`・`lnRef`）を `idx` と中身で読む。無ければ null */
function styleRef(spBody: string, name: 'lnRef' | 'fillRef'): StyleRef | null {
  const el = findElement(spBody, name);
  return el === null ? null : { idx: intAttr(parseAttrs(el.open).idx), body: el.body };
}

/**
 * `xdr:sp`・`xdr:cxnSp` の中身から、回転・反転・枠線と塗りの有無・`a:ln`・形を読む。
 * 色の解釈は `drawing-color.ts`（ここは塗りの要素と `xdr:style` の参照の切り出しまで）。
 * - 枠線が無い＝`a:ln` に `a:noFill`、または `a:ln` に線の塗りの指定が無く `lnRef idx="0"`。それ以外は「あり」（`xdr:style` も線の塗りも無ければあり）
 * - 塗りが無い＝`spPr` 直下に `a:noFill`、または塗りの指定が無く `fillRef idx="0"`。それ以外は「あり」
 * 閉じの無い `spPr`・`a:ln` でも例外にしない。
 */
export function readShapeProps(spBody: string): ShapeProps {
  const spPr = findElement(spBody, 'spPr');
  let pr = spPr?.body ?? '';
  const ext = topLevelExtLst(pr);
  if (ext >= 0) pr = pr.slice(0, ext);

  const xfrm = findOpenTag(pr, 'xfrm');
  const xa = xfrm === null ? null : parseAttrs(xfrm);
  const geom = findElement(pr, 'prstGeom');
  const prst = geom === null ? null : (parseAttrs(geom.open).prst ?? null);
  const cust = geom === null ? findElement(pr, 'custGeom') : null;

  // a:ln を切り出し、残りを「spPr 直下」として塗りを判定する（a:ln の中の塗りと取り違えない）
  const lnEl = findElement(pr, 'ln');
  const ln: LnSpan | null = lnEl === null ? null : { open: lnEl.open, body: lnEl.body };
  let direct = pr;
  if (lnEl !== null) {
    const start = pr.indexOf(lnEl.open);
    direct = pr.slice(0, start) + (lnEl.end >= pr.length ? '' : pr.slice(lnEl.end));
  }

  const lnRefIdx = refIdx(spBody, 'lnRef');
  const lnFill = ln?.body == null ? null : firstFill(ln.body);
  const hasOutline = lnFill !== null ? lnFill !== 'noFill' : lnRefIdx !== 0;

  const fill = firstFill(direct);
  const hasFill = fill !== null ? fill !== 'noFill' : refIdx(spBody, 'fillRef') !== 0;

  return {
    prst,
    rot: normalizeRot(xa?.rot),
    flipH: boolAttr(xa?.flipH),
    flipV: boolAttr(xa?.flipV),
    hasOutline,
    hasFill,
    ln,
    lnRefIdx,
    fillEl: firstFillEl(direct),
    lnFillEl: firstFillEl(ln?.body ?? null),
    fillRef: styleRef(spBody, 'fillRef'),
    lnRef: styleRef(spBody, 'lnRef'),
    geomBody: geom?.body ?? null,
    custGeom: cust === null ? null : { body: cust.body },
  };
}

/** `prst` が線として扱う preset か（`xdr:sp` を線にするかの判定） */
export function isLinePreset(prst: string | null): boolean {
  return prst !== null && LINE_PRESETS.has(prst);
}

/** 矢じりの大きさ（`sm`・`med`・`lg`）。無い・未知の値は `med` */
function arrowSize(s: string | undefined): ArrowSize {
  return s === 'sm' || s === 'lg' ? s : 'med';
}

/** 線の端（`a:headEnd`・`a:tailEnd`）。`type` が無い・`none` は矢じりなし（null）。未知の種類は `triangle` */
function lineEnd(lnBody: string, name: 'headEnd' | 'tailEnd'): LineEnd | null {
  const tag = findOpenTag(lnBody, name);
  if (tag === null) return null;
  const a: Attrs = parseAttrs(tag);
  if (a.type === undefined || a.type === 'none') return null;
  const type = (ARROWS.has(a.type) ? a.type : 'triangle') as ArrowType;
  return { type, w: arrowSize(a.w), len: arrowSize(a.len) };
}

/** 調整値 adj1〜adj3 を読む（`a:gd name="adjN" fmla="val N"` だけ。式は読まない。無ければ 50000。範囲外は丸める） */
function readAdj(geomBody: string | null): [number, number, number] {
  const adj: [number, number, number] = [50000, 50000, 50000];
  if (geomBody === null) return adj;
  forEachElement(geomBody, 'gd', (open) => {
    const a = parseAttrs(open);
    const m = /^adj([123])$/.exec(a.name ?? '');
    const v = /^val\s+(-?\d{1,15})$/.exec((a.fmla ?? '').trim());
    if (m === null || v === null) return;
    adj[Number(m[1]) - 1] = Math.min(Math.max(Number(v[1]), -ADJ_LIMIT), ADJ_LIMIT);
  });
  return adj;
}

/**
 * 枠線（線）の太さと線種。線・矢印（`readLine`）と図形の枠線で同じ規則を使う。
 * 太さは `a:ln w`（0〜20,116,800 に丸める）、無ければ `lnRef idx` 1・2・3 → 6350・12700・19050 EMU、どちらも無ければ 0。
 * 線種は `a:prstDash val`（無い・未知は実線）。
 */
export function readStroke(props: ShapeProps): ShapeOutline {
  const lnAttrs = props.ln === null ? null : parseAttrs(props.ln.open);
  const w = intAttr(lnAttrs?.w);
  const widthEmu =
    w !== null
      ? Math.min(Math.max(w, 0), MAX_LINE_WIDTH)
      : (LN_REF_WIDTH[props.lnRefIdx ?? -1] ?? 0);
  const lnBody = props.ln?.body ?? null;
  const dashTag = lnBody === null ? null : findOpenTag(lnBody, 'prstDash');
  const dashVal = dashTag === null ? undefined : parseAttrs(dashTag).val;
  const dash = (dashVal !== undefined && DASHES.has(dashVal) ? dashVal : 'solid') as LineDash;
  return { widthEmu, dash };
}

/**
 * 線・矢印の形と見た目を読む。形は `prst`（線の preset 以外・`custGeom` は直線＝`line`）、
 * 太さ・線種は `readStroke`、矢じりは `a:headEnd`（始点）・`a:tailEnd`（終点）。
 */
export function readLine(props: ShapeProps): DrawingLine {
  const geom = (isLinePreset(props.prst) ? props.prst : 'line') as LineGeom;
  const { widthEmu, dash } = readStroke(props);
  const lnBody = props.ln?.body ?? null;
  return {
    geom,
    adj: readAdj(props.geomBody),
    widthEmu,
    dash,
    head: lnBody === null ? null : lineEnd(lnBody, 'headEnd'),
    tail: lnBody === null ? null : lineEnd(lnBody, 'tailEnd'),
  };
}
