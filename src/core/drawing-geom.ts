// 正: InterHub docs/plan/adr/0028-xlsx-preset-shape-geometry.md（§決定 3・7）
/**
 * 図形（`xdr:sp`）の形の読み取り：`a:prstGeom` の調整値（`a:avLst`）と、自由図形（`a:custGeom`）の定義。
 * 自由図形の式・path はファイルの値で非信頼の入力。字句（演算子の表・引数の数・整数か名前）を確かめ、上限（guide・命令の数）で守る。
 * 式の評価と経路は geom の `evaluate.ts`（InterHub ISSUE-90）。ここは評価器がそのまま使える形（`CustGeomDef`）に読むところまで。
 * DOM に依存せず、`findElement`・`forEachElement`・`findOpenTag` と線形時間の正規表現だけで走査する。
 * 名前（guide・調整値）はファイルの文字列。`Map` のキーにだけ使い、DOM・CSS・ログに出さない（console を使わない）。
 */
import { GUIDE_ARITY } from '../geom/defs';
import { ADJ_LIMIT, intAttr } from './drawing-shape';
import type { CustGeomDef, CustGeomPath, GeomArg, GeomPathFill } from './types';
import { NS, findElement, findOpenTag, forEachElement, parseAttrs } from './xml';

/** guide・調整値の名前：英字で始まる英数字 16 文字まで */
const NAME_RE = /^[A-Za-z][A-Za-z0-9]{0,15}$/;
/** 角度の組み込み変数（`3cd4`・`3cd8`・`5cd8`・`7cd8`。数字で始まるので `NAME_RE` に入らない） */
const ANGLE_VAR_RE = /^[357]cd[48]$/;
/** `fmla` の長さの上限（文字数。最長の式でも数十文字）。超えたら字句の誤り */
const MAX_FMLA_LENGTH = 200;

/** 式の演算子ごとの引数の数（定義は geom の `defs.ts`。読み取り側の字句の検査と評価器で同じ表を使う） */
export { GUIDE_ARITY };

/** path 命令の要素名 → `[評価器の命令名, 子の `a:pt` の数]`（`arcTo` は属性 4 つ・`pt` なし） */
const COMMANDS: ReadonlyMap<string, readonly [string, number]> = new Map([
  ['moveTo', ['M', 1]],
  ['lnTo', ['L', 1]],
  ['arcTo', ['A', 0]],
  ['quadBezTo', ['Q', 2]],
  ['cubicBezTo', ['C', 3]],
  ['close', ['Z', 0]],
]);

/** `a:path fill` の値（`none` 以外は塗りの濃淡の種類） */
const PATH_FILLS: ReadonlySet<string> = new Set<GeomPathFill>([
  'none',
  'norm',
  'lighten',
  'lightenLess',
  'darken',
  'darkenLess',
]);

/** 開きタグ 1 つ（名前空間の接頭辞可。`[^<>]*>` で打ち切る）。path の中の命令・`a:pt` を順に拾う */
const TAG_RE = new RegExp(String.raw`<${NS}([A-Za-z_][\w.-]*)(?=[\s>/])[^<>]*>`, 'g');

/** 整数を ±2^31-1 に丸める（調整値と同じ。値を有限に保つためだけ） */
const clampInt = (n: number): number => Math.min(Math.max(n, -ADJ_LIMIT), ADJ_LIMIT);

/** 引数の字句：整数（15 桁まで。±2^31-1 に丸める）か名前（組み込み変数・guide）。誤りなら null */
function parseArg(s: string | undefined): GeomArg | null {
  if (s === undefined) return null;
  const t = s.trim();
  const n = intAttr(t);
  if (n !== null) return clampInt(n);
  return NAME_RE.test(t) || ANGLE_VAR_RE.test(t) ? t : null;
}

/**
 * `a:prstGeom` の中身から、調整値（`a:avLst` の `a:gd name fmla="val N"`）を全名前で読む。
 * `val N` だけ（ほかの式は読まない）。名前は英字で始まる英数字 16 文字まで。値は ±2,147,483,647 に丸める。
 * `max` 個（既定 16）を超えた分は読まない。
 */
export function readPresetAdjustments(
  geomBody: string | null,
  max: number,
): Array<readonly [string, number]> {
  const adj: Array<readonly [string, number]> = [];
  if (geomBody === null) return adj;
  forEachElement(geomBody, 'gd', (open) => {
    if (adj.length >= max) return;
    const a = parseAttrs(open);
    const name = a.name ?? '';
    const v = /^val\s+(-?\d{1,15})$/.exec((a.fmla ?? '').trim());
    if (!NAME_RE.test(name) || v === null) return;
    adj.push([name, clampInt(Number(v[1]))]);
  });
  return adj;
}

/** `readCustGeom` に渡す上限 */
export interface CustGeomBudget {
  /** 1 図形の guide（`avLst`＋`gdLst`）の数の上限 */
  maxGuides: number;
  /** 1 図形の path 命令の数の上限（全 path の合計） */
  maxCommands: number;
  /** 1 シートに残っている「path 命令の数＋guide の数」（上限から、これまでの自由図形の合計を引いた値） */
  sheetRoom: number;
}

/**
 * `readCustGeom` の結果。
 * - ok：形の定義と、シートの合計（`budget.sheetRoom`）に数える数（path 命令の数＋guide の数。命令の無い path も 1 つと数える）
 * - invalid：字句の誤り・構造の誤り（形を持たせない。通知はしない）
 * - limit：1 図形の guide か命令の数が上限を超えた（形を持たせず通知）
 * - sheet-limit：1 シートの命令と guide の合計が上限を超えた（形を持たせず通知。以後の自由図形も同じ）
 */
export type CustGeomRead =
  | { ok: true; def: CustGeomDef; units: number }
  | { ok: false; reason: 'invalid' | 'limit' | 'sheet-limit' };

/** guide 1 つの式（`fmla`）を `[演算子, 引数…]` に読む。演算子が未知・引数の数が違う・引数の字句が誤りなら null */
function parseFmla(fmla: string | undefined): [string, ...GeomArg[]] | null {
  const t = (fmla ?? '').trim();
  if (t.length === 0 || t.length > MAX_FMLA_LENGTH) return null;
  const parts = t.split(/\s+/);
  const op = parts[0]!;
  if (GUIDE_ARITY.get(op) !== parts.length - 1) return null;
  const out: [string, ...GeomArg[]] = [op];
  for (let i = 1; i < parts.length; i++) {
    const a = parseArg(parts[i]);
    if (a === null) return null;
    out.push(a);
  }
  return out;
}

/** 命令（または命令の無い path）を 1 つ数える。1 図形の上限・シートの残りを超えたら `counter.over` を立てて false */
function charge(
  counter: { n: number; over: 'limit' | 'sheet-limit' | null; base: number },
  budget: CustGeomBudget,
): boolean {
  counter.n++;
  if (counter.n > budget.maxCommands) counter.over = 'limit';
  else if (counter.base + counter.n > budget.sheetRoom) counter.over = 'sheet-limit';
  return counter.over === null;
}

/** `a:path 1 つ` の中身（命令と `a:pt` の並び）を命令の列にする。誤りなら null。`counter.n` に命令の数を足し、上限を超えたら `counter.over` を立てて打ち切る */
function parseCommands(
  body: string | null,
  counter: { n: number; over: 'limit' | 'sheet-limit' | null; base: number },
  budget: CustGeomBudget,
): CustGeomPath['c'] | null {
  const cmds: Array<readonly [string, ...GeomArg[]]> = [];
  if (body === null) return cmds;
  /** 今の命令。`need` 個の `a:pt` を集める（`arcTo` は属性から引数を取る） */
  let cur: { op: string; need: number; args: GeomArg[] } | null = null;
  const finish = (): boolean => {
    if (cur === null) return true;
    const ok = cur.args.length === (cur.op === 'A' ? 4 : cur.need * 2);
    if (ok) cmds.push([cur.op, ...cur.args]);
    cur = null;
    return ok;
  };
  TAG_RE.lastIndex = 0;
  for (let m = TAG_RE.exec(body); m !== null; m = TAG_RE.exec(body)) {
    const name = m[1]!;
    if (name === 'pt') {
      if (cur === null || cur.op === 'A' || cur.args.length >= cur.need * 2) return null;
      const a = parseAttrs(m[0]);
      const x = parseArg(a.x);
      const y = parseArg(a.y);
      if (x === null || y === null) return null;
      cur.args.push(x, y);
      continue;
    }
    const def = COMMANDS.get(name);
    if (def === undefined) return null;
    if (!finish()) return null;
    if (!charge(counter, budget)) return null;
    cur = { op: def[0], need: def[1], args: [] };
    if (def[0] === 'A') {
      const a = parseAttrs(m[0]);
      for (const k of ['wR', 'hR', 'stAng', 'swAng']) {
        const v = parseArg(a[k]);
        if (v === null) return null;
        cur.args.push(v);
      }
    }
  }
  return finish() ? cmds : null;
}

/**
 * `a:custGeom` の中身（`avLst`・`gdLst`・`rect`・`pathLst`）を評価器の入力の形（`CustGeomDef`）に読む。
 * 式は演算子の表（17 種）と引数の数、引数は整数か名前の字句で確かめる。path は `w`・`h`・`fill`・`stroke` と 6 命令（点は整数か名前）。
 * 字句の誤り・path が 1 つも無い・`custGeom` の中身が取れないときは `invalid`。上限（`budget`）を超えたら、残りを読まず `limit`・`sheet-limit`。
 * 時間は中身の大きさに比例する（上限を超えた時点で打ち切る）。
 */
export function readCustGeom(body: string | null, budget: CustGeomBudget): CustGeomRead {
  if (body === null) return { ok: false, reason: 'invalid' };
  let guides = 0;
  let bad = false;

  // avLst：`val N` だけ。gdLst：演算子の表と引数の数で確かめる
  const av: Array<readonly [string, number]> = [];
  const avBody = findElement(body, 'avLst')?.body ?? null;
  if (avBody !== null) {
    forEachElement(avBody, 'gd', (open) => {
      if (bad || ++guides > budget.maxGuides) return;
      const a = parseAttrs(open);
      const v = /^val\s+(-?\d{1,15})$/.exec((a.fmla ?? '').trim());
      if (!NAME_RE.test(a.name ?? '') || v === null) {
        bad = true;
        return;
      }
      av.push([a.name!, clampInt(Number(v[1]))]);
    });
  }
  const gd: Array<readonly [string, string, ...GeomArg[]]> = [];
  const gdBody = findElement(body, 'gdLst')?.body ?? null;
  if (gdBody !== null) {
    forEachElement(gdBody, 'gd', (open) => {
      if (bad || ++guides > budget.maxGuides) return;
      const a = parseAttrs(open);
      const f = parseFmla(a.fmla);
      if (!NAME_RE.test(a.name ?? '') || f === null) {
        bad = true;
        return;
      }
      gd.push([a.name!, ...f]);
    });
  }
  if (guides > budget.maxGuides) return { ok: false, reason: 'limit' };
  if (guides > budget.sheetRoom) return { ok: false, reason: 'sheet-limit' };
  if (bad) return { ok: false, reason: 'invalid' };

  // rect：l・t・r・b の 4 つがそろい、すべて整数か名前
  let rect: GeomArg[] | undefined;
  const rectTag = findOpenTag(body, 'rect');
  if (rectTag !== null) {
    const a = parseAttrs(rectTag);
    const r = [parseArg(a.l), parseArg(a.t), parseArg(a.r), parseArg(a.b)];
    if (r.some((v) => v === null)) return { ok: false, reason: 'invalid' };
    rect = r as GeomArg[];
  }

  // pathLst：path ごとに命令を読む。path の数も命令の上限で抑える（空の path の大量の並び）
  const paths: CustGeomPath[] = [];
  const counter: { n: number; over: 'limit' | 'sheet-limit' | null; base: number } = {
    n: 0,
    over: null,
    base: guides,
  };
  const pathLst = findElement(body, 'pathLst')?.body ?? null;
  if (pathLst === null) return { ok: false, reason: 'invalid' };
  forEachElement(pathLst, 'path', (open, pathBody) => {
    if (bad || counter.over !== null) return;
    if (paths.length >= budget.maxCommands) {
      counter.over = 'limit';
      return;
    }
    const c = parseCommands(pathBody, counter, budget);
    if (c === null) {
      if (counter.over === null) bad = true;
      return;
    }
    // 命令の無い path も 1 つと数える（空の path の大量の並びでシートの保持量が膨らまないように）
    if (c.length === 0 && !charge(counter, budget)) return;
    const a = parseAttrs(open);
    const path: CustGeomPath = { c };
    const w = intAttr(a.w);
    const h = intAttr(a.h);
    if (w !== null && w > 0) path.w = clampInt(w);
    if (h !== null && h > 0) path.h = clampInt(h);
    if (a.fill !== undefined && PATH_FILLS.has(a.fill)) path.fill = a.fill as GeomPathFill;
    if (a.stroke === '0' || a.stroke === 'false') path.stroke = false;
    paths.push(path);
  });
  if (counter.over !== null) return { ok: false, reason: counter.over };
  if (bad || paths.length === 0) return { ok: false, reason: 'invalid' };

  const def: CustGeomDef = rect === undefined ? { av, gd, paths } : { av, gd, rect, paths };
  return { ok: true, def, units: guides + counter.n };
}
