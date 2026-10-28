// 正: InterHub docs/plan/adr/0027-xlsx-shape-text-and-connectors.md（§決定 5・6）
/**
 * グループ（`xdr:grpSp`）の座標系と、子の `a:xfrm` の読み取り（ISSUE-89）。
 * DOM に依存せず、`findElement`・`findOpenTag` だけで線形時間に走査する。値は EMU の整数（位置は負でもよい）。
 * グループ自身の回転・反転は読まない（ADR-0027 §決定 5）。図の名前・文字はログに出さない（console を使わない）。
 */
import { intAttr } from './drawing-shape';
import type { GroupRect } from './types';
import { NS, findElement, findOpenTag, parseAttrs, type ElementSpan } from './xml';

/** グループの中身の先頭にある `nvGrpSpPr` の開きタグ（先頭の空白可） */
const NV_RE = new RegExp(String.raw`^\s*<${NS}nvGrpSpPr\b[^<>]*>`);
/** `nvGrpSpPr` の直後にある `grpSpPr` の開きタグ（空白可。`lastIndex` を設定して `y` で照合する） */
const PR_RE = new RegExp(String.raw`\s*<${NS}grpSpPr\b[^<>]*>`, 'y');

/**
 * `xfrm` の中身から、`offName`（`off`・`chOff`）と `extName`（`ext`・`chExt`）の 2 つの開きタグで矩形を作る。
 * どちらかが無い・整数でない・大きさが負なら null。
 */
function rectOf(
  body: string,
  offName: 'off' | 'chOff',
  extName: 'ext' | 'chExt',
): GroupRect | null {
  const off = findOpenTag(body, offName);
  const ext = findOpenTag(body, extName);
  if (off === null || ext === null) return null;
  const o = parseAttrs(off);
  const e = parseAttrs(ext);
  const x = intAttr(o.x);
  const y = intAttr(o.y);
  const cx = intAttr(e.cx);
  const cy = intAttr(e.cy);
  if (x === null || y === null || cx === null || cy === null || cx < 0 || cy < 0) return null;
  return { x, y, cx, cy };
}

/**
 * 最初の `a:xfrm`（グラフの `xdr:xfrm` も同じ）の位置と大きさ（`a:off`・`a:ext`）を読む。
 * `xml` は図 1 つ分（`sp`・`cxnSp`・`pic`・`graphicFrame`）か、グループの `grpSpPr` の中身。無い・読めないなら null。
 */
export function readXfrm(xml: string): GroupRect | null {
  const body = findElement(xml, 'xfrm')?.body;
  return body === null || body === undefined ? null : rectOf(body, 'off', 'ext');
}

/** `readGroupHead` の結果 */
export interface GroupHead {
  /** グループ自身の `cNvPr` の開きタグ（無ければ null） */
  cNvPr: string | null;
  /** グループ自身の `grpSpPr`（無い・直後に無いなら null） */
  pr: ElementSpan | null;
}

/**
 * グループ（`xdr:grpSp` の中身）の頭にある、グループ自身の `nvGrpSpPr`（の `cNvPr`）と、その直後の `grpSpPr` を取り出す。
 * 先頭（`nvGrpSpPr`）と直後（`grpSpPr`）だけを見る：壊れたグループ（`nvGrpSpPr`・`grpSpPr` が無い）で、
 * 子（入れ子のグループ）の `cNvPr`・`grpSpPr` を自分のものとして拾わないため。
 */
export function readGroupHead(grpBody: string): GroupHead {
  const open = NV_RE.exec(grpBody);
  if (open === null) return { cNvPr: null, pr: null };
  const nv = findElement(grpBody, 'nvGrpSpPr', 0);
  if (nv === null) return { cNvPr: null, pr: null };
  const cNvPr = nv.body === null ? null : findOpenTag(nv.body, 'cNvPr');
  PR_RE.lastIndex = nv.end;
  const m = PR_RE.exec(grpBody);
  return { cNvPr, pr: m === null ? null : findElement(grpBody, 'grpSpPr', m.index) };
}

/** `readGroupFrame` の結果 */
export interface GroupFrame {
  /** 子が使う座標系（`a:chOff`・`a:chExt`。大きさは 1 以上） */
  ch: GroupRect;
  /** このグループ自身の位置と大きさ（`a:off`・`a:ext`。親のグループの子の座標系）。`needOwn` が false のときは null */
  own: GroupRect | null;
  /** `grpSpPr` の終わりの位置（`grpBody` の中。ここから子を走査する） */
  end: number;
}

/**
 * グループ（`xdr:grpSp` の中身）の座標系を読む。展開できないグループは null：`grpSpPr`・`a:xfrm`・`a:chOff`・`a:chExt` が無い、
 * 数が整数でない、`chExt` の幅か高さが 0 以下（割り算ができない）、`needOwn` なのに `a:off`・`a:ext` が無い。
 * @param pr グループ自身の `grpSpPr`（`readGroupHead` の結果）
 * @param needOwn 入れ子のグループか（親の座標系での位置が要る。最も外側のグループはアンカーの矩形を使うので要らない）
 */
export function readGroupFrame(pr: ElementSpan | null, needOwn: boolean): GroupFrame | null {
  if (pr === null || pr.body === null) return null;
  const xfrm = findElement(pr.body, 'xfrm')?.body;
  if (xfrm === null || xfrm === undefined) return null;
  const ch = rectOf(xfrm, 'chOff', 'chExt');
  if (ch === null || ch.cx <= 0 || ch.cy <= 0) return null;
  const own = needOwn ? rectOf(xfrm, 'off', 'ext') : null;
  if (needOwn && own === null) return null;
  return { ch, own, end: pr.end };
}
