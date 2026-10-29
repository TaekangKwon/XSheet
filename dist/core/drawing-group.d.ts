import type { GroupRect } from './types';
import { type ElementSpan } from './xml';
/**
 * 最初の `a:xfrm`（グラフの `xdr:xfrm` も同じ）の位置と大きさ（`a:off`・`a:ext`）を読む。
 * `xml` は図 1 つ分（`sp`・`cxnSp`・`pic`・`graphicFrame`）か、グループの `grpSpPr` の中身。無い・読めないなら null。
 */
export declare function readXfrm(xml: string): GroupRect | null;
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
export declare function readGroupHead(grpBody: string): GroupHead;
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
export declare function readGroupFrame(pr: ElementSpan | null, needOwn: boolean): GroupFrame | null;
