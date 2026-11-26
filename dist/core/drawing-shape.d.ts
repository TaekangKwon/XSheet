/**
 * 図形（`xdr:sp`）・線（`xdr:cxnSp`）の見た目の読み取り：回転・反転、枠線と塗りの有無、線の形・調整値・太さ・線種・矢じり。
 * 色の解釈は `drawing-color.ts`、形の定義は `drawing-geom.ts`（ここは塗り・参照の切り出しまで）。DOM に依存せず、`findElement`・`forEachElement`・`findOpenTag` だけで線形時間に走査する。
 * 図形の名前・文字はログに出さない（console を使わない）。
 */
import type { DrawingLine, ShapeOutline } from './types';
/** 線として扱う preset（`xdr:sp` の `prstGeom prst` がこれなら種類 `line`）。`xdr:cxnSp` はこれ以外の形でも線 */
export declare const LINE_PRESETS: ReadonlySet<string>;
/** 調整値・自由図形の整数の引数の上限（±2^31-1。値を有限に保つためだけ。Excel の U 字線は端の揃った形で adj1=1,800,000 を書く。異常値の px 対策は表示側） */
export declare const ADJ_LIMIT = 2147483647;
/** 整数の文字（前後の空白可）を数にする。整数でない・桁が多すぎるなら null */
export declare function intAttr(s: string | undefined): number | null;
/** XML の真偽（`1`・`true`）。無ければ false */
export declare function boolAttr(s: string | undefined): boolean;
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
    custGeom: {
        body: string | null;
    } | null;
}
/**
 * `xdr:sp`・`xdr:cxnSp` の中身から、回転・反転・枠線と塗りの有無・`a:ln`・形を読む。
 * 色の解釈は `drawing-color.ts`（ここは塗りの要素と `xdr:style` の参照の切り出しまで）。
 * - 枠線が無い＝`a:ln` に `a:noFill`、または `a:ln` に線の塗りの指定が無く `lnRef idx="0"`。それ以外は「あり」（`xdr:style` も線の塗りも無ければあり）
 * - 塗りが無い＝`spPr` 直下に `a:noFill`、または塗りの指定が無く `fillRef idx="0"`。それ以外は「あり」
 * 閉じの無い `spPr`・`a:ln` でも例外にしない。
 */
export declare function readShapeProps(spBody: string): ShapeProps;
/** `prst` が線として扱う preset か（`xdr:sp` を線にするかの判定） */
export declare function isLinePreset(prst: string | null): boolean;
/**
 * 枠線（線）の太さと線種。線・矢印（`readLine`）と図形の枠線で同じ規則を使う。
 * 太さは `a:ln w`（0〜20,116,800 に丸める）、無ければ `lnRef idx` 1・2・3 → 6350・12700・19050 EMU、どちらも無ければ 0。
 * 線種は `a:prstDash val`（無い・未知は実線）。
 */
export declare function readStroke(props: ShapeProps): ShapeOutline;
/**
 * 線・矢印の形と見た目を読む。形は `prst`（線の preset 以外・`custGeom` は直線＝`line`）、
 * 太さ・線種は `readStroke`、矢じりは `a:headEnd`（始点）・`a:tailEnd`（終点）。
 */
export declare function readLine(props: ShapeProps): DrawingLine;
export {};
