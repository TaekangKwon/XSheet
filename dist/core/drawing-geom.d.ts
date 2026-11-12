/**
 * 図形（`xdr:sp`）の形の読み取り：`a:prstGeom` の調整値（`a:avLst`）と、自由図形（`a:custGeom`）の定義。
 * 自由図形の式・path はファイルの値で非信頼の入力。字句（演算子の表・引数の数・整数か名前）を確かめ、上限（guide・命令の数）で守る。
 * 式の評価と経路は geom の `evaluate.ts`（InterHub ISSUE-90）。ここは評価器がそのまま使える形（`CustGeomDef`）に読むところまで。
 * DOM に依存せず、`findElement`・`forEachElement`・`findOpenTag` と線形時間の正規表現だけで走査する。
 * 名前（guide・調整値）はファイルの文字列。`Map` のキーにだけ使い、DOM・CSS・ログに出さない（console を使わない）。
 */
import { GUIDE_ARITY } from '../geom/defs';
import type { CustGeomDef } from './types';
/** 式の演算子ごとの引数の数（定義は geom の `defs.ts`。読み取り側の字句の検査と評価器で同じ表を使う） */
export { GUIDE_ARITY };
/**
 * `a:prstGeom` の中身から、調整値（`a:avLst` の `a:gd name fmla="val N"`）を全名前で読む。
 * `val N` だけ（ほかの式は読まない）。名前は英字で始まる英数字 16 文字まで。値は ±2,147,483,647 に丸める。
 * `max` 個（既定 16）を超えた分は読まない。
 */
export declare function readPresetAdjustments(geomBody: string | null, max: number): Array<readonly [string, number]>;
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
export type CustGeomRead = {
    ok: true;
    def: CustGeomDef;
    units: number;
} | {
    ok: false;
    reason: 'invalid' | 'limit' | 'sheet-limit';
};
/**
 * `a:custGeom` の中身（`avLst`・`gdLst`・`rect`・`pathLst`）を評価器の入力の形（`CustGeomDef`）に読む。
 * 式は演算子の表（17 種）と引数の数、引数は整数か名前の字句で確かめる。path は `w`・`h`・`fill`・`stroke` と 6 命令（点は整数か名前）。
 * 字句の誤り・path が 1 つも無い・`custGeom` の中身が取れないときは `invalid`。上限（`budget`）を超えたら、残りを読まず `limit`・`sheet-limit`。
 * 時間は中身の大きさに比例する（上限を超えた時点で打ち切る）。
 */
export declare function readCustGeom(body: string | null, budget: CustGeomBudget): CustGeomRead;
