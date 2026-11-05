/**
 * `.xlsx` の図形の「形」の定義（DrawingML の guide の式と path 命令）を評価する純ロジック（DOM 非依存・InterHub ISSUE-90）。
 * 流れ：形の定義（表の preset か `custGeom`）＋調整値＋回転前の矩形の幅・高さ（px）
 * → 組み込み変数 → `avLst`（既定を調整値で上書き）→ `gdLst` を上から 1 回ずつ → path 命令を経路（直線・3 次ベジェ）へ。
 * 式は自前の解釈器で評価し、`eval`・`new Function` は使わない（docs/01_requirements.md N1）。
 * 名前（guide・調整値・`prst`）は `Map` のキーにだけ使い、普通のオブジェクトのキーにしない。
 * 反転・回転・シート範囲への丸め・`d` の組み立ては描画側（ISSUE-92・`xlsx-shapes.ts` の写像）が行う。ここは回転前の矩形の中の経路まで。
 * 形の表は生成物 `preset-geom.json`（変換器 `tools/geom/build-preset-geom.mjs`・出典はリポジトリ直下の NOTICE）。
 * コメント中の ISSUE-N・ADR-N は InterHub の番号。
 */
import { MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES, type PathSeg, type Pt } from './defs';
export { MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES };
/** path の塗りの濃淡の種類（`none`＝塗らない） */
export type GeomFill = 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';
/** 式・命令の引数：整数、または組み込み変数・guide の名前 */
export type GeomArg = number | string;
/** guide 1 つ：`[名前, 演算子, 引数…]`（例 `['x1', '+-', 'w', 'adj1', 0]`） */
export type GeomGuide = readonly [name: string, op: string, ...args: GeomArg[]];
/** path 命令 1 つ：`['M', x, y]`・`['L', x, y]`・`['A', wR, hR, stAng, swAng]`・`['Q', x1, y1, x, y]`・`['C', x1, y1, x2, y2, x, y]`・`['Z']` */
export type GeomCmd = readonly [op: string, ...args: GeomArg[]];
/** path 1 つの定義 */
export interface GeomPathDef {
    /** 座標系の幅・高さ（無い・0 なら形の座標系＝EMU のまま） */
    w?: number;
    h?: number;
    /** 塗りの濃淡の種類（無ければ `norm`） */
    fill?: GeomFill;
    /** 線を引くか（無ければ true） */
    stroke?: boolean;
    /** 命令の列 */
    c: readonly GeomCmd[];
}
/** 形の定義（表の preset も `custGeom` もこの形に揃える） */
export interface GeomDef {
    /** 調整値の既定（`avLst`）：`[名前, 値]` */
    av?: ReadonlyArray<readonly [string, number]>;
    /** guide の列（`gdLst`）。名前は「それより前」に定義したものだけ引ける */
    gd?: readonly GeomGuide[];
    /** 文字領域（l・t・r・b の引数。無ければ矩形全体） */
    rect?: readonly GeomArg[];
    /** path の列 */
    paths: readonly GeomPathDef[];
    /**
     * 評価した経路・文字領域を矩形の中で上下反転する（y → h − y）。原本に定義の無い `upArrow` を `downArrow` から作るために使う（`getPresetGeom`）。
     * ファイルの `custGeom` は持たない
     */
    flipY?: boolean;
}
/** 評価の失敗の理由（代替の描き方は ISSUE-92）。`name`＝未知の名前・前方参照、`op`＝未知の演算子、`args`＝引数の数の誤り、`cmd`＝未知の path 命令、`limit`＝上限超え */
export type GeomFailure = 'name' | 'op' | 'args' | 'cmd' | 'limit';
/** 部分経路：始点・区間の列・閉じるか（区間は `xlsx-shapes.ts` の直線 `L`・3 次ベジェ `C` と同じ表し方） */
export interface GeomSubpath {
    start: Pt;
    segs: PathSeg[];
    closed: boolean;
}
/** path 1 つの経路（px・回転前の矩形の左上が原点・y は下向き） */
export interface GeomPath {
    fill: GeomFill;
    stroke: boolean;
    subpaths: GeomSubpath[];
}
/** 文字領域（px・回転前の矩形の左上が原点） */
export interface GeomTextRect {
    l: number;
    t: number;
    r: number;
    b: number;
}
/** 評価の結果：成功なら経路と文字領域、失敗なら理由（例外にはしない） */
export type GeomResult = {
    ok: true;
    paths: GeomPath[];
    textRect: GeomTextRect;
} | {
    ok: false;
    reason: GeomFailure;
};
/** guide の評価の結果（値は EMU） */
export type GuidesResult = {
    ok: true;
    values: Map<string, number>;
} | {
    ok: false;
    reason: GeomFailure;
};
/**
 * guide を評価して名前 → 値（EMU）の表にする。
 * 順序は 組み込み変数 → `av`（`adjust` の同じ名前の値で上書き。`av` に無い名前は無視）→ `gd` を上から 1 回ずつ。
 * 前方参照・未知の名前・未知の演算子・引数の数の誤り・上限超えは例外にせず失敗として返す。
 *
 * @param def 形の定義
 * @param adjust ファイルの調整値（名前 → 値）。無ければ既定のまま
 * @param w 回転前の矩形の幅（px）
 * @param h 回転前の矩形の高さ（px）
 */
export declare function evaluateGuides(def: GeomDef, adjust: ReadonlyMap<string, number> | undefined, w: number, h: number): GuidesResult;
/**
 * 形の定義を評価して、回転前の矩形の中の経路と文字領域を返す（座標は px・矩形の左上が原点）。
 * 失敗（未知の名前・演算子・命令、引数の数の誤り、上限超え）は例外にせず `{ ok: false, reason }` で返す。
 * 経路は直線と 3 次ベジェだけ（`quadBezTo` は誤差なしで変換・`arcTo` は 90° 以下に分けて近似。誤差は半径の 0.03% 未満）。
 * 文字領域は `rect` の l・t・r・b を評価した値。`rect` が無い・幅か高さが 0 以下なら矩形全体。
 * 反転した `rect`（r ≦ l など）は、調整値の極端な値で式が裏返ったものとみなして矩形全体へ戻す
 * （幅 0 以下の領域には文字が並ばず、図形ごと文字が消えるより、図形の矩形に並べる方が読める）。
 *
 * @param def 形の定義（表の preset か `custGeom`）
 * @param adjust ファイルの調整値（名前 → 値）
 * @param w 回転前の矩形の幅（px）
 * @param h 回転前の矩形の高さ（px）
 */
export declare function evaluateGeom(def: GeomDef, adjust: ReadonlyMap<string, number> | undefined, w: number, h: number): GeomResult;
/** preset 名（`prst`）から形の定義を引く。表に無い（線の 10 種・未知の名前）なら undefined。`upArrow` は原本に無いので `downArrow` の上下反転（`flipY`）で返す */
export declare function getPresetGeom(prst: string): GeomDef | undefined;
/** 表にある preset 名の一覧（原本の出現順。別名の `upArrow` は含めない） */
export declare function presetNames(): string[];
