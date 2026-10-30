// src/geom/defs.ts
/**
* geom が持つ共通の定義（演算子の表・上限・経路の型）。geom は他の層に依存しないので、ここに置いて core・render から引く。
* InterHub では評価器（viewers/xlsx-geom.ts）がこれらを core（`lib/xlsx/drawing-geom.ts`・`limits.ts`）と
* 描画側（`viewers/xlsx-shapes.ts`）から import していた（層の逆転）。
* - `GUIDE_ARITY`：core の `drawing-geom.ts`（自由図形の字句の検査）が引く（M1-2）。
* - `MAX_GEOM_GUIDES`・`MAX_GEOM_COMMANDS`：core の `limits.ts` が引いて `DEFAULT_LIMITS` に入れる（M1-2）。
* - `Pt`・`PathSeg`：render の `xlsx-shapes.ts` が引く（M2-1）。
* 形の表（`preset-geom.json`）は読まない（core から引いても表がバンドルに入らないように、評価器と分ける）。
*/
/** 式の演算子ごとの引数の数（DrawingML の guide の 17 種。読み取り側の字句の検査と評価器で同じ表を使う） */
var GUIDE_ARITY = /* @__PURE__ */ new Map([
	["*/", 3],
	["+-", 3],
	["+/", 3],
	["?:", 3],
	["abs", 1],
	["sqrt", 1],
	["val", 1],
	["max", 2],
	["min", 2],
	["mod", 3],
	["pin", 3],
	["sat2", 3],
	["cat2", 3],
	["sin", 2],
	["cos", 2],
	["tan", 2],
	["at2", 2]
]);
/** 1 図形の guide の数の上限（自由図形の `avLst`＋`gdLst`。preset の最大は gear9 の 244。ADR-0028 §決定 7）。超えたら形を描かず代替 */
var MAX_GEOM_GUIDES = 512;
/** 1 図形の path 命令の数の上限（全 path の合計。ADR-0028 §決定 7）。超えたら形を描かず代替 */
var MAX_GEOM_COMMANDS = 4096;
export { MAX_GEOM_COMMANDS as n, MAX_GEOM_GUIDES as r, GUIDE_ARITY as t };
