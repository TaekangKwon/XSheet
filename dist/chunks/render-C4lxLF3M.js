import { l as XlsxError, p as colName, t as applyFillMode } from "./color-sme5jHOH.js";
import { c as isPaperFill, r as createGeometry, t as openXlsx } from "./core-CGDQdCgf.js";
// src/render/messages.ts
/** 日本語（既定） */
var MESSAGES_JA = Object.freeze({
	libLoadFailed: "図形の形の表を読み込めなかったため、図形は簡略化して表示しています。",
	parseFailed: "Excel ファイルの解析に失敗しました: {detail}",
	noSheets: "（シートがありません）",
	noData: "（データがありません）",
	encrypted: "このファイルは暗号化されているか、旧形式（.xls）の可能性があるため表示できません。",
	tooLarge: "Excel ファイルが大きすぎるため表示できません（内部の上限を超えています）。",
	unsupported: "この形式の Excel ファイル（Strict Open XML・.xlsb など）には対応していません。",
	hiddenSheet: "非表示のシート",
	tabsPrev: "前のシート",
	tabsNext: "次のシート",
	chart: "グラフ",
	smartart: "SmartArt",
	shape: "図形",
	textbox: "テキストボックス",
	group: "グループ",
	imageFormat: "画像（{format}）",
	imageBroken: "画像（表示できない）",
	noticeTooLarge: "図のデータが大きすぎるため、図は表示していません（表は表示しています）。",
	noticeBroken: "図を読み取れなかったため、図は表示していません（表は表示しています）。",
	noticeTooMany: "図が多いため、先頭の 1,000 個だけ表示しています。",
	noticeTextTruncated: "図形の文字が多いため、一部を省略して表示しています。",
	noticeShapeSimplified: "図形の形を描ききれないものがあるため、一部の図形は簡略化して表示しています。",
	noticeOverBudget: "画像の合計が大きいため、一部の画像は枠だけで表示しています。"
});
/** 英語 */
var MESSAGES_EN = Object.freeze({
	libLoadFailed: "Failed to load the shape definitions, so shapes are shown in a simplified form.",
	parseFailed: "Failed to parse the Excel file: {detail}",
	noSheets: "(No sheets)",
	noData: "(No data)",
	encrypted: "This file is encrypted, or may be in the old .xls format, so it cannot be displayed.",
	tooLarge: "The Excel file is too large to display (it exceeds an internal limit).",
	unsupported: "This Excel format (Strict Open XML, .xlsb, etc.) is not supported.",
	hiddenSheet: "Hidden sheet",
	tabsPrev: "Previous sheets",
	tabsNext: "Next sheets",
	chart: "Chart",
	smartart: "SmartArt",
	shape: "Shape",
	textbox: "Text box",
	group: "Group",
	imageFormat: "Image ({format})",
	imageBroken: "Image (cannot be displayed)",
	noticeTooLarge: "The drawing data is too large, so drawings are not shown (the table is shown).",
	noticeBroken: "The drawings could not be read, so they are not shown (the table is shown).",
	noticeTooMany: "Too many drawings: only the first 1,000 are shown.",
	noticeTextTruncated: "The shapes contain a lot of text, so some of it is omitted.",
	noticeShapeSimplified: "Some shapes are too complex to draw exactly, so they are shown in a simplified form.",
	noticeOverBudget: "The images are too large in total, so some images are shown as frames only."
});
/** 韓国語（M6-2・ADR-0003 Q4） */
var MESSAGES_KO = Object.freeze({
	libLoadFailed: "도형 정의를 불러오지 못해 도형을 간략하게 표시합니다.",
	parseFailed: "Excel 파일을 해석하지 못했습니다: {detail}",
	noSheets: "(시트가 없습니다)",
	noData: "(데이터가 없습니다)",
	encrypted: "암호화된 파일이거나 이전 형식(.xls)일 수 있어 표시할 수 없습니다.",
	tooLarge: "Excel 파일이 너무 커서 표시할 수 없습니다(내부 상한을 넘었습니다).",
	unsupported: "이 형식의 Excel 파일(Strict Open XML·.xlsb 등)은 지원하지 않습니다.",
	hiddenSheet: "숨겨진 시트",
	tabsPrev: "이전 시트",
	tabsNext: "다음 시트",
	chart: "차트",
	smartart: "SmartArt",
	shape: "도형",
	textbox: "텍스트 상자",
	group: "그룹",
	imageFormat: "이미지({format})",
	imageBroken: "이미지(표시할 수 없음)",
	noticeTooLarge: "그림 데이터가 너무 커서 그림은 표시하지 않았습니다(표는 표시했습니다).",
	noticeBroken: "그림을 읽지 못해 그림은 표시하지 않았습니다(표는 표시했습니다).",
	noticeTooMany: "그림이 많아 처음 1,000개만 표시합니다.",
	noticeTextTruncated: "도형의 글자가 많아 일부를 생략하고 표시합니다.",
	noticeShapeSimplified: "정확히 그릴 수 없는 도형이 있어 일부 도형은 간략하게 표시합니다.",
	noticeOverBudget: "이미지의 합계가 커서 일부 이미지는 틀만 표시합니다."
});
/**
* 使う文言の表を作る：既定（日本語）に、渡された分だけを上書きする。
* 文字列でない値・表に無いキーは無視する（使う側の誤りで表示が壊れないように）。
*/
function resolveMessages(messages, base = MESSAGES_JA) {
	if (!messages) return base;
	const out = { ...base };
	for (const key of Object.keys(base)) {
		if (!Object.prototype.hasOwnProperty.call(messages, key)) continue;
		const v = messages[key];
		if (typeof v === "string") out[key] = v;
	}
	return Object.freeze(out);
}
/** `{name}` を差し込む。値の無い `{name}` はそのまま残す（InterHub の `t()` と同じ） */
function formatMessage(template, params = {}) {
	return template.replace(/\{(\w+)\}/g, (all, name) => Object.prototype.hasOwnProperty.call(params, name) ? String(params[name]) : all);
}
// src/render/border-layer.ts
var SVG_NS$1 = "http://www.w3.org/2000/svg";
/** SVG の path を作り、属性を入れる（値は数値・列挙値・数値だけの `d` のみ） */
function pathEl(attrs) {
	const e = document.createElementNS(SVG_NS$1, "path");
	e.setAttribute("fill", "none");
	for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
	return e;
}
/** 線の path（二重線以外は `d`・幅・パターン、二重線は 1px の線 2 本の `lineD`）。色は墨色なら class、明示色なら CSSOM */
function linePath(p) {
	const double = p.lineD !== void 0;
	const d = double ? p.lineD : p.d;
	if (d === "") return null;
	const attrs = {
		class: p.ink ? "xs-border-path is-ink" : "xs-border-path",
		d,
		"stroke-width": double ? 1 : p.width
	};
	if (!double && p.dasharray !== null) attrs["stroke-dasharray"] = p.dasharray;
	if (!p.diagonal) attrs["shape-rendering"] = "crispEdges";
	const el = pathEl(attrs);
	if (!p.ink && p.color !== null) {
		el.style.setProperty("stroke", p.color);
		if (p.alpha < 1) el.style.setProperty("stroke-opacity", String(p.alpha));
	}
	return el;
}
/** 二重線の隙間の path（セルの面の色で、線 2 本の間とグリッド線を覆う） */
function gapPath(p) {
	const attrs = {
		class: "xs-border-path is-gap",
		d: p.d,
		"stroke-width": p.width
	};
	if (!p.diagonal) attrs["shape-rendering"] = "crispEdges";
	return pathEl(attrs);
}
/** 罫線層を作る（svg は常に 1 枚で、中身だけを `render` で入れ替える） */
function createBorderLayer() {
	const element = document.createElementNS(SVG_NS$1, "svg");
	element.setAttribute("class", "xs-borders");
	element.setAttribute("aria-hidden", "true");
	element.setAttribute("focusable", "false");
	const clear = () => {
		element.replaceChildren();
		element.removeAttribute("viewBox");
		element.setAttribute("width", "0");
		element.setAttribute("height", "0");
	};
	clear();
	return {
		element,
		clear,
		render(result, originX, originY) {
			const b = result.bounds;
			if (!b || result.paths.length === 0) {
				clear();
				return;
			}
			const x = Math.floor(b.left);
			const y = Math.floor(b.top);
			const w = Math.max(0, Math.ceil(b.right) - x);
			const h = Math.max(0, Math.ceil(b.bottom) - y);
			element.style.left = `${originX + x}px`;
			element.style.top = `${originY + y}px`;
			element.setAttribute("width", String(w));
			element.setAttribute("height", String(h));
			element.setAttribute("viewBox", `${x} ${y} ${w} ${h}`);
			const children = [];
			for (const p of result.paths) if (p.lineD !== void 0 && p.d !== "") children.push(gapPath(p));
			for (const p of result.paths) {
				const el = linePath(p);
				if (el) children.push(el);
			}
			element.replaceChildren(...children);
		}
	};
}
// src/render/borders.ts
/** 1 回の描画で作る区間（線 1 本・斜線 1 本）の数の上限。超えた分は描かない（ADR-0029 §決定 7） */
var MAX_BORDER_SEGMENTS = 5e4;
/** 結合セルの範囲を辺の判定用に印す作業量の上限（重なった巨大な結合を並べた細工でも固まらないため） */
var MERGE_MARK_BUDGET = 2e6;
/** 線種の見た目の分類（同じ太さの順位の中の優先：double ＞ 実線 ＞ 破線系 ＞ dotted ＞ hair。ADR-0029 §決定 4） */
var KIND_HAIR = 0;
var KIND_DOTTED = 1;
var KIND_DASHED = 2;
var KIND_SOLID = 3;
/**
* 13 線種の幅・パターン・優先の表（ADR-0029 §決定 3。暫定値で、E2 の照合で直す定数 1 か所）。
* パターンの数値は px。`slantDashDot` の斜めの切り口は再現しない。
*/
var BORDER_STYLES = {
	thin: {
		width: 1,
		dash: null,
		weight: 2,
		kind: KIND_SOLID
	},
	hair: {
		width: 1,
		dash: [1, 1],
		weight: 1,
		kind: KIND_HAIR
	},
	dotted: {
		width: 1,
		dash: [1, 2],
		weight: 2,
		kind: KIND_DOTTED
	},
	dashed: {
		width: 1,
		dash: [3, 1],
		weight: 2,
		kind: KIND_DASHED
	},
	dashDot: {
		width: 1,
		dash: [
			8,
			3,
			3,
			3
		],
		weight: 2,
		kind: KIND_DASHED
	},
	dashDotDot: {
		width: 1,
		dash: [
			8,
			3,
			3,
			3,
			3,
			3
		],
		weight: 2,
		kind: KIND_DASHED
	},
	medium: {
		width: 2,
		dash: null,
		weight: 3,
		kind: KIND_SOLID
	},
	mediumDashed: {
		width: 2,
		dash: [9, 3],
		weight: 3,
		kind: KIND_DASHED
	},
	mediumDashDot: {
		width: 2,
		dash: [
			9,
			3,
			3,
			3
		],
		weight: 3,
		kind: KIND_DASHED
	},
	mediumDashDotDot: {
		width: 2,
		dash: [
			9,
			3,
			3,
			3,
			3,
			3
		],
		weight: 3,
		kind: KIND_DASHED
	},
	slantDashDot: {
		width: 2,
		dash: [
			11,
			1,
			5,
			1
		],
		weight: 3,
		kind: KIND_DASHED
	},
	thick: {
		width: 3,
		dash: null,
		weight: 4,
		kind: KIND_SOLID
	},
	double: {
		width: 3,
		dash: null,
		weight: 4,
		kind: 4,
		double: true
	}
};
/** 線種の表の 1 件（未知の線種・Object.prototype の名前は undefined） */
function specOf(style) {
	return Object.prototype.hasOwnProperty.call(BORDER_STYLES, style) ? BORDER_STYLES[style] : void 0;
}
/**
* 同じ辺の 2 つの指定から、描く 1 つを選ぶ（ADR-0029 §決定 4）。
* 片方だけならそれ。両方あれば、太さの順位が大きい方 → 同じなら線種の分類（double ＞ 実線 ＞ 破線系 ＞ dotted ＞ hair）
* → それでも同じなら左（上）のセルの指定 `a`。線種が表に無い指定は無いものとして扱う。
*
* @param a 左（上）のセルの、この辺（右・下）の指定
* @param b 右（下）のセルの、この辺（左・上）の指定
*/
function pickBorderEdge(a, b) {
	const sa = a ? specOf(a.style) : void 0;
	const sb = b ? specOf(b.style) : void 0;
	if (!sa) return sb ? b : void 0;
	if (!sb) return a;
	if (sa.weight !== sb.weight) return sa.weight > sb.weight ? a : b;
	if (sa.kind !== sb.kind) return sa.kind > sb.kind ? a : b;
	return a;
}
/** 数値を小数 3 桁に丸める（path の文字列を短く・安定させる） */
var num$1 = (v) => Math.round(v * 1e3) / 1e3;
/**
* 辺の色が墨色（自動・黒）か、明示色（`#rrggbb`＋不透明度）かを決める。
* null（自動）・不正な値・不透明な `#000000` は墨色。文字列は整数から組み立てる（ファイルの文字列は通らない）。
*/
function colorOf(edge) {
	const c = edge.color;
	if (!c || !Number.isInteger(c.rgb) || c.rgb < 0 || c.rgb > 16777215 || !Number.isFinite(c.alpha)) return {
		ink: true,
		color: null,
		alpha: 1
	};
	const alpha = Math.round(Math.min(1, Math.max(0, c.alpha)) * 1e3) / 1e3;
	if (c.rgb === 0 && alpha >= 1) return {
		ink: true,
		color: null,
		alpha: 1
	};
	return {
		ink: false,
		color: `#${c.rgb.toString(16).padStart(6, "0")}`,
		alpha
	};
}
/** 隣り合う線が「同じ線」か（線種と色が同じ）を参照の比較で済ませるための、線の解決器 */
function createLineResolver() {
	const byEdge = /* @__PURE__ */ new Map();
	const byKey = /* @__PURE__ */ new Map();
	return (edge) => {
		if (!edge) return void 0;
		const cached = byEdge.get(edge);
		if (cached !== void 0) return cached ?? void 0;
		const spec = specOf(edge.style);
		let line = null;
		if (spec) {
			const { ink, color, alpha } = colorOf(edge);
			const key = `${edge.style}|${ink ? "ink" : `${color}@${alpha}`}`;
			line = byKey.get(key) ?? null;
			if (!line) {
				line = {
					key,
					style: edge.style,
					spec,
					ink,
					color,
					alpha
				};
				byKey.set(key, line);
			}
		}
		byEdge.set(edge, line);
		return line ?? void 0;
	};
}
/**
* 線 1 本の「グリッド線の画素の位置 g」から、stroke の中心の座標を出す。
* 1px はその画素の中心（g + 0.5）、2px はその画素と手前（左・上）の 1 画素（中心は g）、3px はその画素を中心に 1 画素ずつ（g + 0.5）。
* 手前の端が 0 より前に出る時（A 列の左・1 行目の上）は、線の幅だけ内側へ寄せて、端を 0 にそろえる（sticky の見出しの下に隠れない）。
*/
function centerOf(g, width) {
	const c = width === 2 ? g : g + .5;
	return c - width / 2 < 0 ? width / 2 : c;
}
/**
* 描いた範囲の罫線を、線種×色ごとの path にする。
*
* - 範囲の左端の縦の辺は左隣の列、右端の縦の辺は右隣の列のセル（範囲外でも `rows` は全列を持つ）の指定も見る。
*   上端の横の辺は下のセルの上の辺だけ、下端の横の辺は上のセルの下の辺だけで決める。
* - 結合セルは `GridRow` の左上のセルの `colSpan`・`rowSpan`（行の畳みで縮んだ見える行の数）から範囲を復元し、
*   内側の辺は描かない。外周は面する構成セルの辺から取る。斜線は左上のセルの指定で結合範囲の対角に引く。
* - 幅 0 の列・行の辺は描かない（`rows` は見える行だけ）。非表示の列・行をはさむ辺は、前後の見える列・行の指定で決める。
* - 同じ境目の上で同じ線が続く区間は 1 本（`M x1 y H x2`・`M x y1 V y2`）にまとめる。区間の数が `maxSegments` を超えたら
*   以後は描かない（`truncated`）。
*
* @param rows 描く行（見える行だけ・行番号の昇順。全列を持つ）
* @param table 罫線の表（`XlsxSheet.borders`）
* @param geo 寸法（`colLeft`・`colWidth`・`rowTop`）
* @param range 描く列の範囲（1 始まり・両端を含む）
* @param maxSegments 区間の数の上限（既定 {@link MAX_BORDER_SEGMENTS}）
*/
function buildBorderPaths(rows, table, geo, range, maxSegments = MAX_BORDER_SEGMENTS) {
	const out = {
		paths: [],
		bounds: null,
		segments: 0,
		truncated: false
	};
	const { first, last } = range;
	const R = rows.length;
	if (R === 0 || table.length === 0 || last < first || first < 1) return out;
	if (!rows.some((row) => row.borders !== void 0)) return out;
	const resolve = createLineResolver();
	const acc = /* @__PURE__ */ new Map();
	let left = Infinity;
	let top = Infinity;
	let right = -Infinity;
	let bottom = -Infinity;
	const cellBorder = (i, c) => {
		if (c < 1) return void 0;
		const idx = rows[i].borders?.[c - 1];
		return idx === void 0 ? void 0 : table[idx];
	};
	/** 区間 1 つを積む。上限なら false（以後を描かない） */
	const room = () => {
		if (out.segments >= maxSegments) {
			out.truncated = true;
			return false;
		}
		out.segments++;
		return true;
	};
	const bucket = (line, diagonal) => {
		const key = `${diagonal ? "d" : "o"}|${line.key}`;
		let b = acc.get(key);
		if (!b) {
			b = {
				line,
				diagonal,
				d: [],
				lineD: []
			};
			acc.set(key, b);
		}
		return b;
	};
	/** 外接矩形を広げる（padX・padY は線の幅の半分。線の向きに沿う側には足さない） */
	const grow = (x0, y0, x1, y1, padX, padY) => {
		left = Math.min(left, x0 - padX, x1 - padX);
		right = Math.max(right, x0 + padX, x1 + padX);
		top = Math.min(top, y0 - padY, y1 - padY);
		bottom = Math.max(bottom, y0 + padY, y1 + padY);
	};
	/** 横線：y の中心を g（グリッド線の画素）から出し、x0〜x1 を 1 区間にする */
	const emitH = (line, g, x0, x1) => {
		if (!room()) return;
		const w = line.spec.width;
		const cy = centerOf(g, w);
		const b = bucket(line, false);
		b.d.push(`M${num$1(x0)} ${num$1(cy)}H${num$1(x1)}`);
		if (line.spec.double) b.lineD.push(`M${num$1(x0)} ${num$1(cy - 1)}H${num$1(x1)}M${num$1(x0)} ${num$1(cy + 1)}H${num$1(x1)}`);
		grow(x0, cy, x1, cy, 0, w / 2);
	};
	/** 縦線：x の中心を g から出し、y0〜y1 を 1 区間にする */
	const emitV = (line, g, y0, y1) => {
		if (!room()) return;
		const w = line.spec.width;
		const cx = centerOf(g, w);
		const b = bucket(line, false);
		b.d.push(`M${num$1(cx)} ${num$1(y0)}V${num$1(y1)}`);
		if (line.spec.double) b.lineD.push(`M${num$1(cx - 1)} ${num$1(y0)}V${num$1(y1)}M${num$1(cx + 1)} ${num$1(y0)}V${num$1(y1)}`);
		grow(cx, y0, cx, y1, w / 2, 0);
	};
	/** 斜線：(ax, ay) → (bx, by)。二重線は法線方向に 1px ずつ離した 2 本 */
	const emitDiag = (line, ax, ay, bx, by) => {
		if (!room()) return;
		const b = bucket(line, true);
		b.d.push(`M${num$1(ax)} ${num$1(ay)}L${num$1(bx)} ${num$1(by)}`);
		if (line.spec.double) {
			const len = Math.hypot(bx - ax, by - ay) || 1;
			const nx = -(by - ay) / len;
			const ny = (bx - ax) / len;
			b.lineD.push(`M${num$1(ax + nx)} ${num$1(ay + ny)}L${num$1(bx + nx)} ${num$1(by + ny)}M${num$1(ax - nx)} ${num$1(ay - ny)}L${num$1(bx - nx)} ${num$1(by - ny)}`);
		}
		grow(ax, ay, bx, by, line.spec.width / 2, line.spec.width / 2);
	};
	const W = last - first + 3;
	const col0 = first - 1;
	const ids = new Uint32Array(R * W);
	const merges = [];
	let budget = MERGE_MARK_BUDGET;
	for (let i = 0; i < R; i++) {
		const cells = rows[i].cells;
		const upto = Math.min(cells.length, last);
		for (let ci = 0; ci < upto; ci++) {
			const cell = cells[ci];
			if (!cell || cell.colSpan <= 1 && cell.rowSpan <= 1) continue;
			const c1 = ci + 1;
			const c2 = c1 + Math.max(1, cell.colSpan) - 1;
			if (c2 < col0) continue;
			const i2 = Math.min(R - 1, i + Math.max(1, cell.rowSpan) - 1);
			const m = {
				i1: i,
				i2,
				c1,
				c2,
				id: merges.length + 1
			};
			merges.push(m);
			const a = Math.max(c1, col0) - col0;
			const z = Math.min(c2, last + 1) - col0;
			for (let k = i; k <= i2 && budget > 0; k++) for (let x = a; x <= z && budget > 0; x++, budget--) ids[k * W + x] = m.id;
		}
	}
	const idAt = (i, c) => {
		const x = c - col0;
		return x < 0 || x >= W ? 0 : ids[i * W + x];
	};
	const vis = [];
	for (let c = first; c <= last; c++) if (geo.colWidth(c) > 0) vis.push(c);
	let widest = 0;
	for (const row of rows) widest = Math.max(widest, row.cells.length, row.borders?.length ?? 0);
	let prevVis = 0;
	for (let c = first - 1; c >= 1; c--) if (geo.colWidth(c) > 0) {
		prevVis = c;
		break;
	}
	let nextVis = 0;
	for (let c = last + 1; c <= widest; c++) if (geo.colWidth(c) > 0) {
		nextVis = c;
		break;
	}
	/** 連続する番号 0..n-1 に、線 `at(k)` を割り当て、同じ線の続く区間ごとに `emit(start, end, line)` を呼ぶ */
	const runs = (n, at, emit) => {
		let cur;
		let start = 0;
		for (let k = 0; k < n; k++) {
			const line = at(k);
			if (line !== cur) {
				if (cur) emit(start, k - 1, cur);
				cur = line;
				start = k;
				if (out.truncated) return;
			}
		}
		if (cur && !out.truncated) emit(start, n - 1, cur);
	};
	for (let t = 0; t <= vis.length && !out.truncated; t++) {
		const lc = t === 0 ? prevVis : vis[t - 1];
		const rc = t < vis.length ? vis[t] : nextVis;
		if (vis.length === 0) break;
		const gx = (t < vis.length ? geo.colLeft(rc) : geo.colLeft(lc + 1)) - 1;
		runs(R, (i) => {
			if (lc > 0 && rc > 0) {
				const id = idAt(i, lc);
				if (id !== 0 && id === idAt(i, rc)) return void 0;
			}
			const a = lc > 0 ? cellBorder(i, lc)?.right : void 0;
			const b = rc > 0 ? cellBorder(i, rc)?.left : void 0;
			return resolve(pickBorderEdge(a, b));
		}, (s, e, line) => emitV(line, gx, geo.rowTop(rows[s].r), geo.rowTop(rows[e].r + 1)));
	}
	for (let j = 0; j <= R && !out.truncated && vis.length > 0; j++) {
		const gy = (j < R ? geo.rowTop(rows[j].r) : geo.rowTop(rows[R - 1].r + 1)) - 1;
		runs(vis.length, (t) => {
			const c = vis[t];
			if (j > 0 && j < R) {
				const id = idAt(j - 1, c);
				if (id !== 0 && id === idAt(j, c)) return void 0;
			}
			const a = j > 0 ? cellBorder(j - 1, c)?.bottom : void 0;
			const b = j < R ? cellBorder(j, c)?.top : void 0;
			return resolve(pickBorderEdge(a, b));
		}, (s, e, line) => emitH(line, gy, geo.colLeft(vis[s]), geo.colLeft(vis[e] + 1)));
	}
	const diag = (cb, i1, i2, c1, c2) => {
		if (!cb?.diagonal || !cb.diagonalUp && !cb.diagonalDown) return;
		const line = resolve(cb.diagonal);
		if (!line) return;
		const x0 = geo.colLeft(c1);
		const x1 = geo.colLeft(c2 + 1);
		const y0 = geo.rowTop(rows[i1].r);
		const y1 = geo.rowTop(rows[i2].r + 1);
		if (x1 <= x0 || y1 <= y0) return;
		if (cb.diagonalDown) emitDiag(line, x0, y0, x1, y1);
		if (cb.diagonalUp) emitDiag(line, x0, y1, x1, y0);
	};
	for (const m of merges) {
		if (out.truncated) break;
		if (m.c1 > last || m.c2 < first) continue;
		diag(cellBorder(m.i1, m.c1), m.i1, m.i2, m.c1, m.c2);
	}
	for (let i = 0; i < R && !out.truncated; i++) {
		const row = rows[i];
		if (row.borders === void 0) continue;
		for (const c of vis) {
			const cell = row.cells[c - 1];
			if (cell === null || cell && (cell.colSpan > 1 || cell.rowSpan > 1)) continue;
			diag(cellBorder(i, c), i, i, c, c);
		}
	}
	for (const b of acc.values()) {
		const { line } = b;
		const path = {
			style: line.style,
			ink: line.ink,
			color: line.color,
			alpha: line.alpha,
			width: line.spec.width,
			dasharray: line.spec.dash ? line.spec.dash.join(" ") : null,
			diagonal: b.diagonal,
			d: b.d.join("")
		};
		if (line.spec.double) path.lineD = b.lineD.join("");
		out.paths.push(path);
	}
	if (out.paths.length > 0) out.bounds = {
		left,
		top,
		right,
		bottom
	};
	return out;
}
// src/render/format-css.ts
/**
* `.xlsx` のセルの書式（塗り・文字の装飾）の CSS を作る純関数（ISSUE-100・DOM 非依存）。
* - 書式ごとに規則 1 行：`.xs-sheet[data-xs-fmt="<連番>"] td.xs-xf-<番号> { … }`。
*   連番の属性で絞るため、既存の `.xs-grid td`（グリッド線）より詳細度が強く、同じ文書の別の表示と混ざらない。
* - 文字列は固定の語・整数・整数から組んだ `#rrggbb`・小数（文字サイズの比）・固定のテーマ変数の参照
*   （白の塗りの `var(--xs-cell-bg)` だけ）で組む。ファイルの文字列は入れない。
* - ダーク表示の描き分けは規則に焼き込む（テーマ変数は増やさず、切替で描き直さない）：
*   白以外の塗り＝Excel の色の面と黒の文字（明示色はそのまま）、塗りなし・白の塗り（紙）＝画面の面と文字色（明示色だけ書く）。
*/
/** 白の塗り（紙）の面の色。`style.css` の `.xs-body` が定義する変数（ライト・ダークで切り替わる） */
var PAPER_BG = "var(--xs-cell-bg)";
/** 文字サイズの比として描く範囲（`cell-format.ts` の丸めと同じ 0.5〜4） */
var MIN_RATIO = .5;
var MAX_RATIO = 4;
/**
* 書式の番号 → td に足すクラス名（`xs-xf-<番号>`）。
* @param index `XlsxSheet.formats` の位置
*/
function formatClassName(index) {
	return `xs-xf-${index}`;
}
/**
* 「紙でない」塗りがあるか。はみ出す文字の下の面（グリッド線隠し）を敷くかどうかの判定に使う（ADR-0030 §決定 9）。
* 白の塗り・塗りなし・書式なし（undefined）は false。
*/
function hasOverFill(format) {
	return format?.fill !== void 0 && !isPaperFill(format.fill);
}
/** 塗りなし・紙の上の「自動」の文字色＝画面の文字色（`style.css` の `--xs-fg`。ライト・ダークで切り替わる） */
var AUTO_INK = "var(--xs-fg)";
/** 白以外の塗りの上の「自動」の文字色（Excel の自動＝黒。F6） */
var FILL_INK = "#000000";
/**
* はみ出した「自動」の文字色を、列ごとの区間で塗り分けるグラデーション（`linear-gradient`）を作る（ADR-0030 §決定 9）。
* 塗りの上の自動色は黒、塗りなし・紙の上は画面の文字色で、はみ出した文字が塗りの状態の違う列をまたぐと、
* 単色ではダーク表示でどちらかの面の上に読めない色が残る（黄塗りの黒文字が暗い面へ・明るい文字が薄青の面へ）ため、列の面に合わせる。
* - 呼び出し側は `background-image` にこれを入れ、`background-clip: text; color: transparent` を足す（span の左端＝元の左端）。
* - 元の文字が明示色（`CellFormat.color`。黒は自動に寄せ済み）なら null（Excel と同じく単色）。
* - 区間は元の列から右へ。隣り合う同じ色は 1 区間にまとめ、幅 0（非表示の列）は数えない。区間が 1 つなら null（塗り分け不要）。
* - 文字列は固定の語・整数の px・`#000000`・`var(--xs-fg)` だけ（ファイルの文字列は入らない）。
*
* @param table 書式の表（`XlsxSheet.formats`）
* @param formats 行の書式の番号（`GridRow.formats`。描く列に揃えたもの。無ければ全列「塗りなし」）
* @param c はみ出す元の列（0 始まり。`formats` の添字）
* @param widths 元の列から右へ数えた列幅（px）。`spillColumns` の列数ぶん。不正な値は 0 とみなす
* @returns グラデーションの文字列。塗り分けが要らなければ null
*/
function spillInkGradient(table, formats, c, widths) {
	const at = (i) => {
		const f = formats?.[c + i];
		return f === void 0 ? void 0 : table[f];
	};
	if (at(0)?.color !== void 0) return null;
	const segs = [];
	let x = 0;
	for (let i = 0; i < widths.length; i++) {
		const w = widths[i];
		if (!(Number.isFinite(w) && w > 0)) continue;
		x += w;
		const ink = hasOverFill(at(i)) ? FILL_INK : AUTO_INK;
		const last = segs[segs.length - 1];
		if (last && last.ink === ink) last.end = x;
		else segs.push({
			ink,
			end: x
		});
	}
	if (segs.length < 2) return null;
	let start = 0;
	return `linear-gradient(to right, ${segs.map((s) => {
		const from = start;
		start = Math.round(s.end);
		return `${s.ink} ${from}px ${start}px`;
	}).join(", ")})`;
}
/** 色の整数 → `#rrggbb`（整数から組むため、ファイルの文字列は入らない） */
function hex(rgb) {
	return `#${(rgb & 16777215).toString(16).padStart(6, "0")}`;
}
/** 書式 1 件の宣言（`プロパティ: 値` の並び）。描くものが無ければ空 */
function declarations(f) {
	const out = [];
	const fill = f.fill;
	if (fill !== void 0) {
		const bg = isPaperFill(fill) ? PAPER_BG : hex(fill.rgb);
		out.push(`background-color: ${bg}`, `border-right-color: ${bg}`, `border-bottom-color: ${bg}`);
	}
	if (f.color !== void 0) out.push(`color: ${hex(f.color.rgb)}`);
	else if (fill !== void 0 && !isPaperFill(fill)) out.push("color: #000000");
	if (f.bold) out.push("font-weight: 700");
	if (f.italic) out.push("font-style: italic");
	const lines = [];
	if (f.underline === "single" || f.underline === "double") lines.push("underline");
	if (f.strike) lines.push("line-through");
	if (lines.length > 0) {
		out.push(`text-decoration-line: ${lines.join(" ")}`);
		if (f.underline === "double") out.push("text-decoration-style: double");
	}
	const ratio = f.sizeRatio;
	if (ratio !== void 0 && Number.isFinite(ratio) && ratio >= MIN_RATIO && ratio <= MAX_RATIO) out.push(`font-size: calc(${Number(ratio.toFixed(2))} * 1em)`);
	return out;
}
/**
* 書式の表から CSS の文字列を作る（ブックを開いた時に 1 回）。
* 書式の番号は `formats` の位置のまま（宣言が 1 つも出ない書式は規則を作らないが、後ろの番号はずれない）。
*
* @param formats `XlsxSheet.formats`（ブックで共有する書式の表）
* @param seq 表示ごとの連番（`.xs-sheet` の `data-xs-fmt`）。整数でなければ 0
* @returns 規則を改行でつないだ文字列。書式が無ければ空の文字列
*/
function buildFormatCss(formats, seq) {
	const prefix = `.xs-sheet[data-xs-fmt="${Number.isInteger(seq) ? seq : 0}"] td.`;
	const rules = [];
	formats.forEach((f, i) => {
		const d = declarations(f);
		if (d.length > 0) rules.push(`${prefix}${formatClassName(i)} { ${d.join("; ")} }`);
	});
	return rules.join("\n");
}
// src/render/grid.ts
/**
* `.xlsx` グリッド（`spreadsheet.ts`）の純ロジック（DOM 非依存・ISSUE-84）。
* 表を Excel の寸法（`lib/xlsx/geometry.ts`）で固定して描くための計算だけを切り出す。
* - 描く行の範囲・仮想化の上下の余白・セルの内側の高さは、すべて geometry から出す（DOM の実測はしない）。
* - 非表示の行（高さ 0）は DOM に出さない。描く範囲と前後の余分は「見える行の数」で数え、結合の rowSpan は
*   描く見える行の数へ縮める（ADR-0026 §決定 5 の予備案。列の畳みは `<col>` の `visibility: collapse` のまま）。
* - 表の範囲は、データの行・列より広げて、スクロール枠を埋める（Excel のシートのように空セルが続く）。
* - 文字列のはみ出し（右隣が空なら、その列幅の合計まで文字を出す）は、使える列数と幅の純関数にする。
*/
/** 下罫線の太さ（px）。セルの高さ・幅に含める（`style.css` の `border-width` と同じ値） */
var BORDER = 1;
/** Excel の最大行数・最大列数（表を画面いっぱいに広げるときの上限） */
var MAX_GRID_ROWS = 1048576;
var MAX_GRID_COLS = 16384;
/** 文字がはみ出して使える列数の上限（描く行の計算量を抑える。HTML の `colspan` 上限と同じ値） */
var MAX_SPILL_COLS = 1e3;
/**
* 行番号の列の幅（px）を行数の桁数から出す。1 桁 8px＋左右の余白 16px、最小 32px。
* 列幅は固定レイアウトなので、最大の行番号（`rowCount`）が収まる幅を先に決める。
*
* @param rowCount シートの総行数
*/
function rowNumWidth(rowCount) {
	return Math.max(32, String(Number.isFinite(rowCount) && rowCount >= 1 ? Math.trunc(rowCount) : 1).length * 8 + 16);
}
/**
* 行 r の次の見える行（高さ > 0）。無ければ 0。非表示の行が続いても `rowAt`（二分探索）で跳ぶ。
* r の下端の y を含む行は、同じ y に並ぶ非表示の行より後ろの、次の見える行になる（geometry の `rowAt` の仕様）。
*
* @param total 行数の上限
*/
function nextVisibleRow(geo, r, total) {
	const n = geo.rowAt(geo.rowTop(r) + geo.rowHeight(r));
	return n > r && n <= total && geo.rowHeight(n) > 0 ? n : 0;
}
/** 行 r の前の見える行（高さ > 0）。無ければ 0 */
function prevVisibleRow(geo, r) {
	const top = geo.rowTop(r);
	if (!(top > 0)) return 0;
	const p = geo.rowAt(top - 1);
	return p < r && geo.rowHeight(p) > 0 ? p : 0;
}
/**
* スクロール位置から、見えている行と、描く行（前後に「見える行」で overscan 行ぶん足した範囲）を出す。
* 行の高さは可変（Excel の行高・非表示）なので、`rowAt`（geometry の二分探索）で引く。
* 非表示の行は数えない（フィルターで大半が畳まれていても、描く見える行は一定数）。
*
* @param geo `createGeometry` の結果
* @param scrollTop スクロール枠の `scrollTop`（表の最初の行の上端を 0 とみなす。見出し行は sticky で重なるため無視してよい）
* @param clientHeight スクロール枠の表示高さ（px）
* @param total 行数
* @param overscan 見えている範囲の前後に足す見える行の数
*/
function visibleRowRange(geo, scrollTop, clientHeight, total, overscan) {
	if (total < 1) return {
		top: 0,
		bottom: 0,
		first: 0,
		last: 0
	};
	/** 末尾にだけ非表示の行が続くとき、rowAt は非表示の行を返す。見える行へ戻す */
	const visibleAt = (y) => {
		const r = geo.rowAt(y);
		return geo.rowHeight(r) > 0 ? r : prevVisibleRow(geo, r);
	};
	const top = visibleAt(scrollTop);
	if (top === 0) return {
		top: 0,
		bottom: 0,
		first: 0,
		last: 0
	};
	const bottom = Math.max(top, visibleAt(scrollTop + clientHeight));
	let first = top;
	for (let i = 0; i < overscan; i++) {
		const p = prevVisibleRow(geo, first);
		if (p === 0) break;
		first = p;
	}
	let last = bottom;
	for (let i = 0; i < overscan; i++) {
		const n = nextVisibleRow(geo, last, total);
		if (n === 0) break;
		last = n;
	}
	return {
		top,
		bottom,
		first,
		last
	};
}
/**
* 仮想化で描かない上下の行ぶんの高さ（px）。上＝`rowTop(first)`、下＝`totalHeight − rowTop(last + 1)`。
* 上の余白＋描く行の高さ＋下の余白が、いつも `totalHeight` に一致する。
*/
function spacerHeights(geo, first, last) {
	return {
		top: Math.max(0, geo.rowTop(first)),
		bottom: Math.max(0, geo.totalHeight - geo.rowTop(last + 1))
	};
}
/** 枠の幅が分からない（0＝まだレイアウト前）とき、描く列の範囲に使う代わりの幅（px） */
var COL_FALLBACK_VIEW_WIDTH = 1280;
/**
* スクロール位置から、見えている列と、描く列（左右に 1 画面分の余白を足した範囲）を出す（列の仮想化）。
* 列幅は可変（Excel の列幅・非表示）なので、`colAt`（geometry の二分探索）で引く。範囲は px で決めるため、
* 列が極端に狭くても広くても描く列数は画面幅に比例する（遠い列に図や値があっても、16,384 列を描かない）。
*
* @param geo `createGeometry` の結果（列数は表の列数）
* @param scrollLeft スクロール枠の `scrollLeft`
* @param clientWidth スクロール枠の表示幅（px）。0 以下ならレイアウト前とみなして代わりの幅を使う
* @param originX 表の A1 の左上の x（行番号の列の幅。スクロール枠の内容の左端から）
* @param total 列数
*/
function visibleColRange(geo, scrollLeft, clientWidth, originX, total) {
	if (total < 1) return {
		left: 0,
		right: 0,
		first: 0,
		last: 0
	};
	const w = clientWidth > 0 ? clientWidth : COL_FALLBACK_VIEW_WIDTH;
	const x = scrollLeft - originX;
	const at = (v) => geo.colAt(Math.max(0, v));
	return {
		left: at(x),
		right: at(x + w),
		first: at(x - w),
		last: at(x + 2 * w)
	};
}
/**
* 描かない左右の列ぶんの幅（px）。左＝`colLeft(first)`、右＝`totalWidth − colLeft(last + 1)`。
* 左の余白＋描く列の幅＋右の余白が、いつも `totalWidth` に一致する（`spacerHeights` の列版）。
*/
function colSpacerWidths(geo, first, last) {
	return {
		left: Math.max(0, geo.colLeft(first)),
		right: Math.max(0, geo.totalWidth - geo.colLeft(last + 1))
	};
}
/**
* 行のセルを、描く列 first〜last（1 始まり）だけにする。返す行の `cells[i]` は列 `first + i`。
* - 描く範囲の左端が結合セルの途中（結合の起点が範囲の左外）なら、起点のセルを範囲の左端へ移し、`colSpan` を範囲内へ縮める。
*   起点が上の行にある（`rowSpan` で覆われる）セルは、これまでどおり null のまま（td を作らない）。
* - 右端で切れる結合セルは、範囲の右端までに `colSpan` を縮める。縮めたセルは結合のまま扱う（はみ出さない）。
*
* @param row 全列を持つ行
* @param first 描く最初の列
* @param last 描く最後の列
*/
function clipRowToColumns(row, first, last) {
	const a = first - 1;
	const n = Math.max(0, last - a);
	const cells = new Array(n);
	/** 範囲の左端へ移した結合の起点の列（0 始まり）。移していなければ -1 */
	let movedFrom = -1;
	for (let i = 0; i < n; i++) cells[i] = row.cells[a + i];
	if (a > 0 && n > 0 && row.cells[a] === null) {
		let o = a - 1;
		while (o >= 0 && row.cells[o] === null) o--;
		const origin = o >= 0 ? row.cells[o] : void 0;
		if (origin && origin.colSpan > 1 && o + origin.colSpan - 1 >= a) {
			cells[0] = {
				...origin,
				colSpan: o + origin.colSpan - a,
				merged: true
			};
			movedFrom = o;
		}
	}
	for (let i = 0; i < n; i++) {
		const cell = cells[i];
		if (cell && cell.colSpan > 1 && i + cell.colSpan > n) cells[i] = {
			...cell,
			colSpan: n - i,
			merged: true
		};
	}
	const clipped = {
		r: row.r,
		cells
	};
	if (row.formats) {
		const formats = new Array(n);
		for (let i = 0; i < n; i++) formats[i] = row.formats[a + i];
		if (movedFrom >= 0) formats[0] = row.formats[movedFrom];
		clipped.formats = formats;
	}
	return clipped;
}
/**
* 表の範囲（データの行・列）の外にある余白の位置へ、行の書式 → 列の書式の順で書式の番号を当てる（ADR-0030 §決定 7 の余白への拡張）。
* 読み取り（`getRows`）は表の範囲の中だけを解くため、余白の行（`r > dataRows`）と余白の列（`列 > dataCols`）は、
* Excel で行・列全体に続く塗りがここで抜ける。表の範囲は広げない（描く余白の位置に番号を足すだけ）。
* - 行の書式（`rowFormat`）があればそれ（-1＝書式なしは何も当てず、列の書式も打ち消す）。無ければ列の書式（`colFormat`・-1 は無し）。
* - 描く列 first〜last のうち、余白の位置だけを見る。表の中の位置（読み取りが解いた番号）は変えない。
* - 余白に触れる行にだけ `rowFormat` を 1 回呼ぶ。`colFormat` は列ごとに 1 回（結果を覚える）。入力の行は書き換えず、
*   足した行だけ新しい `GridRow` にする（足すものが無い行は同じ行をそのまま返す）。
*
* @param rows 描く行（`collectVisibleRows` の結果）
* @param first 描く最初の列（1 始まり）
* @param last 描く最後の列
* @param dataRows 表のデータの行数（`XlsxSheet.rowCount`）
* @param dataCols 表のデータの列数（`XlsxSheet.colCount`）
* @param rowFormat 行 r の行の書式の番号（`XlsxSheet.rowFormat`）
* @param colFormat 列 c（1 始まり）の列の書式の番号（`XlsxSheet.colFormat`）
*/
function applyMarginFormats(rows, first, last, dataRows, dataCols, rowFormat, colFormat) {
	if (last < first) return rows.slice();
	/** 列 c の列の書式（列ごとに 1 回だけ引く） */
	const colCache = /* @__PURE__ */ new Map();
	const colOf = (c) => {
		let f = colCache.get(c);
		if (f === void 0) {
			f = colFormat(c);
			colCache.set(c, f);
		}
		return f;
	};
	return rows.map((row) => {
		const from = row.r > dataRows ? first : Math.max(first, dataCols + 1);
		if (from > last) return row;
		const rf = rowFormat(row.r);
		if (rf !== void 0 && rf < 0) return row;
		let formats;
		for (let c = from; c <= last; c++) {
			if (row.formats?.[c - 1] !== void 0) continue;
			const f = rf ?? colOf(c);
			if (f >= 0) {
				formats ??= row.formats ? row.formats.slice() : [];
				formats[c - 1] = f;
			}
		}
		return formats ? {
			...row,
			formats
		} : row;
	});
}
/**
* セルの内側の要素の高さ（px）。セル（td）の高さ（行高、結合は描く行の高さの合計）から下罫線 1px を引く。0 を下回らない。
*/
function cellInnerHeight(cellHeight) {
	return Math.max(0, cellHeight - BORDER);
}
/**
* 描く行のうち、i 番目の行から span 行ぶん（結合した、描く見える行の数）の高さの合計（px）。
*
* @param rows 描く行（見える行だけ。上から順）
*/
function spanHeight(geo, rows, i, span) {
	let h = 0;
	for (let k = i; k < Math.min(rows.length, i + Math.max(1, span)); k++) h += geo.rowHeight(rows[k].r);
	return h;
}
/**
* 空のセルか（値が無く、結合も無い）。`undefined`（セルの無い列）と、型の無い（`kind` 無し）文字が空で結合の無い `CellView`。
* 数式の結果が空文字の文字列セル（`kind: 's'`・text ""）は「値のあるセル」なので空に含めない（Excel は文字のはみ出しをそこで止める）。
* 結合で隠れるセル（`null`）も空に含めない（はみ出しの先にできない）。
*/
function isEmptyCell(cell) {
	return cell === void 0 || cell !== null && cell.text === "" && cell.kind === void 0 && cell.colSpan === 1 && cell.rowSpan === 1 && cell.merged !== true;
}
/**
* 文字列がはみ出して使える列数（自分の列を含む）を決める。Excel の「文字列は右隣が空ならはみ出し、
* 値のあるセルの手前で切る」を表す。td は 1 セル 1 個のまま（罫線を残す）で、文字だけをはみ出させる。
* - 文字列のとき（`kind === 's'` で非空、結合の無いセル）だけ右隣の空セルを数える。型で判定する
*   （数字だけの文字列・「▲1,234」は文字列なのではみ出す。数値・真偽・エラー・日付、型の無いセルは自分の列で切る）。
* - 結合セル・空のセルは 1（自分の列だけ）。
* - 値のあるセル（数式の結果が "" の文字列セルを含む）・結合で隠れるセル（null）・結合セルの手前で止める。HTML の `colspan` 上限と揃えて 1000 列まで。
*
* @param cells 行のセル（`GridRow.cells`）
* @param c 対象の列（0 始まり）
* @param cols 表の列数。`cells` より右の列（足した空の列）も空のセルとして数える
* @returns 使える列数（1 以上）。1 ならはみ出さない
*/
function spillColumns(cells, c, cols = cells.length) {
	const cell = cells[c];
	if (cell === void 0 || cell === null || cell.merged === true || cell.colSpan > 1 || cell.rowSpan > 1 || cell.text === "" || cell.kind !== "s") return 1;
	let n = 1;
	while (n < MAX_SPILL_COLS && c + n < cols && isEmptyCell(cells[c + n])) n++;
	return n;
}
/**
* はみ出しに「紙でない塗り」が絡むか（ADR-0030 §決定 9）。はみ出す元のセル（列 c）と、`spillColumns` が数えた右の列のどれかに
* 白以外の塗りがあれば true（入れ物へ `is-over-fill` を付け、文字の下に面の色を敷かない）。白の塗りは塗りなしとして数える。
* 追加の走査はせず、数えた列の `formats` を見るだけ。
*
* @param table 書式の表（`XlsxSheet.formats`）
* @param formats 行の書式の番号（`GridRow.formats`。描く列に揃えたもの。無ければ false）
* @param c はみ出す元の列（0 始まり）
* @param columns `spillColumns` の結果（自分の列を含む列数）
*/
function spillOverFill(table, formats, c, columns) {
	if (!formats) return false;
	const end = Math.min(formats.length, c + columns);
	for (let i = c; i < end; i++) {
		const f = formats[i];
		if (f !== void 0 && hasOverFill(table[f])) return true;
	}
	return false;
}
/**
* はみ出す文字の入れ物の幅（px）。自分の列から `columns` 列ぶんの列幅の合計から、右端の罫線 1px を引く。
* 非表示の列（幅 0）は合計に入らない。自分の列が非表示（幅 0）なら 0（隠れたセルの文字を出さない）。
*
* @param geo `createGeometry` の結果
* @param c 対象の列（0 始まり）
* @param columns `spillColumns` の結果（自分の列を含む列数）
*/
function spillWidth(geo, c, columns) {
	if (geo.colWidth(c + 1) === 0) return 0;
	return Math.max(0, geo.colLeft(c + 1 + columns) - geo.colLeft(c + 1) - BORDER);
}
/**
* 表を画面いっぱいに広げたときの行数・列数。Excel のシートのように、データの先にも空のセル
* （既定の幅と高さ。geometry は範囲外の番号も既定の寸法で延ばす）が続くようにする。
* 行数＝max(データの行数, `minRows`, スクロール枠の高さを埋める行数)、列数も同じ。上限は Excel の最大（1,048,576 行・16,384 列）。
* 後続の画像（図のある範囲まで表を広げる）は `minRows`・`minCols` に図の最下行・最右列を渡す。
*
* @param layout `XlsxSheet.layout`
* @param dataRows データの行数（`XlsxSheet.rowCount`）
* @param dataCols データの列数（`XlsxSheet.colCount`）
* @param viewWidth スクロール枠の表示幅（px。`clientWidth`）
* @param viewHeight スクロール枠の表示高さ（px。`clientHeight`）
* @param minRows 表示する行数の最小（既定 0）
* @param minCols 表示する列数の最小（既定 0）
*/
function gridExtent(layout, dataRows, dataCols, viewWidth, viewHeight, minRows = 0, minCols = 0) {
	const geo = createGeometry(layout, MAX_GRID_ROWS, MAX_GRID_COLS);
	const fillRows = viewHeight > 0 ? geo.rowAt(viewHeight) : 0;
	const rows = Math.min(MAX_GRID_ROWS, Math.max(1, dataRows, minRows, fillRows));
	const x = viewWidth - rowNumWidth(rows);
	const fillCols = x > 0 ? geo.colAt(x) : 0;
	return {
		rows,
		cols: Math.min(MAX_GRID_COLS, Math.max(1, dataCols, minCols, fillCols))
	};
}
/**
* 行 first〜last のうち、見える行（高さ > 0）だけを DOM 用の行にして返す。
* 非表示の行は返さない。結合セル（`rowSpan` > 1）は、描く見える行の数へ縮める：
* - 先頭の行が非表示の結合は、結合内の最初に見える行へセルを移す（値はその行の td に出る）。
* - 非表示の行を跨ぐ結合の `rowSpan` は、見える行の数にする。
* 隣り合う見える行の間の非表示の行が `gapMax` 行までなら、まとめて 1 回で読む（結合の途中を正しく扱うため）。
* これを超える隙間（フィルターで大半を畳んだシートなど）は別々に読み、隙間の向こうへ続く結合は、
* 次の見える行に（同じ文字の）結合セルが現れる。
*
* @param geo `createGeometry` の結果
* @param first 描く最初の行
* @param last 描く最後の行
* @param total 行数
* @param readRows 行 a〜b（連続）を読む関数（`XlsxSheet.getRows` と同じ形・範囲外の行は呼び出し側で空の行にする）
* @param gapMax まとめて読む隙間の上限
*/
function collectVisibleRows(geo, first, last, total, readRows, gapMax = 200) {
	if (first < 1 || last < first) return [];
	const nums = [];
	for (let r = geo.rowHeight(first) > 0 ? first : nextVisibleRow(geo, first, total); r > 0 && r <= last; r = nextVisibleRow(geo, r, total)) nums.push(r);
	const out = [];
	let a = nums[0];
	let prev = nums[0];
	for (let i = 1; i < nums.length; i++) {
		const n = nums[i];
		if (n - prev - 1 > gapMax) {
			out.push(...readSegment(geo, a, prev, readRows));
			a = n;
		}
		prev = n;
	}
	if (a !== void 0 && prev !== void 0) out.push(...readSegment(geo, a, prev, readRows));
	return out;
}
/** 連続した行 a〜b を読み、結合を見える行へ寄せて、見える行だけ返す（{@link collectVisibleRows} の本体） */
function readSegment(geo, a, b, readRows) {
	const rows = readRows(a, b);
	const visible = rows.map((row) => geo.rowHeight(row.r) > 0);
	const cum = [0];
	for (const v of visible) cum.push(cum[cum.length - 1] + (v ? 1 : 0));
	const merges = [];
	rows.forEach((row, i) => {
		row.cells.forEach((cell, c) => {
			if (cell && cell.rowSpan > 1) merges.push({
				i,
				c,
				cell
			});
		});
	});
	for (const { i, c, cell } of merges) {
		const count = cum[Math.min(rows.length - 1, i + cell.rowSpan - 1) + 1] - cum[i];
		const fixed = {
			...cell,
			rowSpan: Math.max(1, count),
			merged: true
		};
		if (visible[i]) rows[i].cells[c] = fixed;
		else if (count > 0) {
			let j = i + 1;
			while (!visible[j]) j++;
			const target = rows[j];
			target.cells[c] = fixed;
			const f = rows[i].formats?.[c];
			if (f !== void 0 || target.formats?.[c] !== void 0) (target.formats ??= new Array(target.cells.length).fill(void 0))[c] = f;
		}
	}
	return rows.filter((_, i) => visible[i]);
}
/** 幅が最大桁幅 MDW の半分ほどの文字（小数点・桁区切り・コロン・スラッシュ・空白など） */
var NARROW_CHARS = /* @__PURE__ */ new Set([
	".",
	",",
	":",
	";",
	"'",
	" ",
	"/"
]);
/** MDW として使える値（不正なら Calibri 11 相当の 7） */
function safeMdw(mdw) {
	return Number.isFinite(mdw) && mdw >= 1 ? mdw : 7;
}
/**
* 数値の表示文字の推定幅（px）。DOM は測らず「文字数 × MDW」で見積もる。
* 数字・符号・E・％・通貨記号などは MDW、小数点・桁区切り・コロン・スラッシュ（日付）・空白は MDW の半分とする。
*
* @param text 表示文字
* @param mdw 最大桁幅 MDW（`SheetLayout.mdw`）
*/
function estimateNumberWidth(text, mdw) {
	const m = safeMdw(mdw);
	let w = 0;
	for (const ch of text) w += NARROW_CHARS.has(ch) ? m / 2 : m;
	return w;
}
/**
* 収まらない数値を `###` にする（Excel と同じ）。数値・日付（`kind` が `n`・`d`）の表示文字の推定幅が
* セルの内側の幅（セル幅 − 右罫線 1px − 左右の余白）を超えたら、`#` の繰り返し（内側の幅 ÷ MDW ÷ 文字サイズの比（1 以上）の切り捨て個数、最低 1）を返す。
* 収まるときは表示文字のまま。文字列・真偽・エラーは対象外（そのまま）。
*
* @param text 表示文字
* @param kind セルの型
* @param cellWidth セルの幅（px。結合は結合範囲の列幅の合計）
* @param mdw 最大桁幅 MDW
* @param sizeRatio 文字サイズの比（`CellFormat.sizeRatio`。推定幅に掛ける。既定 1。太字は掛けない・ADR-0030 §決定 10）
*/
function fitNumberText(text, kind, cellWidth, mdw, sizeRatio = 1) {
	if (kind !== "n" && kind !== "d") return text;
	const inner = cellWidth - BORDER - 6;
	if (estimateNumberWidth(text, mdw) * sizeRatio <= inner) return text;
	return "#".repeat(Math.max(1, Math.floor(inner / (safeMdw(mdw) * Math.max(1, sizeRatio)))));
}
/**
* セルの型ごとの既定の配置（Excel の「標準」。セルの配置指定は読まない）。
* 数値・日付は右、真偽・エラーは中央、文字列・型なしは左（クラス無し）。
*/
function cellAlignClass(kind) {
	if (kind === "n" || kind === "d") return "is-right";
	if (kind === "b" || kind === "e") return "is-center";
	return "";
}
// src/render/shapes.ts
/** 読み込んだ geom（読む前・失敗は null。null の間、形のある図形は代替＝種類名の破線枠で描く） */
var geomLib = null;
/** 読み込み中・読み込み済みの約束（失敗したら捨て、次の呼び出しで読み直す） */
var geomLoading = null;
/**
* 形の表と評価器（geom）を読む（M2-3）。表が大きいため render からは動的 import にし、形のある図形を描くシートを開く時だけ読む。
* 何度呼んでも読むのは 1 回（約束を共有する）。読めなければ例外（呼び出し側は代替で描き、通知を出す）。
*/
function loadGeom() {
	if (!geomLoading) geomLoading = import("../geom/index.js").then((lib) => {
		geomLib = lib;
	}, (e) => {
		geomLoading = null;
		throw e;
	});
	return geomLoading;
}
/** 1px あたりの EMU（`drawings.ts` の `EMU_PER_PX` と同じ値。循環 import を避けるためここにも持つ） */
var EMU_PER_PX$1 = 9525;
/** 調整値の既定（`adj` が無い・有限でないとき。割合 50%） */
var DEFAULT_ADJ = 5e4;
/** 調整値の分母（1/100000 の割合） */
var ADJ_DEN = 1e5;
/** 線の太さの範囲（px） */
var MIN_STROKE = 1;
var MAX_STROKE = 16;
/** 矢じりの大きさの下限の基準（線の太さがこれより細くても、この値に係数を掛ける）と上限（px） */
var MIN_ARROW_BASE = 2;
var MAX_ARROW = 48;
/** 矢じりの大きさの係数（`sm`・`med`・`lg`） */
var ARROW_FACTOR = {
	sm: 2,
	med: 3,
	lg: 5
};
/** `stealth` の切れ込みの位置（先端から長さの何割後ろか） */
var STEALTH_NOTCH = .7;
/** 1/60000 度あたりの 1 周（360 度）・各境界 */
var ROT_FULL = 216e5;
var ROT_45 = 27e5;
var ROT_135 = 81e5;
var ROT_225 = 135e5;
var ROT_315 = 189e5;
/** 同じ点とみなす距離（px）。端の接線の向きを求めるとき、これ以下の差の点は飛ばす */
var SAME_POINT = 1e-6;
/** 有限でない値を 0 にする */
function fin(n) {
	return Number.isFinite(n) ? n : 0;
}
/** 回転（1/60000 度）を [0, 360 度) に正規化する（有限でない値は 0） */
function normRot(rot) {
	if (rot === void 0 || !Number.isFinite(rot)) return 0;
	return (Math.round(rot) % ROT_FULL + ROT_FULL) % ROT_FULL;
}
/**
* アンカーの矩形から「回転前の矩形」を出す。図の回転が 45° 以上 135° 未満、または 225° 以上 315° 未満なら、
* アンカー（回転後の外接）は縦横が入れ替わっているため、幅と高さを入れ替える。中心はアンカーと同じ（ADR-0027 §決定 4）。
*
* 45°・135°・225°・315° ちょうどの境界は、Excel のアンカーが縦横を入れ替えたか入れ替えていないかが回転の値だけでは決まらない（45° ちょうどの
* 図形は入れ替えていない）。`ext`（`a:xfrm` の `a:ext`＝回転前の大きさ。px）があれば、アンカーの幅・高さを `ext` に近い向きへ合わせて決める
* （入れ替えた方が `ext` に近ければ入れ替える。差が 1px 以内＝正方形に近いなど決められないときは、回転の値で決める。ISSUE-93）。
*
* @param rect アンカーの矩形（px。幅・高さは 0 でもよい）
* @param rot 回転（1/60000 度。無ければ 0）
* @param ext 回転前の大きさ（px。無ければ回転の値だけで決める）
*/
function preRotationRect(rect, rot, ext) {
	const r = normRot(rot);
	let swap = r >= ROT_45 && r < ROT_135 || r >= ROT_225 && r < ROT_315;
	if (ext !== void 0 && ext.w > 0 && ext.h > 0) {
		const keep = Math.abs(rect.w - ext.w) + Math.abs(rect.h - ext.h);
		const swapped = Math.abs(rect.w - ext.h) + Math.abs(rect.h - ext.w);
		if (Math.abs(keep - swapped) > 1) swap = swapped < keep;
	}
	if (!swap) return { ...rect };
	const cx = rect.x + rect.w / 2;
	const cy = rect.y + rect.h / 2;
	return {
		x: cx - rect.h / 2,
		y: cy - rect.w / 2,
		w: rect.h,
		h: rect.w
	};
}
/** 図形の回転前の大きさ（`xfrmExt`。EMU）を px にする。持たない・読めないなら undefined（`preRotationRect` の `ext` に渡す） */
function xfrmExtPx(drawing) {
	const e = drawing.xfrmExt;
	return e === void 0 ? void 0 : {
		w: e.cx / EMU_PER_PX$1,
		h: e.cy / EMU_PER_PX$1
	};
}
/** 調整値 `i`（0 始まり）を割合（1.0＝100%）にする。無い・有限でないときは 50% */
function adjRatio(adj, i) {
	const v = adj?.[i];
	return (v !== void 0 && Number.isFinite(v) ? v : DEFAULT_ADJ) / ADJ_DEN;
}
/**
* 形（`prst`）と調整値から、回転前の矩形（幅 w・高さ h・左上が原点）の中の経路を作る。式は ECMA-376 の preset 定義に合わせる。
* - `line`・`straightConnector1`・未知の値：(0,0)→(w,h) の直線。
* - `bentConnector2〜5`：折れ線。`curvedConnector4・5` は同じ調整値の `bentConnector4・5` の折れ線で代替する。
* - `curvedConnector2・3`：3 次ベジェ。
* 調整値は割合のまま使い（範囲外も張り出しとして許す）、丸めは回転後の点に対して行う（`clampPath`）。
*
* @param adj 調整値 adj1〜adj3（1/100000。無い要素は 50000）
*/
function shapePath(geom, adj, w, h) {
	const L = (x, y) => ({
		t: "L",
		x,
		y
	});
	const a1 = adjRatio(adj, 0);
	const a2 = adjRatio(adj, 1);
	const a3 = adjRatio(adj, 2);
	switch (geom) {
		case "bentConnector2": return {
			start: {
				x: 0,
				y: 0
			},
			segs: [L(w, 0), L(w, h)]
		};
		case "bentConnector3": {
			const x1 = w * a1;
			return {
				start: {
					x: 0,
					y: 0
				},
				segs: [
					L(x1, 0),
					L(x1, h),
					L(w, h)
				]
			};
		}
		case "bentConnector4":
		case "curvedConnector4": {
			const x1 = w * a1;
			const y2 = h * a2;
			return {
				start: {
					x: 0,
					y: 0
				},
				segs: [
					L(x1, 0),
					L(x1, y2),
					L(w, y2),
					L(w, h)
				]
			};
		}
		case "bentConnector5":
		case "curvedConnector5": {
			const x1 = w * a1;
			const y2 = h * a2;
			const x3 = w * a3;
			return {
				start: {
					x: 0,
					y: 0
				},
				segs: [
					L(x1, 0),
					L(x1, y2),
					L(x3, y2),
					L(x3, h),
					L(w, h)
				]
			};
		}
		case "curvedConnector2": return {
			start: {
				x: 0,
				y: 0
			},
			segs: [{
				t: "C",
				x1: w / 2,
				y1: 0,
				x2: w,
				y2: h / 2,
				x: w,
				y: h
			}]
		};
		case "curvedConnector3": {
			const x2 = w * a1;
			return {
				start: {
					x: 0,
					y: 0
				},
				segs: [{
					t: "C",
					x1: x2 / 2,
					y1: 0,
					x2,
					y2: h / 4,
					x: x2,
					y: h / 2
				}, {
					t: "C",
					x1: x2,
					y1: h * 3 / 4,
					x2: (w + x2) / 2,
					y2: h,
					x: w,
					y: h
				}]
			};
		}
		default: return {
			start: {
				x: 0,
				y: 0
			},
			segs: [L(w, h)]
		};
	}
}
/** 経路の点（制御点を含む）を全部 `fn` で写した新しい経路にする */
function mapPath(path, fn) {
	const out = {
		start: fn(path.start),
		segs: path.segs.map((s) => {
			if (s.t === "L") {
				const p = fn(s);
				return {
					t: "L",
					x: p.x,
					y: p.y
				};
			}
			const c1 = fn({
				x: s.x1,
				y: s.y1
			});
			const c2 = fn({
				x: s.x2,
				y: s.y2
			});
			const e = fn(s);
			return {
				t: "C",
				x1: c1.x,
				y1: c1.y,
				x2: c2.x,
				y2: c2.y,
				x: e.x,
				y: e.y
			};
		})
	};
	if (path.closed) out.closed = true;
	return out;
}
/**
* 矩形の中の経路を、反転（`flipH`：x→w-x、`flipV`：y→h-y）してから、矩形の中心で時計回りに `rot` 回転し、
* シートの座標（px）にする。画面座標（y 下向き）では `x' = cx + dx·cos − dy·sin`、`y' = cy + dx·sin + dy·cos`。
*
* @param path 矩形の左上を原点とする経路（`shapePath` の結果）
* @param box 回転前の矩形（`preRotationRect` の結果）
*/
function transformPath(path, box, xfrm) {
	const rad = normRot(xfrm.rot) / 6e4 * (Math.PI / 180);
	const cos = Math.cos(rad);
	const sin = Math.sin(rad);
	const cx = box.x + box.w / 2;
	const cy = box.y + box.h / 2;
	return mapPath(path, (p) => {
		const lx = xfrm.flipH ? box.w - p.x : p.x;
		const ly = xfrm.flipV ? box.h - p.y : p.y;
		const dx = box.x + lx - cx;
		const dy = box.y + ly - cy;
		return {
			x: cx + dx * cos - dy * sin,
			y: cy + dx * sin + dy * cos
		};
	});
}
/**
* 経路の全点（制御点を含む）をシートの範囲（0〜全幅・0〜全高）に丸める。
* 細工した極端な調整値で、座標・`<svg>` の大きさ・表の範囲が巨大にならないための安全弁（ADR-0027 §決定 4）。
* 範囲の中の点は変えない（Excel の U 字線の張り出しは縮まない）。
*/
function clampPath(path, sheet) {
	const maxX = Math.max(0, fin(sheet.totalWidth));
	const maxY = Math.max(0, fin(sheet.totalHeight));
	return mapPath(path, (p) => ({
		x: Math.min(maxX, Math.max(0, fin(p.x))),
		y: Math.min(maxY, Math.max(0, fin(p.y)))
	}));
}
/** 経路の全点（始点・制御点・終点）を並べる */
function pathPoints(path) {
	const pts = [path.start];
	for (const s of path.segs) {
		if (s.t === "C") pts.push({
			x: s.x1,
			y: s.y1
		}, {
			x: s.x2,
			y: s.y2
		});
		pts.push({
			x: s.x,
			y: s.y
		});
	}
	return pts;
}
/** 線の太さ（px）：EMU / 9525 を 1〜16 に丸める。0・未指定・有限でない値は 1px */
function strokeWidthPx(widthEmu) {
	if (widthEmu === void 0 || !Number.isFinite(widthEmu) || widthEmu <= 0) return MIN_STROKE;
	return Math.min(MAX_STROKE, Math.max(MIN_STROKE, widthEmu / EMU_PER_PX$1));
}
/**
* 線種を `stroke-dasharray` の数値にする。`solid` は空（実線）、`dot`・`sysDot` は点線 [w, 2w]、
* それ以外は破線 [4w, 3w]（w は線の太さ px）。
*/
function dashArray(dash, strokeWidth) {
	if (dash === void 0 || dash === "solid") return [];
	if (dash === "dot" || dash === "sysDot") return [strokeWidth, strokeWidth * 2];
	return [strokeWidth * 4, strokeWidth * 3];
}
/** 矢じりの大きさ（px）：max(線の太さ, 2px) × 係数（`sm` 2・`med` 3・`lg` 5。未知は `med`）。上限 48px */
function arrowDimension(strokeWidth, size) {
	const f = size === "sm" || size === "lg" ? ARROW_FACTOR[size] : ARROW_FACTOR.med;
	return Math.min(MAX_ARROW, Math.max(strokeWidth, MIN_ARROW_BASE) * f);
}
/** `a` から `b` への距離 */
function dist(a, b) {
	return Math.hypot(b.x - a.x, b.y - a.y);
}
/**
* 端の接線の向き（単位ベクトル）：経路の外へ向かう向き。
* `tail`＝終点で、一つ前の（終点と違う）点から終点への向き。`head`＝始点で、次の（始点と違う）点から始点への向き。
* 直線・折れ線は端の線分、ベジェは端に隣る制御点→端点。長さ 0 なら更に遡る。全点が同じなら (1, 0)。
*/
function endDirection(path, end) {
	const pts = pathPoints(path);
	const tip = end === "head" ? pts[0] : pts[pts.length - 1];
	const order = end === "head" ? pts.slice(1) : pts.slice(0, -1).reverse();
	for (const p of order) {
		const d = dist(p, tip);
		if (d > SAME_POINT) return {
			x: (tip.x - p.x) / d,
			y: (tip.y - p.y) / d
		};
	}
	return {
		x: 1,
		y: 0
	};
}
/** 矢じりの型ごとの、線の端を縮める長さ（矢じりの長さに対する割合）。塗り三角は全長、stealth は切れ込みまで、他は縮めない */
function shrinkRatio(type) {
	switch (type) {
		case "arrow":
		case "diamond":
		case "oval": return 0;
		case "stealth": return STEALTH_NOTCH;
		default: return 1;
	}
}
/** 端の矢じりの大きさ（幅・長さ。px） */
function arrowSize(end, strokeWidth) {
	return {
		width: arrowDimension(strokeWidth, end.w),
		length: arrowDimension(strokeWidth, end.len)
	};
}
/**
* 矢じり 1 つの形を、端点 `tip` と外向きの単位ベクトル `dir` から作る（座標はシートの原点基準）。
* - `triangle`（と未知の値）：塗った三角。先端が端点。
* - `stealth`：後ろが切れ込んだ塗った四角形。先端が端点。
* - `arrow`：開いた V（線だけ）。先端が端点。
* - `diamond`・`oval`：塗ったひし形・楕円。中心が端点。
*
* @param strokeWidth 線の太さ（px）。大きさ＝max(太さ, 2px)×係数（上限 48px）
*/
function arrowShape(end, spec, tip, dir, strokeWidth) {
	const { width, length } = arrowSize(spec, strokeWidth);
	const at = (a, b) => ({
		x: tip.x + dir.x * a - dir.y * b,
		y: tip.y + dir.y * a + dir.x * b
	});
	const hw = width / 2;
	switch (spec.type) {
		case "stealth": return {
			kind: "polygon",
			end,
			points: [
				at(0, 0),
				at(-length, hw),
				at(-length * STEALTH_NOTCH, 0),
				at(-length, -hw)
			],
			filled: true,
			closed: true
		};
		case "arrow": return {
			kind: "polygon",
			end,
			points: [
				at(-length, hw),
				at(0, 0),
				at(-length, -hw)
			],
			filled: false,
			closed: false
		};
		case "diamond": return {
			kind: "polygon",
			end,
			points: [
				at(length / 2, 0),
				at(0, hw),
				at(-length / 2, 0),
				at(0, -hw)
			],
			filled: true,
			closed: true
		};
		case "oval": return {
			kind: "ellipse",
			end,
			cx: tip.x,
			cy: tip.y,
			rx: length / 2,
			ry: hw,
			angle: Math.atan2(dir.y, dir.x) * 180 / Math.PI
		};
		default: return {
			kind: "polygon",
			end,
			points: [
				at(0, 0),
				at(-length, hw),
				at(-length, -hw)
			],
			filled: true,
			closed: true
		};
	}
}
/** 矢じりの座標を (dx, dy) だけ動かす */
function shiftArrow(a, dx, dy) {
	if (a.kind === "ellipse") return {
		...a,
		cx: a.cx + dx,
		cy: a.cy + dy
	};
	return {
		...a,
		points: a.points.map((p) => ({
			x: p.x + dx,
			y: p.y + dy
		}))
	};
}
/** 矢じりの外接矩形（左・上・右・下） */
function arrowExtent(a) {
	if (a.kind === "ellipse") {
		const rad = a.angle * Math.PI / 180;
		const ex = Math.hypot(a.rx * Math.cos(rad), a.ry * Math.sin(rad));
		const ey = Math.hypot(a.rx * Math.sin(rad), a.ry * Math.cos(rad));
		return {
			l: a.cx - ex,
			t: a.cy - ey,
			r: a.cx + ex,
			b: a.cy + ey
		};
	}
	const xs = a.points.map((p) => p.x);
	const ys = a.points.map((p) => p.y);
	return {
		l: Math.min(...xs),
		t: Math.min(...ys),
		r: Math.max(...xs),
		b: Math.max(...ys)
	};
}
/**
* 線の端を縮めた経路を返す（塗った三角・stealth の下に線が突き出ないため）。
* 始点側は始点と最初の区間の制御点、終点側は終点と最後の区間の制御点を、端の接線に沿って動かす。
* 縮める長さは、その端の区間の長さ（弦）を超えない。区間が 1 つで両端を縮めるときは、合計が区間の長さに収まるよう按分する。
*
* @param headShrink 始点側を縮める長さ（px）
* @param tailShrink 終点側を縮める長さ（px）
*/
function shrinkPath(path, headShrink, tailShrink) {
	if (path.segs.length === 0 || headShrink <= 0 && tailShrink <= 0) return path;
	const segs = path.segs.map((s) => ({ ...s }));
	const start = { ...path.start };
	const firstSeg = segs[0];
	const lastSeg = segs[segs.length - 1];
	const firstEnd = {
		x: firstSeg.x,
		y: firstSeg.y
	};
	const prevOfLast = segs.length >= 2 ? {
		x: segs[segs.length - 2].x,
		y: segs[segs.length - 2].y
	} : {
		x: start.x,
		y: start.y
	};
	const headLen = dist(start, firstEnd);
	const tailLen = dist(prevOfLast, {
		x: lastSeg.x,
		y: lastSeg.y
	});
	let sh = Math.min(Math.max(0, headShrink), headLen);
	let st = Math.min(Math.max(0, tailShrink), tailLen);
	if (segs.length === 1 && sh + st > headLen) {
		const k = headLen / (sh + st);
		sh *= k;
		st *= k;
	}
	const hd = endDirection(path, "head");
	const td = endDirection(path, "tail");
	if (sh > 0) {
		start.x -= hd.x * sh;
		start.y -= hd.y * sh;
		if (firstSeg.t === "C") {
			firstSeg.x1 -= hd.x * sh;
			firstSeg.y1 -= hd.y * sh;
		}
	}
	if (st > 0) {
		lastSeg.x -= td.x * st;
		lastSeg.y -= td.y * st;
		if (lastSeg.t === "C") {
			lastSeg.x2 -= td.x * st;
			lastSeg.y2 -= td.y * st;
		}
	}
	return {
		start,
		segs
	};
}
/** 数値を SVG 用の文字列にする（小数 2 桁。有限でない値は 0。`-0` は `0`） */
function fmt(n) {
	return String(Number(fin(n).toFixed(2)));
}
/**
* 経路を `<path d>` の文字列にする（数値だけから組み立てる。NaN・Infinity は出さない）。閉じた経路は末尾に `Z`。
*
* @param originX 座標の原点の x（この値を引く。外接矩形の左上を原点にするときに渡す）
* @param originY 座標の原点の y
*/
function pathData(path, originX = 0, originY = 0) {
	const px = (x) => fmt(x - originX);
	const py = (y) => fmt(y - originY);
	const parts = [`M${px(path.start.x)} ${py(path.start.y)}`];
	for (const s of path.segs) if (s.t === "L") parts.push(`L${px(s.x)} ${py(s.y)}`);
	else parts.push(`C${px(s.x1)} ${py(s.y1)} ${px(s.x2)} ${py(s.y2)} ${px(s.x)} ${py(s.y)}`);
	if (path.closed) parts.push("Z");
	return parts.join(" ");
}
/** 部分経路の列を 1 つの `<path d>` にする（図形の path 1 つ。部分経路を空白でつなぐ。空なら空文字） */
function subpathsData(paths, originX = 0, originY = 0) {
	return paths.map((p) => pathData(p, originX, originY)).join(" ");
}
/** 多角形の矢じりを `<path d>` の文字列にする（閉じる図形は `Z`）。座標はそのまま（原点の移動は `renderLine` が済ませる） */
function arrowPolygonData(a) {
	const body = a.points.map((p, i) => `${i === 0 ? "M" : "L"}${fmt(p.x)} ${fmt(p.y)}`).join(" ");
	return a.closed ? `${body} Z` : body;
}
/**
* 1 本の線を、SVG に描ける数値にまとめる。
* アンカーの矩形 → 回転前の矩形 → 経路 → 反転 → 回転 → シートの範囲へ丸め → 矢じり・縮め・外接矩形。
* 外接矩形は、丸めた経路の点（制御点を含む）と矢じりの範囲の外接矩形に、線の太さの半分と矢じりの大きさの半分
* （大きい方）を四方に足し、シートの範囲に丸めたもの。水平・垂直の線（幅か高さが 0）でも太さ分の大きさを持つ。
*
* @param rect アンカーの矩形（px。回転後の外接。幅か高さが 0 でもよい）。`preRotated` が true なら回転前の矩形そのもの
* @param line 線の形と見た目
* @param xfrm 回転・反転
* @param sheet シートの範囲
* @param preRotated `rect` が回転前の矩形か（グループの中の線。`xfrm` の矩形そのもので、縦横を入れ替えない。ISSUE-89）
*/
function renderLine(rect, line, xfrm, sheet, preRotated = false) {
	const box = preRotated ? { ...rect } : preRotationRect(rect, xfrm.rot);
	const path = clampPath(transformPath(shapePath(line.geom, line.adj, box.w, box.h), box, xfrm), sheet);
	const strokeWidth = strokeWidthPx(line.widthEmu);
	const dash = dashArray(line.dash, strokeWidth);
	const pts = pathPoints(path);
	const tipHead = path.start;
	const tipTail = pts[pts.length - 1];
	const arrows = [];
	let headShrink = 0;
	let tailShrink = 0;
	let arrowMax = 0;
	if (line.head) {
		arrows.push(arrowShape("head", line.head, tipHead, endDirection(path, "head"), strokeWidth));
		const sz = arrowSize(line.head, strokeWidth);
		headShrink = sz.length * shrinkRatio(line.head.type);
		arrowMax = Math.max(arrowMax, sz.width, sz.length);
	}
	if (line.tail) {
		arrows.push(arrowShape("tail", line.tail, tipTail, endDirection(path, "tail"), strokeWidth));
		const sz = arrowSize(line.tail, strokeWidth);
		tailShrink = sz.length * shrinkRatio(line.tail.type);
		arrowMax = Math.max(arrowMax, sz.width, sz.length);
	}
	const drawn = shrinkPath(path, headShrink, tailShrink);
	const pad = Math.max(strokeWidth / 2, arrowMax / 2);
	let l = Math.min(...pts.map((p) => p.x)) - pad;
	let t = Math.min(...pts.map((p) => p.y)) - pad;
	let r = Math.max(...pts.map((p) => p.x)) + pad;
	let b = Math.max(...pts.map((p) => p.y)) + pad;
	for (const a of arrows) {
		const e = arrowExtent(a);
		l = Math.min(l, e.l - strokeWidth / 2);
		t = Math.min(t, e.t - strokeWidth / 2);
		r = Math.max(r, e.r + strokeWidth / 2);
		b = Math.max(b, e.b + strokeWidth / 2);
	}
	const maxX = Math.max(0, fin(sheet.totalWidth));
	const maxY = Math.max(0, fin(sheet.totalHeight));
	const x0 = Math.max(0, Math.floor(l));
	const y0 = Math.max(0, Math.floor(t));
	const x1 = Math.min(maxX, Math.ceil(r));
	const y1 = Math.min(maxY, Math.ceil(b));
	const bbox = {
		x: x0,
		y: y0,
		w: Math.max(0, x1 - x0),
		h: Math.max(0, y1 - y0)
	};
	return {
		bbox,
		d: pathData(drawn, bbox.x, bbox.y),
		strokeWidth,
		dash,
		heads: arrows.map((a) => shiftArrow(a, -bbox.x, -bbox.y)),
		path
	};
}
var SHADES = /* @__PURE__ */ new Set([
	"darken",
	"darkenLess",
	"lighten",
	"lightenLess"
]);
/** 図形が表示する文字を持つか（ランの文字が 1 文字以上ある。改行だけ・空のランは含めない） */
function shapeHasText(drawing) {
	return (drawing.kind === "shape" || drawing.kind === "textbox") && !!drawing.text && drawing.text.paragraphs.some((p) => p.runs.some((r) => !r.br && r.text.length > 0));
}
/** 色の不透明度が 0 でない（色が無い・不透明度が有限でないときは見えるとみなす。暫定の色で描かれるため） */
function opaqueColor(c) {
	return c === void 0 || !(c.alpha <= 0);
}
/** 塗りがあり、その色が見える（不透明度 0 超。色が無い＝暫定の色は見える）。下地の塗りの条件（ISSUE-96） */
function filledShown(drawing) {
	return drawing.hasFill === true && opaqueColor(drawing.fillColor);
}
/** 見える文字がある（文字があり、不透明度 0 超のランが 1 つ以上ある。`a:noFill` の文字は見えない）。下地の塗りの条件（ISSUE-96） */
function textShown(drawing) {
	return shapeHasText(drawing) && drawing.text.paragraphs.some((p) => p.runs.some((r) => !r.br && r.text.length > 0 && opaqueColor(r.color)));
}
/** 形（`SheetDrawing.geom`）を評価器の定義と調整値にする。表に無い `prst`・形なし・geom が未読は null */
function geomDefOf(geom) {
	if (!geom || !geomLib) return null;
	if (geom.type === "preset") {
		const def = geomLib.getPresetGeom(geom.prst);
		return def ? {
			def,
			adjust: new Map(geom.adj)
		} : null;
	}
	return {
		def: geom.def,
		adjust: void 0
	};
}
/**
* 図形 1 つを、SVG に描ける数値にまとめる（ADR-0028 §決定 2・4・5・8）。
* アンカーの矩形 → 回転前の矩形 → 形の経路（`evaluateGeom`）→ 反転 → 中心で回転 → シートの範囲に丸め → 外接矩形。
* 外接矩形は、丸めた経路の点（制御点を含む）の外接矩形に、枠線があれば太さの半分を足して整数に広げ、アンカーの矩形と合わせたもの。
* 枠も塗りも無く文字だけの図形は経路を作らず、外接矩形はアンカーの矩形。
* 文字領域は形の `rect`（回転前の矩形の中）を、反転に合わせて矩形の中で鏡写しにしたもの（`flipH` なら x を w − r へ）。
*
* @param rect アンカーの矩形（px・整数。回転後の外接）。`preRotated` が true なら回転前の矩形そのもの
* @param drawing 図形（`geom`・`outline`・`hasOutline`・`hasFill`・`rot`・`flipH`・`flipV`・`text`）
* @param sheet シートの範囲
* @param preRotated `rect` が回転前の矩形か（グループの中の図形。縦横を入れ替えない。ISSUE-89）
*/
function renderShape(rect, drawing, sheet, preRotated = false) {
	const resolved = geomDefOf(drawing.geom);
	if (!resolved) return { kind: "fallback" };
	const box = preRotated ? { ...rect } : preRotationRect(rect, drawing.rot, xfrmExtPx(drawing));
	const ev = geomLib.evaluateGeom(resolved.def, resolved.adjust, box.w, box.h);
	if (!ev.ok) return { kind: "fallback" };
	const outline = drawing.hasOutline === true;
	const filled = drawing.hasFill === true;
	if (!outline && !filled && !shapeHasText(drawing)) return { kind: "hidden" };
	const xfrm = {
		rot: drawing.rot,
		flipH: drawing.flipH,
		flipV: drawing.flipV
	};
	const strokeWidth = strokeWidthPx(drawing.outline?.widthEmu);
	const dash = dashArray(drawing.outline?.dash, strokeWidth);
	const hasText = textShown(drawing);
	const wantBase = filledShown(drawing) || hasText;
	const drawn = [];
	for (const p of ev.paths) {
		const fillable = p.fill !== "none";
		const fill = filled && fillable;
		const stroke = outline && p.stroke;
		const base = wantBase && fillable;
		if (!fill && !stroke && !base) continue;
		drawn.push({
			fill,
			stroke,
			base,
			paint: fill || stroke,
			shade: fill && SHADES.has(p.fill) ? p.fill : null,
			subpaths: p.subpaths.map((sp) => clampPath(transformPath({
				start: sp.start,
				segs: sp.segs,
				closed: sp.closed
			}, box, xfrm), sheet))
		});
	}
	const painted = drawn.filter((f) => f.paint);
	let l = rect.x;
	let t = rect.y;
	let r = rect.x + rect.w;
	let b = rect.y + rect.h;
	const pts = painted.flatMap((f) => f.subpaths.flatMap(pathPoints));
	if (pts.length > 0) {
		const pad = outline ? 2 * strokeWidth : 0;
		const maxX = Math.max(0, fin(sheet.totalWidth));
		const maxY = Math.max(0, fin(sheet.totalHeight));
		let pl = Infinity;
		let pt = Infinity;
		let pr = -Infinity;
		let pb = -Infinity;
		for (const p of pts) {
			pl = Math.min(pl, p.x);
			pt = Math.min(pt, p.y);
			pr = Math.max(pr, p.x);
			pb = Math.max(pb, p.y);
		}
		l = Math.min(l, Math.max(0, Math.floor(pl - pad)));
		t = Math.min(t, Math.max(0, Math.floor(pt - pad)));
		r = Math.max(r, Math.min(maxX, Math.ceil(pr + pad)));
		b = Math.max(b, Math.min(maxY, Math.ceil(pb + pad)));
	}
	const bbox = {
		x: l,
		y: t,
		w: r - l,
		h: b - t
	};
	const tr = ev.textRect;
	const textRect = {
		x: drawing.flipH ? box.w - tr.r : tr.l,
		y: drawing.flipV ? box.h - tr.b : tr.t,
		w: tr.r - tr.l,
		h: tr.b - tr.t
	};
	const ds = new Map(painted.map((f) => [f, subpathsData(f.subpaths, bbox.x, bbox.y)]));
	const backdrop = [];
	for (const f of drawn) if (f.base || f.stroke) backdrop.push({
		d: ds.get(f) ?? subpathsData(f.subpaths, bbox.x, bbox.y),
		fill: f.base,
		stroke: f.stroke
	});
	if (hasText && !backdrop.some((b) => b.fill)) {
		const d = pathData(clampPath(transformPath({
			start: {
				x: 0,
				y: 0
			},
			segs: [
				{
					t: "L",
					x: box.w,
					y: 0
				},
				{
					t: "L",
					x: box.w,
					y: box.h
				},
				{
					t: "L",
					x: 0,
					y: box.h
				}
			],
			closed: true
		}, box, xfrm), sheet), bbox.x, bbox.y);
		backdrop.unshift({
			d,
			fill: true,
			stroke: false
		});
	}
	return {
		kind: "drawn",
		rect: bbox,
		box,
		textRect,
		faces: painted.map((f) => ({
			d: ds.get(f),
			fill: f.fill,
			stroke: f.stroke,
			shade: f.shade,
			subpaths: f.subpaths
		})),
		backdrop,
		strokeWidth,
		dash
	};
}
// src/render/drawings.ts
/** 1px あたりの EMU（Office の英語メートル単位。px＝EMU / 9525） */
var EMU_PER_PX = 9525;
/** 座標・大きさの上限（px）。ブラウザのレイアウトの上限（約 33,554,431px）を超えないようにする */
var MAX_PX = 33e6;
/** `SheetDrawing.line` が無い線（壊れたデータ）に使う既定：細い直線・矢じりなし */
var DEFAULT_LINE = {
	geom: "line",
	adj: [
		5e4,
		5e4,
		5e4
	],
	widthEmu: 0,
	dash: "solid",
	head: null,
	tail: null
};
/** EMU を px（整数）にする。有限でない値は 0、負は 0、上限に丸める */
function emuToPx(emu) {
	if (!Number.isFinite(emu) || emu <= 0) return 0;
	return Math.min(MAX_PX, Math.round(emu / EMU_PER_PX));
}
/** 0 始まりの番号を [0, max] の整数に丸める（有限でない値は 0） */
function clampIndex(n, max) {
	if (!Number.isFinite(n)) return 0;
	return Math.min(max, Math.max(0, Math.trunc(n)));
}
/**
* アンカーの点の x（px）：その列の左端＋ずれ。ずれが列の幅を超える時は列の右端で止める
* （列幅の僅差で隣のセルへ飛び出さない。ADR-0026 §決定 8）。
* geometry の番号は 1 始まり、アンカーの `col` は 0 始まり。
*/
function anchorPointX(geo, p) {
	const c = clampIndex(p.col, MAX_GRID_COLS - 1) + 1;
	return geo.colLeft(c) + Math.min(emuToPx(p.colOff), geo.colWidth(c));
}
/** アンカーの点の y（px）：その行の上端＋ずれ（行の高さを超える時は行の下端で止める） */
function anchorPointY(geo, p) {
	const r = clampIndex(p.row, MAX_GRID_ROWS - 1) + 1;
	return geo.rowTop(r) + Math.min(emuToPx(p.rowOff), geo.rowHeight(r));
}
/**
* アンカーを A1 の左上基準の矩形（px）にする。幅・高さは 0 や負の値のままで、大きさの判定はしない
* （線は幅か高さが 0 でも描くため。判定は `anchorRect` と `layoutDrawings` が行う）。未知の種類は null。
*/
function rawAnchorRect(geo, anchor) {
	let rect;
	switch (anchor.type) {
		case "twoCell": {
			const x = anchorPointX(geo, anchor.from);
			const y = anchorPointY(geo, anchor.from);
			rect = {
				x,
				y,
				w: anchorPointX(geo, anchor.to) - x,
				h: anchorPointY(geo, anchor.to) - y
			};
			break;
		}
		case "oneCell":
			rect = {
				x: anchorPointX(geo, anchor.from),
				y: anchorPointY(geo, anchor.from),
				w: emuToPx(anchor.cx),
				h: emuToPx(anchor.cy)
			};
			break;
		case "absolute":
			rect = {
				x: emuToPx(anchor.x),
				y: emuToPx(anchor.y),
				w: emuToPx(anchor.cx),
				h: emuToPx(anchor.cy)
			};
			break;
		default: return null;
	}
	return rect;
}
/** 子の座標系（`ch`）の矩形 `r` を、親の矩形 `parent`（px）へ線形に写す。`ch` の大きさは 1 以上 */
function mapChild(parent, ch, r) {
	const kx = parent.w / ch.cx;
	const ky = parent.h / ch.cy;
	return {
		x: parent.x + (r.x - ch.x) * kx,
		y: parent.y + (r.y - ch.y) * ky,
		w: r.cx * kx,
		h: r.cy * ky
	};
}
/**
* グループの中の図の矩形（px・整数）を出す（ADR-0027 §決定 5）。
* `anchor` は最も外側のグループのアンカーの矩形。グループの子の座標系（`chOff`・`chExt`）から親の矩形へ、外側から順に線形に写す：
* `x = 親.x + (子.off.x − chOff.x) × 親.w / chExt.cx`（y・幅・高さも同じ）。入れ子のグループは、内側のグループ自身の `off`・`ext` を外側の座標系から写して、その矩形を次の親にする。
* グループ自身の回転・反転は無視する。図の `xfrm` は回転前の矩形そのものなので、縦横の入れ替えはしない。
* 端は四捨五入して幅・高さは端の差で出す。矩形は A1 の左・上へはみ出しても切らない（切ると線の傾き・文字の折り返しが変わる。はみ出しは
* 線の `clampPath` とレイヤー（`overflow: hidden`）が隠す）。位置は ±MAX_PX、大きさは MAX_PX までで止める。座標系が壊れている・有限でない値になるときは null。
*/
function groupChildRect(anchor, group) {
	const { levels, xfrm } = group;
	if (levels.length === 0) return null;
	let cur = {
		x: anchor.x,
		y: anchor.y,
		w: Math.max(0, anchor.w),
		h: Math.max(0, anchor.h)
	};
	for (let i = 0; i < levels.length; i++) {
		const lv = levels[i];
		if (!(lv.ch.cx > 0 && lv.ch.cy > 0)) return null;
		const target = i + 1 < levels.length ? levels[i + 1].own : xfrm;
		if (target === null) return null;
		cur = mapChild(cur, lv.ch, target);
	}
	const edge = (v) => Math.min(MAX_PX, Math.max(-33e6, Math.round(v)));
	const left = edge(cur.x);
	const top = edge(cur.y);
	const right = edge(cur.x + cur.w);
	const bottom = edge(cur.y + cur.h);
	if (![
		left,
		top,
		right,
		bottom
	].every(Number.isFinite)) return null;
	return {
		x: left,
		y: top,
		w: Math.min(MAX_PX, Math.max(0, right - left)),
		h: Math.min(MAX_PX, Math.max(0, bottom - top))
	};
}
/**
* 図の矩形（px）。トップレベルの図はアンカーの矩形、グループの中の図は最も外側のグループのアンカーから写した矩形
* （幅・高さは 0 や負のままで、大きさの判定はしない。判定は `layoutDrawings` が行う）。
*/
function placeDrawing(geo, drawing) {
	const raw = rawAnchorRect(geo, drawing.anchor);
	return raw === null || drawing.group === void 0 ? raw : groupChildRect(raw, drawing.group);
}
/**
* 図の一覧を配置する。`index` を元の位置のまま持つので重ね順が保たれる。
* - グループの中の図（`group` を持つ）：矩形は最も外側のグループのアンカーから写した値（`placeDrawing`）。`xfrm` は回転前の矩形なので、線の描画では縦横を入れ替えない。
* - 図形（kind が shape・textbox で形を持つもの。ISSUE-92）：大きさが 0 以下の図は除く。形を評価できれば `rect` は経路の外接矩形とアンカーの矩形を合わせた矩形で、
*   `shape` に描画結果を持つ。枠も塗りも文字も無い図形（Excel でも見えない）は除く。形を評価できない図形は、次の「線以外」と同じ（代替の描き方にする）。
* - 線以外：大きさが 0 以下の図は除く。`rect` はアンカーの矩形（グループの中の図は写した矩形）。
* - 線（kind が line）：アンカーの幅か高さが 0 でも残す（水平・垂直の線）。幅・高さが両方 0 以下のときだけ除く
*   （負の値＝逆向きの twoCell は 0 とみなす）。`rect` は `renderLine` の外接矩形（シートの範囲に丸め済み）で、
*   `line` に描画結果を持つ。外接矩形が空（シートの範囲の外へ追い出された）なら除く。
*   `SheetDrawing.line` が無い線（壊れたデータ）は既定の直線として扱う。
*/
function layoutDrawings(geo, drawings) {
	const boxes = [];
	drawings.forEach((d, index) => {
		if (d.kind === "line") {
			const raw = placeDrawing(geo, d);
			if (!raw || raw.w <= 0 && raw.h <= 0) return;
			const line = renderLine({
				x: raw.x,
				y: raw.y,
				w: Math.max(0, raw.w),
				h: Math.max(0, raw.h)
			}, d.line ?? DEFAULT_LINE, {
				rot: d.rot,
				flipH: d.flipH,
				flipV: d.flipV
			}, {
				totalWidth: geo.totalWidth,
				totalHeight: geo.totalHeight
			}, d.group !== void 0);
			if (line.bbox.w > 0 && line.bbox.h > 0) boxes.push({
				index,
				rect: line.bbox,
				line
			});
			return;
		}
		const rect = placeDrawing(geo, d);
		if (rect && rect.w > 0 && rect.h > 0) {
			if ((d.kind === "shape" || d.kind === "textbox") && d.geom !== void 0) {
				const shape = renderShape(rect, d, {
					totalWidth: geo.totalWidth,
					totalHeight: geo.totalHeight
				}, d.group !== void 0);
				if (shape.kind === "hidden") return;
				if (shape.kind === "drawn") {
					boxes.push({
						index,
						rect: shape.rect,
						shape
					});
					return;
				}
			}
			boxes.push({
				index,
				rect
			});
		}
	});
	return boxes;
}
/**
* 図を収めるのに必要な表の行数・列数（最下行・最右列。1 始まり）。図が無ければ 0。
* `gridExtent` の `minRows`・`minCols` に渡して、図のある範囲まで表を広げる。
*
* @param geo 行数・列数の上限を Excel の最大にした geometry（`rowAt`・`colAt` が図の端を含む行・列へ届くように）
*/
function drawingExtent(geo, boxes) {
	let rows = 0;
	let cols = 0;
	for (const { rect } of boxes) {
		rows = Math.max(rows, geo.rowAt(Math.max(0, rect.y + rect.h - 1)));
		cols = Math.max(cols, geo.colAt(Math.max(0, rect.x + rect.w - 1)));
	}
	return {
		rows,
		cols
	};
}
/**
* スクロール枠の可視範囲（A1 の左上基準・px）に、上下左右 1 画面分の余白を足した範囲。
* 表の原点（A1 の左上）はスクロール枠の内容の (`originX`, `originY`)＝行番号の列の幅・見出し行の高さ。
* 枠の大きさが 0（まだ DOM に付いていない）なら幅・高さ 0 の範囲を返す（何も交差しない）。
*/
function viewWindow(scrollLeft, scrollTop, clientWidth, clientHeight, originX, originY) {
	const w = Math.max(0, clientWidth);
	const h = Math.max(0, clientHeight);
	return {
		x: scrollLeft - originX - w,
		y: scrollTop - originY - h,
		w: w === 0 ? 0 : w * 3,
		h: h === 0 ? 0 : h * 3
	};
}
/** 2 つの矩形が交わるか（接するだけは交わらない。幅・高さ 0 の矩形は何とも交わらない） */
function rectsIntersect(a, b) {
	return a.w > 0 && a.h > 0 && b.w > 0 && b.h > 0 && a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;
}
/** 範囲 `win` と交わる図だけを、元の順（重ね順）で返す。図は最大 1,000 個なので毎回の線形走査でよい */
function pickVisible(boxes, win) {
	return boxes.filter((b) => rectsIntersect(b.rect, win));
}
// src/render/grid-table.ts
/** 上下の描かない行の高さを持つ空の行（行仮想化の余白）。 */
function spacerRow(height, cols) {
	const tr = document.createElement("tr");
	tr.className = "xs-spacer";
	tr.setAttribute("aria-hidden", "true");
	const px = `${Math.max(0, height)}px`;
	tr.style.height = px;
	const td = document.createElement("td");
	td.colSpan = cols + 1;
	td.style.height = px;
	tr.append(td);
	return tr;
}
/**
* 文字がはみ出す入れ物にする：幅を右の空セルの列幅の合計まで広げ（`overflow: hidden` で値のあるセルの手前で切る）、
* 文字を span で包む（`style.css` の `.is-spill .xs-text` が背景色を持ち、文字の範囲だけ下の罫線を覆う。
* 塗りが絡むときは呼び出し側が `is-over-fill` を足し、その背景を外す）。
* td 自体は `overflow: visible` のため、広げた入れ物が右の空セルの上へ出る。
*/
function spillOut(box, text, width) {
	const label = document.createElement("span");
	label.className = "xs-text";
	label.textContent = text;
	box.classList.add("is-spill");
	box.style.width = `${width}px`;
	box.replaceChildren(label);
}
/**
* 塗りの絡むはみ出し（`is-over-fill`）の文字色を、列の面に合わせて塗り分ける（ADR-0030 §決定 9）。
* 元の文字が「自動」色で、またぐ列の面が塗りあり（黒）と塗りなし（画面の文字色）に分かれるときだけ、span に
* 列ごとの区間のグラデーション＋`background-clip: text`＋透明な文字色をインラインで付ける（はみ出しはまれなので少数）。
* 明示色・面が同じ・span の無い入れ物には何もしない（単色のまま）。
*
* @param k 描く列の中での元の位置（`formats` の添字）
* @param c 元の列（0 始まり。geometry は c + 1）
* @param columns `spillColumns` の結果
*/
function inkSpill(box, table, formats, k, geo, c, columns) {
	const label = box.firstElementChild;
	if (!(label instanceof HTMLElement)) return;
	const widths = [];
	for (let i = 0; i < columns; i++) widths.push(geo.colWidth(c + 1 + i));
	const image = spillInkGradient(table, formats, k, widths);
	if (image === null) return;
	label.style.setProperty("background-image", image);
	label.style.setProperty("-webkit-background-clip", "text");
	label.style.setProperty("background-clip", "text");
	label.style.setProperty("color", "transparent");
}
/**
* 内側の要素（高さ固定・はみ出しを隠す）を持つ td の中身を作る。
* 行の高さを固定するため、文字は高さ固定の `div` の中に入れる（`style.css` の `.xs-cell`）。
* セル値は `textContent` のみで入れる。
*/
function cellBox(text, height) {
	const box = document.createElement("div");
	box.className = "xs-cell";
	box.style.height = `${height}px`;
	box.textContent = text;
	return box;
}
/**
* グリッドの table を Excel の寸法（geometry）で固定して組む（計測テスト `format.perf.test.ts` のため export）：
* `colgroup`（列幅）・見出し行（列記号）・上の余白・行（Excel の行番号＋セル）・下の余白。
* - 列幅は `<col>` の `width`、行の高さは `<tr>` の `height` とセルの内側の要素の高さで固定する（`table-layout: fixed`）。
* - 非表示の列は幅 0 とし、`<col>` に `xs-collapsed`（`visibility: collapse`）を付けて畳む（ADR-0026 §決定 5）。
*   非表示の行は `rows` に含まれない（呼び出し側で除く）。結合セルの `rowSpan` は描く行の数。
* - セル値は `textContent` のみで入れる。結合で隠れるセル（null）は td を作らない。td は 1 セル 1 個（罫線を残す）。
*   文字列のはみ出しは、内側の要素を右の空セルの列幅の合計まで広げ、文字の範囲だけ背景色で下の罫線を覆う。
*   セルの幅（結合を含む）が 0 なら内側の要素を作らない（畳んだ列の文字を出さない）。
* - 配置は型ごとの既定（数値・日付は右、真偽・エラーは中央、文字列は左）。収まらない数値・日付は `###` にし、元の値を title に入れる
*   （大きな文字のセルは推定幅に文字サイズの比を掛ける）。
* - 書式（塗り・文字の装飾）のあるセルの td には、className に `xs-xf-<番号>` を 1 語足すだけ（規則は表示の根の `<style>`。style 属性は書かない）。
*   塗り（白以外）が絡むはみ出しの入れ物には `is-over-fill` を付ける（ADR-0030 §決定 9）。
* - 列も仮想化する：描く列は `range`（可視範囲の左右 1 画面分）だけ。描かない左右の列は、幅 `colLeft` の差の
*   余白の列（`<col>` 1 本＋各行に空の td 1 個）で置き換え、表の全幅は geometry の totalWidth のまま保つ。
*   範囲の端で切れる結合セルは `clipRowToColumns` が `colSpan` を縮める。
*
* @param colName 列番号（1 始まり）→ 列記号
* @param cols 表の列数（データの列数より広いことがある）
* @param rows 描く行（見える行だけ。行番号の昇順）
* @param geo 寸法の換算結果（`createGeometry`）
* @param rowCount 表の行数（データの行数より広いことがある）
* @param mdw 最大桁幅 MDW（`SheetLayout.mdw`。収まらない数値を `###` にする推定幅に使う）
* @param range 描く列の範囲（`first`〜`last`・1 始まり）。表の全列を描くときは 1〜`cols`
* @param formats 書式の表（`XlsxSheet.formats`。`rows[i].formats` の値が指す）
*/
function buildGridTable(colName, cols, rows, geo, rowCount, mdw, range, formats) {
	const table = document.createElement("table");
	table.className = "xs-grid";
	table.setAttribute("aria-rowcount", String(rowCount + 1));
	const numWidth = rowNumWidth(rowCount);
	table.style.width = `${numWidth + geo.totalWidth}px`;
	const { first, last } = range;
	const pad = colSpacerWidths(geo, first, last);
	const leftPad = first > 1 && pad.left > 0 ? pad.left : 0;
	const rightPad = last < cols && pad.right > 0 ? pad.right : 0;
	/** DOM 上の列数（行番号の列＋余白の列＋描く列） */
	const domCols = 1 + (leftPad > 0 ? 1 : 0) + (last - first + 1) + (rightPad > 0 ? 1 : 0);
	const padCol = (width) => {
		const col = document.createElement("col");
		col.style.width = `${width}px`;
		return col;
	};
	const group = document.createElement("colgroup");
	const numCol = document.createElement("col");
	numCol.style.width = `${numWidth}px`;
	group.append(numCol);
	if (leftPad > 0) group.append(padCol(leftPad));
	for (let c = first; c <= last; c++) {
		const col = document.createElement("col");
		const w = geo.colWidth(c);
		col.style.width = `${w}px`;
		if (w === 0) col.className = "xs-collapsed";
		group.append(col);
	}
	if (rightPad > 0) group.append(padCol(rightPad));
	table.append(group);
	const head = table.createTHead().insertRow();
	const corner = document.createElement("th");
	corner.className = "xs-rownum";
	head.append(corner);
	if (leftPad > 0) head.append(document.createElement("th"));
	for (let c = first; c <= last; c++) {
		const th = document.createElement("th");
		th.textContent = colName(c);
		head.append(th);
	}
	if (rightPad > 0) head.append(document.createElement("th"));
	const rowPad = spacerHeights(geo, rows[0]?.r ?? 1, rows[rows.length - 1]?.r ?? 0);
	const tbody = table.createTBody();
	tbody.append(spacerRow(rowPad.top, domCols - 1));
	rows.forEach((source, i) => {
		const row = clipRowToColumns(source, first, last);
		const n = last - first + 1;
		const tr = document.createElement("tr");
		tr.setAttribute("aria-rowindex", String(row.r + 1));
		const h = geo.rowHeight(row.r);
		tr.style.height = `${h}px`;
		const rowHead = document.createElement("td");
		rowHead.className = "xs-rownum";
		rowHead.append(cellBox(String(row.r), cellInnerHeight(h)));
		tr.append(rowHead);
		if (leftPad > 0) tr.append(document.createElement("td"));
		for (let k = 0; k < n; k++) {
			const c = first - 1 + k;
			const cell = row.cells[k];
			if (cell === null) continue;
			const td = document.createElement("td");
			const fi = row.formats?.[k];
			const fmt = fi === void 0 ? void 0 : formats[fi];
			if (fi !== void 0 && fmt) td.className = formatClassName(fi);
			if (cell) {
				const colSpan = Math.max(1, cell.colSpan);
				const width = geo.colLeft(c + 1 + colSpan) - geo.colLeft(c + 1);
				if (cell.text !== "" && width > 0) {
					const shown = fitNumberText(cell.text, cell.kind, width, mdw, fmt?.sizeRatio ?? 1);
					const box = cellBox(shown, cellInnerHeight(spanHeight(geo, rows, i, cell.rowSpan)));
					const align = cellAlignClass(cell.kind);
					if (align !== "") box.classList.add(align);
					if (shown !== cell.text) box.title = cell.text;
					const columns = spillColumns(row.cells, k, n);
					if (columns > 1) {
						spillOut(box, cell.text, spillWidth(geo, c, columns));
						if (spillOverFill(formats, row.formats, k, columns)) {
							box.classList.add("is-over-fill");
							inkSpill(box, formats, row.formats, k, geo, c, columns);
						}
					}
					td.append(box);
				}
				if (colSpan > 1) td.colSpan = colSpan;
				if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
			}
			tr.append(td);
		}
		if (rightPad > 0) tr.append(document.createElement("td"));
		tbody.append(tr);
	});
	tbody.append(spacerRow(rowPad.bottom, domCols - 1));
	return table;
}
// src/render/image-store.ts
/** 画像ストアを作る */
function createImageStore(deps) {
	const createUrl = deps.createUrl ?? ((blob) => URL.createObjectURL(blob));
	const revokeUrl = deps.revokeUrl ?? ((url) => URL.revokeObjectURL(url));
	const done = /* @__PURE__ */ new Map();
	const pending = /* @__PURE__ */ new Map();
	let disposed = false;
	let overBudgetNotified = false;
	/** 未開始の読み込み（先頭から 1 本ずつ実行する）。dispose で破棄する */
	const queue = [];
	let busy = false;
	const start = async (path) => {
		if (disposed) return null;
		let result;
		try {
			result = await deps.read(path);
		} catch {
			result = {
				ok: false,
				reason: "broken",
				label: null
			};
		}
		if (disposed) return null;
		let entry;
		if (result.ok) {
			const blob = new Blob([result.bytes], { type: result.mime });
			entry = {
				ok: true,
				url: createUrl(blob)
			};
		} else {
			entry = {
				ok: false,
				reason: result.reason,
				label: result.label
			};
			if (result.reason === "over-budget" && !overBudgetNotified) {
				overBudgetNotified = true;
				deps.onOverBudget?.();
			}
		}
		done.set(path, entry);
		return entry;
	};
	/** 待ち行列を 1 本ずつ処理する。破棄済みなら `read` を呼ばず null で終える */
	const pump = async () => {
		if (busy) return;
		busy = true;
		try {
			for (let job = queue.shift(); job; job = queue.shift()) job.resolve(disposed ? null : await start(job.path).catch(() => null));
		} finally {
			busy = false;
		}
	};
	return {
		peek: (path) => done.get(path),
		load(path) {
			if (disposed) return Promise.resolve(null);
			const known = done.get(path);
			if (known) return Promise.resolve(known);
			let p = pending.get(path);
			if (!p) {
				p = new Promise((resolve) => {
					queue.push({
						path,
						resolve
					});
					pump();
				}).finally(() => pending.delete(path));
				pending.set(path, p);
			}
			return p;
		},
		dispose() {
			if (disposed) return;
			disposed = true;
			for (const job of queue.splice(0)) job.resolve(null);
			for (const entry of done.values()) if (entry.ok) revokeUrl(entry.url);
			done.clear();
			pending.clear();
		},
		get urlCount() {
			let n = 0;
			for (const entry of done.values()) if (entry.ok) n++;
			return n;
		}
	};
}
// src/render/image-layer.ts
/** 図の種類名（破線枠の文字）の文言のキー。画像の `image` は「表示できない画像」の枠に使う */
var KIND_KEYS = {
	image: "imageBroken",
	chart: "chart",
	smartart: "smartart",
	shape: "shape",
	textbox: "textbox",
	group: "group",
	line: "shape"
};
/**
* 破線枠に入れる種類名。画像（`kind: 'image'`）は、描けない形式で形式名が分かれば「画像（EMF）」、それ以外は「画像（表示できない）」。
*
* @param entry 画像の読み込み失敗（画像以外・読み込み前は不要）
* @param messages 文言の表（既定は日本語）
*/
function frameLabel(drawing, entry, messages = MESSAGES_JA) {
	if (drawing.kind === "image" && entry && !entry.ok && entry.reason === "unsupported" && entry.label) return formatMessage(messages.imageFormat, { format: entry.label });
	return messages[KIND_KEYS[drawing.kind]];
}
/** 種類名入りの破線枠にする（中身を置き換える） */
function fillFrame(el, label) {
	const frame = document.createElement("div");
	frame.className = "xs-frame";
	frame.textContent = label;
	el.replaceChildren(frame);
}
var SVG_NS = "http://www.w3.org/2000/svg";
/** CSS の px（1pt = 96/72 px） */
var PX_PER_PT = 96 / 72;
/** 文字の大きさの既定（pt。Excel の既定） */
var DEFAULT_PT = 11;
/** 内側の余白の既定（EMU。Excel の既定 l/r 0.1 インチ・t/b 0.05 インチ） */
var DEFAULT_INSETS = {
	l: 91440,
	t: 45720,
	r: 91440,
	b: 45720
};
/** 行の高さ（文字の大きさの倍率） */
var LINE_HEIGHT = 1.2;
/** 段落の揃え → `text-align`（論理値。縦書きでも合う） */
var TEXT_ALIGN = {
	l: "start",
	ctr: "center",
	r: "end",
	just: "justify",
	dist: "justify"
};
/** 縦位置 → flex の `justify-content`（column の主軸＝ブロック方向なので縦書きでも合う） */
var TEXT_ANCHOR = {
	t: "flex-start",
	ctr: "center",
	b: "flex-end"
};
/** 文字の向き → `writing-mode`・`text-orientation`（表に無い値＝横書き） */
var TEXT_VERT = {
	eaVert: { mode: "vertical-rl" },
	wordArtVertRtl: { mode: "vertical-rl" },
	mongolianVert: { mode: "vertical-lr" },
	vert: {
		mode: "vertical-rl",
		orientation: "sideways"
	},
	vert270: { mode: "sideways-lr" },
	wordArtVert: {
		mode: "vertical-rl",
		orientation: "upright"
	}
};
/**
* 文字を描く図形か（kind が shape・textbox で、表示する文字が 1 文字以上ある。改行だけ・空のランは含めない）。
* 該当しない図形は種類名入りの破線枠になる。
*/
function hasShapeText(drawing) {
	return shapeHasText(drawing);
}
/** 有限でない値を 0 にして小数 2 桁へ丸める（style・属性に入れる数値用） */
function num(n) {
	return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
}
/**
* lib の色（整数 RGB・不透明度）を検証して `Paint` にする。RGB が 0〜0xFFFFFF の整数でない・不透明度が有限でないときは null（入れない）。
* 不透明度は 0〜1 に収める。ファイルの文字列は通らない（数値だけから `#rrggbb` を組み立てる）。
*/
function paintOf(c) {
	if (!c || !Number.isInteger(c.rgb) || c.rgb < 0 || c.rgb > 16777215 || !Number.isFinite(c.alpha)) return null;
	return {
		hex: `#${c.rgb.toString(16).padStart(6, "0")}`,
		alpha: Math.round(Math.min(1, Math.max(0, c.alpha)) * 1e3) / 1e3
	};
}
/** SVG 要素の style に塗り（`fill`）か線（`stroke`）の色を入れる。不透明度が 1 未満なら `-opacity` も入れる（要素の style は CSS のクラスの規則に勝つ） */
function setPaint(el, prop, p) {
	if (!p) return;
	el.style.setProperty(prop, p.hex);
	if (p.alpha < 1) el.style.setProperty(`${prop}-opacity`, String(p.alpha));
}
/** HTML 用の色（`color`・背景・枠線）：不透明なら `#rrggbb`、不透明度があれば `#rrggbbaa` */
function cssColor(p) {
	return p.alpha < 1 ? `${p.hex}${Math.round(p.alpha * 255).toString(16).padStart(2, "0")}` : p.hex;
}
/** 文字の大きさ（pt）を px にする。有限でない・0 以下は既定、`fontScale` を掛ける */
function fontPx(pt, scale) {
	const base = pt !== void 0 && Number.isFinite(pt) && pt > 0 ? pt : DEFAULT_PT;
	const k = scale !== void 0 && Number.isFinite(scale) && scale > 0 ? scale : 1;
	return num(base * PX_PER_PT * k);
}
/** 余白（EMU）を px にする。未指定・有限でない・負は既定値 */
function insetPx(emu, fallback) {
	return num((emu !== void 0 && Number.isFinite(emu) && emu >= 0 ? emu : fallback) / EMU_PER_PX);
}
/** 文字の枠（図形の中身）を作る。段落＝`div`・ラン＝`span`・改行＝`br`。本文はすべて `textContent` */
function fillText(el, text, w, h) {
	const scale = text.fontScale;
	const ins = text.insets ?? DEFAULT_INSETS;
	const cap = (px, size) => num(Math.min(px, Math.max(0, size) / 2));
	el.style.paddingLeft = `${cap(insetPx(ins.l, DEFAULT_INSETS.l), w)}px`;
	el.style.paddingTop = `${cap(insetPx(ins.t, DEFAULT_INSETS.t), h)}px`;
	el.style.paddingRight = `${cap(insetPx(ins.r, DEFAULT_INSETS.r), w)}px`;
	el.style.paddingBottom = `${cap(insetPx(ins.b, DEFAULT_INSETS.b), h)}px`;
	el.style.justifyContent = TEXT_ANCHOR[text.anchor] ?? TEXT_ANCHOR.t;
	el.style.whiteSpace = text.noWrap ? "pre" : "pre-wrap";
	const vert = TEXT_VERT[text.vert];
	if (vert) {
		el.style.setProperty("writing-mode", vert.mode);
		if (vert.orientation) el.style.setProperty("text-orientation", vert.orientation);
	}
	const paras = [];
	for (const p of text.paragraphs) {
		const para = document.createElement("div");
		para.className = "xs-para";
		para.style.textAlign = TEXT_ALIGN[p.align] ?? TEXT_ALIGN.l;
		const shown = p.runs.filter((r) => !r.br && r.text.length > 0);
		if (shown.length === 0) {
			const endPx = fontPx(p.endSizePt, scale);
			para.style.fontSize = `${endPx}px`;
			para.style.setProperty("min-block-size", `${num(endPx * LINE_HEIGHT)}px`);
		} else para.style.fontSize = `${Math.max(...shown.map((r) => fontPx(r.sizePt, scale)))}px`;
		if (p.bullet) {
			const firstRun = p.runs.find((r) => !r.br);
			const b = document.createElement("span");
			b.textContent = `${p.bullet} `;
			b.style.fontSize = `${fontPx(firstRun?.sizePt ?? p.endSizePt, scale)}px`;
			const bulletPaint = paintOf(firstRun?.color);
			if (bulletPaint) b.style.color = cssColor(bulletPaint);
			para.append(b);
		}
		for (const r of p.runs) {
			if (r.br) {
				para.append(document.createElement("br"));
				continue;
			}
			const span = document.createElement("span");
			span.textContent = r.text;
			span.style.fontSize = `${fontPx(r.sizePt, scale)}px`;
			const runPaint = paintOf(r.color);
			if (runPaint) span.style.color = cssColor(runPaint);
			if (r.bold) span.style.fontWeight = "bold";
			if (r.italic) span.style.fontStyle = "italic";
			const deco = [r.underline ? "underline" : "", r.strike ? "line-through" : ""].filter(Boolean).join(" ");
			if (deco) span.style.textDecoration = deco;
			para.append(span);
		}
		if (p.runs.length > 0 && p.runs[p.runs.length - 1].br) para.append(document.createElement("br"));
		paras.push(para);
	}
	el.replaceChildren(...paras);
}
/** 時計回りの回転角（度。0〜360 未満）。`rot` は 1/60000 度で、無い・有限でない値は 0 */
function rotationDeg(rot) {
	if (rot === void 0 || !Number.isFinite(rot)) return 0;
	return num((rot / 6e4 % 360 + 360) % 360);
}
/** SVG の要素を作り、属性を入れる（値は数値・列挙値・数値だけの `d` のみ） */
function svgEl(tag, attrs) {
	const e = document.createElementNS(SVG_NS, tag);
	for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
	return e;
}
/**
* 線 1 本の `<svg>` を作る（外接矩形に置き、矢じりは `<path>`・`<ellipse>`。`<marker>` は使わない）。
*
* @param color 線の色（lib の `DrawingLine.color`）。矢じりの塗り・線も同じ。無い・不正なら style を入れず、属性の `currentColor`（暫定のテーマ変数）のまま
*/
function buildLine(line, rect, color) {
	const paint = paintOf(color);
	const svg = svgEl("svg", {
		width: rect.w,
		height: rect.h,
		viewBox: `0 0 ${num(rect.w)} ${num(rect.h)}`,
		"aria-hidden": "true",
		focusable: "false"
	});
	svg.setAttribute("class", "xs-figure xs-line");
	svg.style.left = `${rect.x}px`;
	svg.style.top = `${rect.y}px`;
	svg.style.width = `${rect.w}px`;
	svg.style.height = `${rect.h}px`;
	const stroke = num(line.strokeWidth);
	const attrs = {
		d: line.d,
		fill: "none",
		stroke: "currentColor",
		"stroke-width": stroke,
		"stroke-linejoin": "round",
		"stroke-linecap": "butt"
	};
	const dash = line.dash.filter((n) => Number.isFinite(n) && n >= 0).map(num);
	if (dash.length > 0) attrs["stroke-dasharray"] = dash.join(" ");
	const main = svgEl("path", attrs);
	setPaint(main, "stroke", paint);
	svg.append(main);
	for (const a of line.heads) if (a.kind === "polygon") {
		const head = svgEl("path", {
			d: arrowPolygonData(a),
			fill: a.filled ? "currentColor" : "none",
			stroke: a.filled ? "none" : "currentColor",
			"stroke-width": stroke,
			"stroke-linejoin": "round",
			"stroke-linecap": "round"
		});
		setPaint(head, a.filled ? "fill" : "stroke", paint);
		svg.append(head);
	} else {
		const oval = svgEl("ellipse", {
			cx: num(a.cx),
			cy: num(a.cy),
			rx: num(a.rx),
			ry: num(a.ry),
			transform: `rotate(${num(a.angle)} ${num(a.cx)} ${num(a.cy)})`,
			fill: "currentColor",
			stroke: "none"
		});
		setPaint(oval, "fill", paint);
		svg.append(oval);
	}
	return svg;
}
/**
* ダーク表示の下地の `<svg>` を作る（ISSUE-96）。外接矩形（図の位置）から四方に `pad` だけ広げて置き、`viewBox` を `-pad` からにして中の座標は図と同じにする。
* クラス `xs-backdrop`：`z-index: -1`・色・不透明度は CSS（テーマ変数）が決める。図の要素ではない（`xs-figure` を付けない）。
*/
function backdropSvg(rect, pad) {
	const w = num(rect.w + 2 * pad);
	const h = num(rect.h + 2 * pad);
	const svg = svgEl("svg", {
		width: w,
		height: h,
		viewBox: `${-pad} ${-pad} ${w} ${h}`,
		"aria-hidden": "true",
		focusable: "false"
	});
	svg.setAttribute("class", "xs-backdrop");
	svg.style.left = `${num(rect.x - pad)}px`;
	svg.style.top = `${num(rect.y - pad)}px`;
	svg.style.width = `${w}px`;
	svg.style.height = `${h}px`;
	return svg;
}
/** 下地の要素（`path`・`ellipse`・`rect`）。`fill` なら形を白く塗り、`strokeWidth` があれば線の縁取り（実線・角と端は丸）。色は付けない（クラスと CSS のテーマ変数） */
function backdropPart(tag, attrs, fill, strokeWidth) {
	const all = { ...attrs };
	const classes = ["xs-backdrop-part"];
	if (fill) classes.push("is-fill");
	if (strokeWidth !== void 0) {
		classes.push("is-stroke");
		all["stroke-width"] = strokeWidth;
		all["stroke-linejoin"] = "round";
		all["stroke-linecap"] = "round";
	}
	all.class = classes.join(" ");
	return svgEl(tag, all);
}
/** 線・矢印の下地：線の path に沿った白い縁取り（太さ＋片側 3px。実線）。矢じりは形を白で塗り、同じ縁取りを付ける（開いた V は縁取りだけ） */
function buildLineBackdrop(line, rect) {
	const svg = backdropSvg(rect, 3);
	const border = num(line.strokeWidth + 6);
	svg.append(backdropPart("path", { d: line.d }, false, border));
	for (const a of line.heads) if (a.kind === "polygon") svg.append(backdropPart("path", { d: arrowPolygonData(a) }, a.filled, border));
	else svg.append(backdropPart("ellipse", {
		cx: num(a.cx),
		cy: num(a.cy),
		rx: num(a.rx),
		ry: num(a.ry),
		transform: `rotate(${num(a.angle)} ${num(a.cx)} ${num(a.cy)})`
	}, true, border));
	return svg;
}
/** 面の濃淡 → クラス（色が無いときだけ CSS が文字色・背景色の薄い膜にする暫定の色。色があれば style の色が勝つ） */
var SHADE_CLASS = {
	darken: "is-darken",
	darkenLess: "is-darken-less",
	lighten: "is-lighten",
	lightenLess: "is-lighten-less"
};
/**
* 図形の形の `<svg>` を作る（外接矩形いっぱい。path 1 つにつき `<path>` 1 つ）。
* 色は Excel の色（`drawing.fillColor`・`outlineColor`）を style の fill・stroke に入れる。面の濃淡は塗りの色に lib の `applyFillMode` を掛けた色。
* 色が無い（省略・不正）ときは style を入れず、クラスと CSS の暫定の色（テーマ変数）のまま。
*/
function buildShapeSvg(shape, rect, drawing) {
	const fillColor = paintOf(drawing.fillColor) ? drawing.fillColor : void 0;
	const outlinePaint = paintOf(drawing.outlineColor);
	const svg = svgEl("svg", {
		width: rect.w,
		height: rect.h,
		viewBox: `0 0 ${num(rect.w)} ${num(rect.h)}`,
		"aria-hidden": "true",
		focusable: "false"
	});
	svg.setAttribute("class", "xs-shape-svg");
	const stroke = num(shape.strokeWidth);
	const dash = shape.dash.filter((n) => Number.isFinite(n) && n >= 0).map(num);
	for (const f of shape.faces) {
		const attrs = { d: f.d };
		const classes = ["xs-shape-path"];
		if (f.fill) {
			classes.push("has-fill");
			if (f.shade) classes.push(SHADE_CLASS[f.shade]);
		}
		if (f.stroke) {
			classes.push("has-stroke");
			attrs["stroke-width"] = stroke;
			attrs["stroke-linejoin"] = "miter";
			attrs["stroke-miterlimit"] = 4;
			if (dash.length > 0) attrs["stroke-dasharray"] = dash.join(" ");
		}
		attrs.class = classes.join(" ");
		const path = svgEl("path", attrs);
		if (f.fill && fillColor) setPaint(path, "fill", paintOf(f.shade ? applyFillMode(fillColor, f.shade) : fillColor));
		if (f.stroke) setPaint(path, "stroke", outlinePaint);
		svg.append(path);
	}
	return svg;
}
/** 図形の下地の `<svg>`（ISSUE-96）：形の塗りの path は白で塗り、枠線は線の縁取り（太さ＋片側 3px・実線）。何も無ければ null */
function buildShapeBackdrop(shape, rect) {
	if (shape.backdrop.length === 0) return null;
	const svg = backdropSvg(rect, 3);
	const border = num(shape.strokeWidth + 6);
	for (const b of shape.backdrop) svg.append(backdropPart("path", { d: b.d }, b.fill, b.stroke ? border : void 0));
	return svg;
}
/**
* 図が Excel の色を 1 つでも持つか（塗り・枠線・線・文字のランの色）。持たない図（色を省略した SheetDrawing）は暫定のテーマ変数
* （`--xs-fg` など）で描かれ、白い下地に乗ると読めないため、下地を敷かない（ISSUE-96）。
*/
function hasAnyColor(drawing) {
	return !!drawing.fillColor || !!drawing.outlineColor || !!drawing.line?.color || !!drawing.text?.paragraphs.some((p) => p.runs.some((r) => !!r.color));
}
/** 代替の矩形の文字枠の下地（ISSUE-96）：回転前の矩形（`r`）の位置に置き、図の要素と同じく中心で回す。余白は付けない（縁取りの無い塗りだけ） */
function buildFrameBackdrop(r, deg) {
	const svg = backdropSvg(r, 0);
	if (deg !== 0) svg.style.transform = `rotate(${deg}deg)`;
	svg.append(backdropPart("rect", {
		width: num(r.w),
		height: num(r.h)
	}, true));
	return svg;
}
/**
* 形で描く図形の要素を作る（ADR-0028 §決定 5）：`div.xs-figure`（外接矩形）＞ `<svg>` ＋ `div.xs-shape-box`（回転前の矩形）＞ `div.xs-shape-text`（文字領域）。
* 文字は図形の矩形（box）で切り、文字領域では切らない。回転は box ごと中心で掛け、反転は文字領域の位置（`textRect`）に反映済みで文字は鏡写しにしない。
* 上下反転（`flipV`・`flipHV`）の図形は、Excel と同じく文字を 180° 回す（文字領域の中心まわり。領域の位置は鏡写し済み。左右反転は文字を回さない。ISSUE-93）。
*/
function buildShape(shape, rect, drawing) {
	const el = document.createElement("div");
	el.className = "xs-figure";
	el.style.left = `${rect.x}px`;
	el.style.top = `${rect.y}px`;
	el.style.width = `${rect.w}px`;
	el.style.height = `${rect.h}px`;
	if (shape.faces.length > 0) el.append(buildShapeSvg(shape, rect, drawing));
	if (shapeHasText(drawing)) {
		const box = document.createElement("div");
		box.className = "xs-shape-box";
		box.style.left = `${num(shape.box.x - rect.x)}px`;
		box.style.top = `${num(shape.box.y - rect.y)}px`;
		box.style.width = `${num(shape.box.w)}px`;
		box.style.height = `${num(shape.box.h)}px`;
		const deg = rotationDeg(drawing.rot);
		if (deg !== 0) box.style.transform = `rotate(${deg}deg)`;
		const tr = shape.textRect;
		const text = document.createElement("div");
		text.className = "xs-shape-text";
		text.style.left = `${num(tr.x)}px`;
		text.style.top = `${num(tr.y)}px`;
		text.style.width = `${num(tr.w)}px`;
		text.style.height = `${num(tr.h)}px`;
		if (drawing.flipV === true) text.style.transform = "rotate(180deg)";
		fillText(text, drawing.text, tr.w, tr.h);
		box.append(text);
		el.append(box);
	}
	return el;
}
/**
* 画像レイヤーを作る。
*
* @param deps テスト用に差し替える依存（blob URL の作成・解放）。通常は省略（`URL.createObjectURL`・`revokeObjectURL`）
*/
function createImageLayer(deps = {}) {
	const messages = deps.messages ?? MESSAGES_JA;
	const element = document.createElement("div");
	element.className = "xs-layer";
	let sheet = null;
	let store = null;
	/** 今 DOM にある図（図の位置 → 要素）。重ね順は `order`（図の位置の昇順） */
	let mounted = /* @__PURE__ */ new Map();
	let order = [];
	let disposed = false;
	/** DOM の図をすべて外す */
	const unmountAll = () => {
		mounted = /* @__PURE__ */ new Map();
		order = [];
		element.replaceChildren();
	};
	/** 画像の読み込み結果を図へ反映する（`<img>` か破線枠） */
	const applyEntry = (fig, entry) => {
		if (entry.ok) {
			const img = document.createElement("img");
			img.decoding = "async";
			img.alt = fig.drawing.descr || fig.drawing.name;
			img.draggable = false;
			img.className = "xs-img";
			img.onerror = () => fillFrame(fig.el, messages.imageBroken);
			img.src = entry.url;
			fig.el.replaceChildren(img);
		} else fillFrame(fig.el, frameLabel(fig.drawing, entry, messages));
	};
	/** 図 1 つの要素を作る。画像は読み込み済みなら反映し、未読なら読み込みを始める（読み込み中は空） */
	const mount = (box) => {
		const drawing = sheet.drawings[box.index];
		if (drawing.kind === "line" && box.line) return {
			box,
			drawing,
			el: buildLine(box.line, box.rect, drawing.line?.color),
			backdrop: hasAnyColor(drawing) ? buildLineBackdrop(box.line, box.rect) : void 0
		};
		if (box.shape) return {
			box,
			drawing,
			el: buildShape(box.shape, box.rect, drawing),
			backdrop: hasAnyColor(drawing) ? buildShapeBackdrop(box.shape, box.rect) ?? void 0 : void 0
		};
		const el = document.createElement("div");
		el.className = "xs-figure";
		const fig = {
			box,
			drawing,
			el
		};
		if (hasShapeText(drawing)) {
			const r = drawing.group ? { ...box.rect } : preRotationRect(box.rect, drawing.rot, xfrmExtPx(drawing));
			el.style.left = `${r.x}px`;
			el.style.top = `${r.y}px`;
			el.style.width = `${r.w}px`;
			el.style.height = `${r.h}px`;
			el.classList.add("xs-shape-text");
			el.classList.toggle("has-outline", drawing.hasOutline === true);
			el.classList.toggle("has-fill", drawing.hasFill === true);
			const deg = rotationDeg(drawing.rot);
			if (deg !== 0) el.style.transform = `rotate(${deg}deg)`;
			const fillPaint = drawing.hasFill === true ? paintOf(drawing.fillColor) : null;
			if (fillPaint) el.style.backgroundColor = cssColor(fillPaint);
			const outlinePaint = drawing.hasOutline === true ? paintOf(drawing.outlineColor) : null;
			if (outlinePaint) el.style.outlineColor = cssColor(outlinePaint);
			fillText(el, drawing.text, r.w, r.h);
			if (hasAnyColor(drawing) && (filledShown(drawing) || textShown(drawing))) fig.backdrop = buildFrameBackdrop(r, deg);
			return fig;
		}
		el.style.left = `${box.rect.x}px`;
		el.style.top = `${box.rect.y}px`;
		el.style.width = `${box.rect.w}px`;
		el.style.height = `${box.rect.h}px`;
		if (drawing.kind !== "image") fillFrame(el, frameLabel(drawing, void 0, messages));
		else if (drawing.mediaPath === null || !store) fillFrame(el, messages.imageBroken);
		else {
			const path = drawing.mediaPath;
			const known = store.peek(path);
			if (known) applyEntry(fig, known);
			else {
				const current = store;
				current.load(path).then((entry) => {
					if (!entry || disposed || store !== current) return;
					if (mounted.get(box.index) === fig) applyEntry(fig, entry);
				});
			}
		}
		return fig;
	};
	return {
		element,
		setSheet(next) {
			if (disposed) return;
			store?.dispose();
			store = null;
			unmountAll();
			sheet = next;
			if (next) store = createImageStore({
				read: next.readImage,
				createUrl: deps.createUrl,
				revokeUrl: deps.revokeUrl,
				onOverBudget: next.onOverBudget
			});
		},
		place(left, top, width, height) {
			element.style.left = `${left}px`;
			element.style.top = `${top}px`;
			element.style.width = `${Math.max(0, width)}px`;
			element.style.height = `${Math.max(0, height)}px`;
		},
		update(view) {
			if (disposed || !sheet) return;
			const wanted = pickVisible(sheet.boxes, view);
			if (wanted.length === order.length && wanted.every((b, i) => b.index === order[i])) return;
			const next = /* @__PURE__ */ new Map();
			for (const box of wanted) next.set(box.index, mounted.get(box.index) ?? mount(box));
			mounted = next;
			order = wanted.map((b) => b.index);
			const nodes = [];
			for (const b of wanted) {
				const fig = next.get(b.index);
				if (fig.backdrop) nodes.push(fig.backdrop);
				nodes.push(fig.el);
			}
			element.replaceChildren(...nodes);
		},
		dispose() {
			if (disposed) return;
			disposed = true;
			store?.dispose();
			store = null;
			sheet = null;
			unmountAll();
		}
	};
}
// src/render/sheet-view.ts
/**
* `.xlsx` の 1 シートの表示（InterHub `spreadsheet.ts` の `renderXlsxView`・`showSheet`・`draw`・`dispose` の .xlsx の部分。M2-3）。
* - 見えている行・列（と前後の余分）だけを DOM にする行・列の仮想化グリッド。表は Excel の列幅・行高（geometry）で固定する。
*   罫線は表の上の SVG 層、図（画像・図形・線）は表の上の図の層に、Excel と同じセルの位置で重ねる。
* - 状態はすべてこの表示（インスタンス）が持つ。InterHub のモジュール変数（`renderGeneration`・`disposeXlsxView`）は無く、
*   同じページに複数の表示を置ける。書式の CSS の連番（`data-xs-fmt`）だけはモジュールで 1 つの連番。
* - 表示の根は自前の `div.xs-sheet`（使う側の `container` には属性を付けない）。子は通知の行（`div.xs-notice`）・
*   スクロール枠（`div.xs-body`。仮想化の窓）・書式の `<style>`。枠の大きさは使う側が決める（`container` の高さ）。
* - 書式の CSS はブックで 1 回だけ作り、シートを切り替えても作り直さない（`renderWorkbook` が同じ表示でシートを替える）。
* - タブは持たない（`renderWorkbook` の仕事）。ブックも閉じない（開いた側が閉じる）。
* ユーザーのファイル内容（セル値）は信頼できない入力として扱い、`innerHTML` は使わず `textContent` だけで DOM を組む。
*/
/** 見えている範囲の上下に余分に描く行数（xrag-Viewer の値を踏襲）。 */
var OVERSCAN = 40;
/** 通知の理由 → 文言のキー */
var NOTICE_KEYS = {
	"part-too-large": "noticeTooLarge",
	broken: "noticeBroken",
	"too-many": "noticeTooMany",
	"text-truncated": "noticeTextTruncated",
	"shape-simplified": "noticeShapeSimplified",
	"over-budget": "noticeOverBudget",
	"geom-failed": "libLoadFailed"
};
/** 書式 CSS の表示ごとの連番（`data-xs-fmt`）。表示を作るたびに 1 増やす（ファイル名・パスは使わない。ADR-0030 §決定 1） */
var formatSeq = 0;
/**
* 読み取りの例外を種類ごとの文言にする（ADR-0025 決定 6）。
* not-zip・broken やそれ以外の例外は「解析に失敗しました」に詳細を添える。
* 詳細（XlsxError のメッセージ）にはセル値・シート名を含まない。
*/
function xlsxErrorMessage(e, messages) {
	if (e instanceof XlsxError) switch (e.kind) {
		case "encrypted": return messages.encrypted;
		case "too-large": return messages.tooLarge;
		case "unsupported": return messages.unsupported;
	}
	const detail = e instanceof Error ? e.message : String(e);
	return formatMessage(messages.parseFailed, { detail });
}
/** データ無しの表示（空のシート）。 */
function emptyView(message) {
	const div = document.createElement("div");
	div.className = "xs-empty";
	div.textContent = message;
	return div;
}
/** 警告（シートを開けない）の要素を作る。 */
function bannerElement(message) {
	const banner = document.createElement("div");
	banner.className = "xs-banner";
	banner.textContent = message;
	return banner;
}
/** 形を評価する図（形の表が要る図）があるか */
function needsGeom(found) {
	return found.drawings.some((d) => d.geom !== void 0);
}
/**
* 1 シートの表示を作り、根を `container` に取り付ける（シートは {@link SheetViewInstance.showSheet} で開く）。
* 保持するのは表示中の 1 シートだけ。シートを替えると前のシートを放す（ADR-0025 決定 4）。
*
* @param container 表示を置く要素（高さが決まっていること）
* @param book 開いたブック（閉じるのは呼び出し側）
*/
function createSheetView(container, book, options) {
	const { messages } = options;
	const root = document.createElement("div");
	root.className = "xs-sheet";
	const fmtSeq = ++formatSeq;
	root.setAttribute("data-xs-fmt", String(fmtSeq));
	const fmtStyle = document.createElement("style");
	if (options.nonce !== void 0) fmtStyle.setAttribute("nonce", options.nonce);
	let fmtReady = false;
	/** スクロール枠（仮想化の窓）。大きさは CSS（`container` の高さ）で決まり、ここでは測るだけ */
	const body = document.createElement("div");
	body.className = "xs-body";
	let sheet = null;
	let activeIndex = null;
	/** 表示中のシートの寸法（Excel の列幅・行高を px にしたもの）。シートを開くたび・表の範囲が変わるたびに作り直す */
	let geo = null;
	/** 表の行数・列数（データより広げて、スクロール枠を埋める）。`geo` はこの範囲で作る */
	let extent = {
		rows: 0,
		cols: 0
	};
	/** 描いた行・列の範囲（描き直しの判定に使う） */
	let drawn = null;
	/** 新しい showSheet・destroy が来たら、前のシート読み込みの結果を捨てるための番号。 */
	let openSeq = 0;
	let frame = 0;
	let initFrame = 0;
	let closed = false;
	const wrapper = document.createElement("div");
	wrapper.className = "xs-wrap";
	const layer = createImageLayer({ messages });
	const borderLayer = createBorderLayer();
	wrapper.append(borderLayer.element, layer.element);
	/** wrapper の中の今の表（無ければ null） */
	let tableEl = null;
	/** 表示中のシートの図の配置（A1 の左上基準 px。大きさ 0 の図は含まない） */
	let boxes = [];
	/** 図を収めるのに必要な表の行数・列数（表を図のある範囲まで広げる） */
	let drawingMin = {
		rows: 0,
		cols: 0
	};
	/** 図の読み取り・上限超過をまとめて知らせる 1 行（無ければ hidden） */
	const noticeBar = document.createElement("div");
	noticeBar.className = "xs-notice";
	noticeBar.hidden = true;
	/** 表示中のシートで知らせる理由 */
	const notices = /* @__PURE__ */ new Set();
	/** 通知の行の文言を notices に合わせる（理由の順に 1 行で並べる） */
	const renderNotices = () => {
		const text = [...notices].map((n) => messages[NOTICE_KEYS[n]]).join(" ");
		noticeBar.textContent = text;
		noticeBar.hidden = text === "";
	};
	/** 表も図も出さず、メッセージ（空・エラー）だけを本文枠に出す。前のシートの blob URL はここで revoke する */
	const showMessage = (el) => {
		layer.setSheet(null);
		boxes = [];
		drawingMin = {
			rows: 0,
			cols: 0
		};
		tableEl = null;
		borderLayer.clear();
		wrapper.replaceChildren(borderLayer.element, layer.element);
		body.replaceChildren(el);
	};
	/**
	* 表の範囲（行数・列数）を、スクロール枠の大きさに合わせて決め直す。データより広げて空セルを足し、画面を埋める
	* （Excel のシートのように、右端の列の文字もはみ出せる）。範囲が変わったら geometry を作り直して true を返す。
	* 図のある範囲まで広げる（ISSUE-86）ため、図の最下行・最右列を gridExtent の minRows・minCols に渡す。
	* 枠の大きさは使う側が決める（InterHub はウィンドウ高に合わせた max-height も見ていたが、ここでは clientWidth・clientHeight だけ）。
	*/
	const refreshExtent = (sh) => {
		const next = gridExtent(sh.layout, sh.rowCount, sh.colCount, body.clientWidth, body.clientHeight, drawingMin.rows, drawingMin.cols);
		if (geo && next.rows === extent.rows && next.cols === extent.cols) return false;
		extent = next;
		geo = createGeometry(sh.layout, next.rows, next.cols);
		return true;
	};
	/** 行 a〜b の読み出し。データより下（足した空の行）は空の行にする。 */
	const readRows = (sh, a, b) => {
		const rows = a <= sh.rowCount ? sh.getRows(a, Math.min(b, sh.rowCount)) : [];
		for (let r = Math.max(a, sh.rowCount + 1); r <= b; r++) rows.push({
			r,
			cells: []
		});
		return rows;
	};
	/**
	* 見えている行（と上下の余分）を描く。force でなければ、描いた範囲に見えている行が収まる間は描き直さない。
	* 行の位置・上下の余白は geometry（Excel の寸法）から出し、描いた行の高さは実測しない。
	* 非表示の行は DOM に出さず、前後の余分も見える行の数で数える（フィルターで大半を畳んだシートでも描く行は一定）。
	*/
	const draw = (force = false) => {
		const sh = sheet;
		if (!sh || closed) return;
		const extended = refreshExtent(sh);
		const g = geo;
		if (!g) return;
		const range = visibleRowRange(g, body.scrollTop, body.clientHeight, extent.rows, OVERSCAN);
		const d = drawn;
		const originX = rowNumWidth(extent.rows);
		const cols = visibleColRange(g, body.scrollLeft, body.clientWidth, originX, extent.cols);
		if (force || extended || !d || range.top < d.first || range.bottom > d.last || cols.left < d.firstCol || cols.right > d.lastCol) {
			const read = collectVisibleRows(g, range.first, range.last, extent.rows, (a, b) => readRows(sh, a, b));
			const rows = sh.formats.length > 0 ? applyMarginFormats(read, cols.first, cols.last, sh.rowCount, sh.colCount, (r) => sh.rowFormat(r), (c) => sh.colFormat(c)) : read;
			const table = buildGridTable(colName, extent.cols, rows, g, extent.rows, sh.layout.mdw, cols, sh.formats);
			if (wrapper.parentNode !== body) body.replaceChildren(wrapper);
			if (tableEl) tableEl.replaceWith(table);
			else wrapper.prepend(table);
			tableEl = table;
			drawn = {
				first: range.first,
				last: range.last,
				firstCol: cols.first,
				lastCol: cols.last
			};
			layer.place(originX, 24, g.totalWidth, g.totalHeight);
			if (sh.borders.length > 0) borderLayer.render(buildBorderPaths(rows, sh.borders, g, {
				first: cols.first,
				last: cols.last
			}), originX, 24);
			else borderLayer.clear();
		}
		layer.update(viewWindow(body.scrollLeft, body.scrollTop, body.clientWidth, body.clientHeight, originX, 24));
	};
	/** スクロール・枠の大きさの変化を次のフレームの描画 1 回にまとめる */
	const schedule = () => {
		if (frame || closed) return;
		frame = requestAnimationFrame(() => {
			frame = 0;
			draw();
		});
	};
	const observer = typeof ResizeObserver === "function" ? new ResizeObserver(schedule) : null;
	const showSheet = async (index) => {
		if (closed) return;
		if (!book.sheets.some((s) => s.index === index)) throw new RangeError("シートの番号が範囲外");
		const seq = ++openSeq;
		const isStale = () => seq !== openSeq || closed;
		let opened;
		try {
			opened = await book.openSheet(index);
		} catch (e) {
			if (isStale()) return;
			sheet = null;
			geo = null;
			activeIndex = null;
			drawn = null;
			notices.clear();
			renderNotices();
			showMessage(bannerElement(xlsxErrorMessage(e, messages)));
			return;
		}
		let found = {
			drawings: [],
			notices: []
		};
		let geomFailed = false;
		if (options.drawings) {
			try {
				found = await book.openDrawings(index);
			} catch {
				found = {
					drawings: [],
					notices: ["broken"]
				};
			}
			if (needsGeom(found)) try {
				await loadGeom();
			} catch {
				geomFailed = true;
			}
		}
		if (isStale()) return;
		if (!fmtReady) {
			fmtReady = true;
			fmtStyle.textContent = buildFormatCss(opened.formats, fmtSeq);
		}
		activeIndex = index;
		drawn = null;
		body.scrollTop = 0;
		body.scrollLeft = 0;
		notices.clear();
		found.notices.forEach((n) => notices.add(n));
		if (geomFailed) notices.add("geom-failed");
		renderNotices();
		const bigGeo = createGeometry(opened.layout, MAX_GRID_ROWS, MAX_GRID_COLS);
		const drawings = found.drawings;
		const placed = layoutDrawings(bigGeo, drawings);
		if (opened.rowCount === 0 && placed.length === 0) {
			sheet = null;
			geo = null;
			showMessage(emptyView(messages.noData));
			return;
		}
		sheet = opened;
		geo = null;
		boxes = placed;
		drawingMin = drawingExtent(bigGeo, boxes);
		layer.setSheet(options.drawings ? {
			drawings,
			boxes,
			readImage: (path) => book.readImage(index, path),
			onOverBudget: () => {
				if (!isStale()) {
					notices.add("over-budget");
					renderNotices();
				}
			}
		} : null);
		draw(true);
		cancelAnimationFrame(initFrame);
		initFrame = requestAnimationFrame(() => {
			initFrame = 0;
			draw(true);
		});
	};
	const destroy = () => {
		if (closed) return;
		closed = true;
		openSeq++;
		layer.dispose();
		cancelAnimationFrame(frame);
		cancelAnimationFrame(initFrame);
		frame = 0;
		initFrame = 0;
		body.removeEventListener("scroll", schedule);
		observer?.disconnect();
		sheet = null;
		geo = null;
		drawn = null;
		activeIndex = null;
		borderLayer.clear();
		boxes = [];
		tableEl = null;
		notices.clear();
		wrapper.replaceChildren();
		body.replaceChildren();
		root.replaceChildren();
		root.removeAttribute("data-xs-fmt");
		root.remove();
	};
	root.append(noticeBar, body, fmtStyle);
	container.append(root);
	body.addEventListener("scroll", schedule, { passive: true });
	observer?.observe(body);
	return {
		root,
		get activeIndex() {
			return activeIndex;
		},
		showSheet,
		destroy
	};
}
// src/render/render-sheet.ts
/**
* ブックの 1 シートを `container` の中に表示する。
* `container` は高さが決まっていること（中のスクロール枠 `.xs-body` が仮想化の窓になる）。
* シートを開けない（壊れている等）ときは、例外にせず枠の中にメッセージを出す。
*
* @param container 表示を置く要素（子として `div.xs-sheet` を足す。`container` 自身には何も付けない）
* @param book `openXlsx` で開いたブック（閉じるのは呼び出し側）
* @param sheetIndex シートの番号（`SheetInfo.index`。1 始まり・非表示のシートも数える）
* @throws RangeError ブックに無い番号（DOM には何も足さない）
*/
async function renderSheet(container, book, sheetIndex, options = {}) {
	if (!book.sheets.some((s) => s.index === sheetIndex)) throw new RangeError("シートの番号が範囲外");
	const view = createSheetView(container, book, {
		messages: resolveMessages(options.messages),
		drawings: options.drawings !== false,
		nonce: options.nonce
	});
	try {
		await view.showSheet(sheetIndex);
	} catch (e) {
		view.destroy();
		throw e;
	}
	return { destroy: () => view.destroy() };
}
// src/render/sheet-tabs.ts
/**
* シートタブのボタンを作る（クリックの処理は呼び出し側で付ける）。
* 長いシート名は CSS で省略表示されるため、全文を title で見られるようにする。
* 非表示のシートは名前に ` ⌀` を足し、`is-hidden-sheet` と、title に注記（`hiddenSheet`）を付ける。
*/
function makeSheetTab(name, hidden, messages) {
	const btn = document.createElement("button");
	btn.type = "button";
	btn.className = "xs-sheet-tab";
	btn.textContent = name + (hidden ? " ⌀" : "");
	btn.title = name;
	if (hidden) {
		btn.classList.add("is-hidden-sheet");
		btn.title = `${name}\n${messages.hiddenSheet}`;
	}
	return btn;
}
/**
* シートタブを 1 行に並べ、入りきらない分は横スクロールさせる列を作る。
* 構造は `div.xs-sheet-strip` ＞ `[◀ button.xs-sheet-nav][div.xs-sheet-tabs（overflow-x:auto）＞ タブ…][▶ button.xs-sheet-nav]`。
* - ◀▶ は押すとスクロール領域の可視幅の約 8 割ぶん滑らかに動かす。端に着いた側は disabled、全タブが収まるときは両方隠す。
* - 状態はスクロール・リサイズ（ResizeObserver。無い環境ではガードして使わない）のたびに更新する。
* - 縦のホイール回転は、あふれているときだけ横スクロールに変換する。
*
* @param buttons シート名のタブ（{@link makeSheetTab}。クリック処理は呼び出し側で設定済みのもの）
*/
function buildSheetTabStrip(buttons, messages) {
	const element = document.createElement("div");
	element.className = "xs-sheet-strip";
	const tabs = document.createElement("div");
	tabs.className = "xs-sheet-tabs";
	tabs.append(...buttons);
	const makeNav = (text, label) => {
		const nav = document.createElement("button");
		nav.type = "button";
		nav.className = "xs-sheet-nav";
		nav.textContent = text;
		nav.title = label;
		nav.setAttribute("aria-label", label);
		return nav;
	};
	const prev = makeNav("◀", messages.tabsPrev);
	const next = makeNav("▶", messages.tabsNext);
	element.append(prev, tabs, next);
	/** スクロール位置を動かす（smooth 指定は scrollTo 未実装の環境では即時に落とす）。 */
	const scrollTabsTo = (left, smooth) => {
		if (typeof tabs.scrollTo === "function") tabs.scrollTo({
			left,
			behavior: smooth ? "smooth" : "auto"
		});
		else tabs.scrollLeft = left;
	};
	/** ◀▶ の表示・disabled を現在のスクロール位置に合わせる。 */
	const updateNav = () => {
		const max = tabs.scrollWidth - tabs.clientWidth;
		const overflow = max > 1;
		prev.hidden = !overflow;
		next.hidden = !overflow;
		prev.disabled = tabs.scrollLeft <= 0;
		next.disabled = tabs.scrollLeft >= max - 1;
	};
	let pending;
	/** 選択中タブが見える範囲へ寄せる（tabs が offsetParent なので offsetLeft はスクロール領域内の座標）。 */
	const applyReveal = (btn, smooth) => {
		const left = btn.offsetLeft;
		const right = left + btn.offsetWidth;
		if (left < tabs.scrollLeft) scrollTabsTo(left, smooth);
		else if (right > tabs.scrollLeft + tabs.clientWidth) scrollTabsTo(right - tabs.clientWidth, smooth);
	};
	const reveal = (btn) => {
		if (!btn) return;
		if (tabs.clientWidth === 0) {
			pending = btn;
			return;
		}
		pending = void 0;
		applyReveal(btn, true);
	};
	const onLayout = () => {
		if (pending && tabs.clientWidth > 0) {
			const target = pending;
			pending = void 0;
			applyReveal(target, false);
		}
		updateNav();
	};
	const onPrev = () => scrollTabsTo(tabs.scrollLeft - tabs.clientWidth * .8, true);
	const onNext = () => scrollTabsTo(tabs.scrollLeft + tabs.clientWidth * .8, true);
	prev.addEventListener("click", onPrev);
	next.addEventListener("click", onNext);
	/** 縦ホイールを横スクロールへ変換する（あふれていて縦成分が主のときだけ。preventDefault のため非 passive）。 */
	const onWheel = (e) => {
		if (tabs.scrollWidth - tabs.clientWidth <= 1 || Math.abs(e.deltaY) <= Math.abs(e.deltaX)) return;
		e.preventDefault();
		tabs.scrollLeft += e.deltaY;
	};
	tabs.addEventListener("scroll", updateNav, { passive: true });
	tabs.addEventListener("wheel", onWheel, { passive: false });
	let frame = 0;
	const onResize = () => {
		if (frame === 0) frame = requestAnimationFrame(() => {
			frame = 0;
			onLayout();
		});
	};
	const observer = typeof ResizeObserver === "function" ? new ResizeObserver(onResize) : null;
	observer?.observe(tabs);
	updateNav();
	const dispose = () => {
		observer?.disconnect();
		if (frame !== 0) {
			cancelAnimationFrame(frame);
			frame = 0;
		}
		prev.removeEventListener("click", onPrev);
		next.removeEventListener("click", onNext);
		tabs.removeEventListener("scroll", updateNav);
		tabs.removeEventListener("wheel", onWheel);
		pending = void 0;
	};
	return {
		element,
		reveal,
		dispose
	};
}
// src/render/render-workbook.ts
/**
* 公開の入口 `renderWorkbook`（M2-4）：xlsx のバイト列を開き、シートタブ付きで `container` の中に表示する。
* InterHub `spreadsheet.ts` の `renderXlsxView` のうち、ブックを開く・タブ・最初のシートの選択・エラー表示の部分。
* - 表示は 1 冊につき 1 つのシートの表示（`createSheetView`）を使い回し、タブでシートを替える
*   （書式の `<style>` はブックで 1 回だけ作り、タブを替えても作り直さない）。
* - 開けない（暗号化・壊れている・大きすぎる・対応しない形式）ときは例外にせず、種類ごとの文言を `div.xs-banner` に出す。
* - `destroy()` でタブ列・シートの表示（blob URL・リスナ）を放し、ブックを閉じ、根を外す。
* シート名はユーザーのファイル内容（信頼できない入力）なので、`textContent` と `title` にだけ入れる。
*/
/** メッセージだけの要素（`div.xs-banner`・`div.xs-empty`） */
function messageElement(className, text) {
	const el = document.createElement("div");
	el.className = className;
	el.textContent = text;
	return el;
}
/** シートを持たない表示（開けない・シートが 0 枚）。destroy は根を外すだけ */
function messageView(root) {
	return {
		sheets: [],
		showSheet: () => Promise.reject(/* @__PURE__ */ new RangeError("シートの番号が範囲外")),
		destroy: () => {
			root.replaceChildren();
			root.remove();
		}
	};
}
/**
* xlsx を開き、シートタブ付きで `container` の中に表示する。
* 最初に出すシートは、最初の表示（非表示でない）シート。全部非表示なら先頭。
* `container` は高さが決まっていること（タブ列の下のスクロール枠 `.xs-body` が仮想化の窓になる）。
*
* @param container 表示を置く要素（子として `div.xs-workbook` を足す。`container` 自身には何も付けない）
* @param data xlsx のバイト列
*/
async function renderWorkbook(container, data, options = {}) {
	const messages = resolveMessages(options.messages);
	const root = document.createElement("div");
	root.className = "xs-workbook";
	container.append(root);
	let book;
	try {
		book = await openXlsx(data, { limits: options.limits });
	} catch (e) {
		root.append(messageElement("xs-banner", xlsxErrorMessage(e, messages)));
		return messageView(root);
	}
	const sheets = book.sheets;
	if (sheets.length === 0) {
		book.close();
		root.append(messageElement("xs-empty", messages.noSheets));
		return messageView(root);
	}
	let closed = false;
	const buttons = sheets.map((info) => makeSheetTab(info.name, info.hidden, messages));
	const strip = buildSheetTabStrip(buttons, messages);
	root.append(strip.element);
	const view = createSheetView(root, book, {
		messages,
		drawings: options.drawings !== false,
		nonce: options.nonce
	});
	/** 指定のシートを開いて表示する（タブ操作・初期表示）。 */
	const showSheet = async (index) => {
		if (closed) return;
		const at = sheets.findIndex((s) => s.index === index);
		if (at < 0) throw new RangeError("シートの番号が範囲外");
		if (view.activeIndex === index) return;
		buttons.forEach((b, bi) => b.classList.toggle("is-active", bi === at));
		strip.reveal(buttons[at]);
		await view.showSheet(index);
	};
	const onClicks = buttons.map((btn, bi) => {
		const onClick = () => {
			showSheet(sheets[bi].index);
		};
		btn.addEventListener("click", onClick);
		return onClick;
	});
	const destroy = () => {
		if (closed) return;
		closed = true;
		buttons.forEach((btn, bi) => btn.removeEventListener("click", onClicks[bi]));
		strip.dispose();
		view.destroy();
		book.close();
		root.replaceChildren();
		root.remove();
	};
	try {
		await showSheet((sheets.find((s) => !s.hidden) ?? sheets[0]).index);
	} catch (e) {
		destroy();
		throw e;
	}
	return {
		sheets,
		showSheet,
		destroy
	};
}
export { MESSAGES_KO as a, MESSAGES_JA as i, renderSheet as n, resolveMessages as o, MESSAGES_EN as r, renderWorkbook as t };
