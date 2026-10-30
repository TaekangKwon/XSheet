import { t as GUIDE_ARITY } from "./defs-BMhuSHU6.js";
import { C as parseCellRef, D as MAX_DRAWING_TEXT, E as MAX_COLS, O as MAX_ROWS, S as parseAttrs, T as MAX_CELL_TEXT, _ as findElement, a as DEFAULT_THEME, b as isTagDelimiter, d as NS, f as attrByLocalName, g as decodeText, h as decodeEntities, i as resolveColor, k as MAX_ROW_PT, l as XlsxError, m as colOfRef, n as mixColors, o as loadTheme, r as mixColorsLinear, s as parseHex6, u as emptyLayout, v as findOpenTag, w as DEFAULT_LIMITS, x as looksStrict, y as forEachElement } from "./color-sme5jHOH.js";
import * as ssfModule from "ssf";
import JSZip from "jszip";
// src/core/cell-color.ts
/** 旧 64 色の表（ECMA-376 §18.8.27。`<indexedColors>` が無いブックの `indexed` 0〜63） */
var INDEXED_COLORS = [
	0,
	16777215,
	16711680,
	65280,
	255,
	16776960,
	16711935,
	65535,
	0,
	16777215,
	16711680,
	65280,
	255,
	16776960,
	16711935,
	65535,
	8388608,
	32768,
	128,
	8421376,
	8388736,
	32896,
	12632256,
	8421504,
	10066431,
	10040166,
	16777164,
	13434879,
	6684774,
	16744576,
	26316,
	13421823,
	128,
	16711935,
	16776960,
	65535,
	8388736,
	8388608,
	32896,
	255,
	52479,
	13434879,
	13434828,
	16777113,
	10079487,
	16751052,
	13408767,
	16764057,
	3368703,
	3394764,
	10079232,
	16763904,
	16750848,
	16737792,
	6710937,
	9868950,
	13158,
	3381606,
	13056,
	3355392,
	10040064,
	10040166,
	3355545,
	3355443
];
/** `indexed` の特別な値：64＝システムの前景色（自動）、65＝システムの背景色（白） */
var INDEXED_FOREGROUND = 64;
var INDEXED_BACKGROUND = 65;
/** `<indexedColors>` の件数の上限（旧パレットは 64 色。ADR-0029 §決定 9） */
var MAX_INDEXED = INDEXED_COLORS.length;
/** `theme` の番号（SpreadsheetML の順）→ テーマの位置。0 と 1・2 と 3 が THEME_SLOTS（dk1・lt1・dk2・lt2）と入れ替わる */
var THEME_BY_INDEX = [
	"lt1",
	"dk1",
	"lt2",
	"dk2",
	"accent1",
	"accent2",
	"accent3",
	"accent4",
	"accent5",
	"accent6",
	"hlink",
	"folHlink"
];
/** 属性の値を整数にする（空・数でない・小数は null） */
function intAttr$3(v) {
	if (v === void 0 || v.trim() === "") return null;
	const n = Number(v);
	return Number.isInteger(n) ? n : null;
}
/** `rgb` 属性（ARGB 8 桁か RGB 6 桁の 16 進）を整数にする。先頭のアルファは無視。不正は null */
function parseRgbAttr(raw) {
	if (raw.length === 8 && parseHex6(raw.slice(0, 6)) !== null) return parseHex6(raw.slice(2));
	if (raw.length === 6) return parseHex6(raw);
	return null;
}
/**
* styles.xml の `<colors><indexedColors>` を読み、64 色の表を返す（既定の表を、書いてある位置だけ上書きする）。
* 無い・空なら null（既定の表を使う）。`<rgbColor rgb>` は 64 件まで。不正な値はその位置だけ既定のまま。
*/
function parseIndexedColors(xml) {
	const block = findElement(xml, "indexedColors");
	if (block === null || block.body === null) return null;
	const palette = INDEXED_COLORS.slice();
	let n = 0;
	forEachElement(block.body, "rgbColor", (open) => {
		if (n >= MAX_INDEXED) return;
		const raw = parseAttrs(open).rgb;
		const rgb = raw === void 0 ? null : parseRgbAttr(raw);
		if (rgb !== null) palette[n] = rgb;
		n++;
	});
	return n === 0 ? null : palette;
}
/** RGB の整数 → HLS（H・L・S とも 0〜1） */
function rgbToHls(rgb) {
	const r = (rgb >> 16 & 255) / 255;
	const g = (rgb >> 8 & 255) / 255;
	const b = (rgb & 255) / 255;
	const max = Math.max(r, g, b);
	const min = Math.min(r, g, b);
	const l = (max + min) / 2;
	if (max === min) return [
		0,
		l,
		0
	];
	const d = max - min;
	const s = l > .5 ? d / (2 - max - min) : d / (max + min);
	let h;
	if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
	else if (max === g) h = (b - r) / d + 2;
	else h = (r - g) / d + 4;
	return [
		h / 6,
		l,
		s
	];
}
/** HLS（0〜1）→ RGB の整数（各成分は四捨五入） */
function hlsToRgb(h, l, s) {
	const channel = (t0, p, q) => {
		let t = t0;
		if (t < 0) t += 1;
		if (t > 1) t -= 1;
		if (t < 1 / 6) return p + (q - p) * 6 * t;
		if (t < 1 / 2) return q;
		if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
		return p;
	};
	let r = l;
	let g = l;
	let b = l;
	if (s !== 0) {
		const q = l < .5 ? l * (1 + s) : l + s - l * s;
		const p = 2 * l - q;
		r = channel(h + 1 / 3, p, q);
		g = channel(h, p, q);
		b = channel(h - 1 / 3, p, q);
	}
	const to8 = (v) => Math.min(255, Math.max(0, Math.round(v * 255)));
	return to8(r) << 16 | to8(g) << 8 | to8(b);
}
/**
* SpreadsheetML の `tint`（HLS の明度）を当てる。負は `L × (1 + tint)`、正は `L × (1 − tint) + tint`。
* `tint` は -1〜1 に丸める（-1 は黒・1 は白）。0 は色を変えない。DrawingML の `tint`（線形 RGB）とは別の式。
*/
function applyTint(rgb, tint) {
	const t = Math.max(-1, Math.min(1, tint));
	if (t === 0) return rgb;
	const [h, l, s] = rgbToHls(rgb);
	return hlsToRgb(h, t < 0 ? l * (1 + t) : l * (1 - t) + t, s);
}
/**
* `<color>` の属性 → 色。null は自動（`auto="1"`・属性なし・`indexed="64"`）か解けない値。
* @param attrs `parseAttrs` で読んだ `<color>` の属性（無ければ undefined＝自動）
* @param theme ブックのテーマ（`theme` の番号を解く）
* @param palette `parseIndexedColors` の結果（null なら既定の 64 色）
*/
function resolveCellColor(attrs, theme, palette) {
	if (attrs === void 0) return null;
	if (attrs.auto === "1" || attrs.auto === "true") return null;
	let rgb = null;
	if (attrs.rgb !== void 0) rgb = parseRgbAttr(attrs.rgb);
	else if (attrs.indexed !== void 0) {
		const i = intAttr$3(attrs.indexed);
		if (i === INDEXED_BACKGROUND) rgb = 16777215;
		else if (i !== null && i >= 0 && i < MAX_INDEXED && i !== INDEXED_FOREGROUND) rgb = (palette ?? INDEXED_COLORS)[i] ?? null;
	} else if (attrs.theme !== void 0) {
		const i = intAttr$3(attrs.theme);
		const slot = i === null || i < 0 ? void 0 : THEME_BY_INDEX[i];
		if (slot !== void 0) rgb = theme.get(slot) ?? null;
	}
	if (rgb === null) return null;
	if (attrs.tint !== void 0 && attrs.tint.trim() !== "") {
		const tint = Number(attrs.tint);
		if (Number.isFinite(tint)) rgb = applyTint(rgb, tint);
	}
	return {
		rgb,
		alpha: 1
	};
}
// src/core/borders.ts
/** 罫線の無い表（styles.xml が無い・壊れているとき。共有する・変更しない） */
var EMPTY_BORDER_TABLE = Object.freeze({
	borders: Object.freeze([]),
	xfBorder: /* @__PURE__ */ new Int32Array(0)
});
/** 線種の名前（`none`・未知は含めない）。`Set` なので `__proto__` などの名前でも汚染されない */
var BORDER_STYLES = /* @__PURE__ */ new Set([
	"thin",
	"hair",
	"dotted",
	"dashed",
	"dashDot",
	"dashDotDot",
	"medium",
	"mediumDashed",
	"mediumDashDot",
	"mediumDashDotDot",
	"slantDashDot",
	"thick",
	"double"
]);
/** 真偽の属性（`1`・`true` が真） */
var boolAttr$2 = (v) => v === "1" || v === "true";
/** 整数の属性（空・数でない・小数は null） */
function intAttr$2(v) {
	if (v === void 0 || v.trim() === "") return null;
	const n = Number(v);
	return Number.isInteger(n) ? n : null;
}
/**
* `<border>` の中から辺 1 つ（要素名 name）を読む。要素が無い・線種が無い・未知・`none` なら undefined。
* 色は辺の中の最初の `<color>`。無ければ null（自動）。
*/
function readEdge(body, name, theme, palette) {
	let edge;
	let seen = false;
	forEachElement(body, name, (open, inner) => {
		if (seen) return;
		seen = true;
		const style = parseAttrs(open).style;
		if (style === void 0 || !BORDER_STYLES.has(style)) return;
		const colorTag = inner === null ? null : findOpenTag(inner, "color");
		edge = {
			style,
			color: colorTag === null ? null : resolveCellColor(parseAttrs(colorTag), theme, palette)
		};
	});
	return edge;
}
/** `<border>` 1 件（開きタグと中身）→ 罫線。どの辺にも線が無ければ null */
function readBorder(open, body, theme, palette) {
	if (body === null) return null;
	const attrs = parseAttrs(open);
	const border = {
		diagonalUp: boolAttr$2(attrs.diagonalUp),
		diagonalDown: boolAttr$2(attrs.diagonalDown)
	};
	const left = readEdge(body, "left", theme, palette) ?? readEdge(body, "start", theme, palette);
	const right = readEdge(body, "right", theme, palette) ?? readEdge(body, "end", theme, palette);
	const top = readEdge(body, "top", theme, palette);
	const bottom = readEdge(body, "bottom", theme, palette);
	const diagonal = readEdge(body, "diagonal", theme, palette);
	if (left) border.left = left;
	if (right) border.right = right;
	if (top) border.top = top;
	if (bottom) border.bottom = bottom;
	if (diagonal) border.diagonal = diagonal;
	return left || right || top || bottom || diagonal ? border : null;
}
/**
* styles.xml から罫線の表を作る（ブックで 1 回）。
* @param xml styles.xml（無ければ null＝罫線なし）
* @param theme ブックのテーマ（色の `theme` を解く）
* @param limits `maxBorders` を超えた `<border>` は読まない（それを参照するセルは罫線なし）
*/
function parseBorders(xml, theme, limits) {
	if (xml === null) return EMPTY_BORDER_TABLE;
	const block = findElement(xml, "borders");
	const cellXfs = findElement(xml, "cellXfs");
	if (block === null || block.body === null || cellXfs === null || cellXfs.body === null) return EMPTY_BORDER_TABLE;
	const palette = parseIndexedColors(xml);
	const borders = [];
	const idToIndex = [];
	forEachElement(block.body, "border", (open, body) => {
		if (idToIndex.length >= limits.maxBorders) return;
		const b = readBorder(open, body, theme, palette);
		if (b === null) idToIndex.push(-1);
		else {
			idToIndex.push(borders.length);
			borders.push(b);
		}
	});
	const xfBorder = [];
	for (const m of cellXfs.body.matchAll(new RegExp(String.raw`<${NS}xf\b[^<>]*>`, "g"))) {
		const id = intAttr$2(parseAttrs(m[0]).borderId);
		xfBorder.push(id !== null && id >= 0 ? idToIndex[id] ?? -1 : -1);
	}
	return {
		borders,
		xfBorder: Int32Array.from(xfBorder)
	};
}
// src/core/cell-format.ts
/**
* styles.xml の書式の読み取り：`<fills>`（塗り）・`<fonts>`（文字の装飾）と、cellXfs の `fontId`・`fillId`。
* ブックで 1 回だけ読み、セルの `s` 属性（cellXfs の位置）→ 書式の番号 → 塗りと文字の装飾、と引けるようにする（`FormatTable`）。
* - 見える書式（塗り・太字・斜体・下線・取り消し線・自動でない色・比が 1 でない文字サイズのどれか）が無い xf は「書式なし」にまとめる。
* - 同じ見た目の書式は 1 つにまとめる（キーは数値だけ。ファイルの文字列は持たない）。まとめた後の数は `maxCellFormats` まで。
* - 塗りは 1 色に畳む：`solid`＝`fgColor`、ほかの模様＝前景と背景の混色、グラデーション＝50% の位置の色。
* - `applyFill`・`applyFont` は見ない（F5）。cellStyleXfs は読まない。フォント名は読まない。
* - 色は SpreadsheetML の `<color>`（cell-color.ts）をテーマで解いてから持つ。解けない値は例外にせず既定の色・書式なしに倒す。
* - 走査は線形（開きタグは `[^<>]*>`、閉じは `indexOf`）。
*/
/** 書式の無い表（styles.xml が無い・壊れているとき。共有する・変更しない） */
var EMPTY_FORMAT_TABLE = Object.freeze({
	formats: Object.freeze([]),
	xfFormat: /* @__PURE__ */ new Int32Array(0)
});
/** 整数の属性（空・数でない・小数は null） */
function intAttr$1(v) {
	if (v === void 0 || v.trim() === "") return null;
	const n = Number(v);
	return Number.isInteger(n) ? n : null;
}
/** 数の属性（空・数でない・無限は null） */
function numAttr$1(v) {
	if (v === void 0 || v.trim() === "") return null;
	const n = Number(v);
	return Number.isFinite(n) ? n : null;
}
/** 模様（`patternType`）→ 前景色の割合（ADR-0030 §決定 3 の表。`solid`・`none`・未知は含めない） */
var PATTERN_RATIO$1 = /* @__PURE__ */ new Map([
	["darkGray", .75],
	["mediumGray", .5],
	["lightGray", .25],
	["gray125", .125],
	["gray0625", .0625],
	["darkHorizontal", .5],
	["darkVertical", .5],
	["darkDown", .5],
	["darkUp", .5],
	["darkGrid", .75],
	["darkTrellis", .75],
	["lightHorizontal", .25],
	["lightVertical", .25],
	["lightDown", .25],
	["lightUp", .25],
	["lightGrid", 7 / 16],
	["lightTrellis", 6 / 16]
]);
/** 色の既定：前景・`solid` の塗りは黒、背景は白（F3） */
var BLACK = {
	rgb: 0,
	alpha: 1
};
var WHITE = {
	rgb: 16777215,
	alpha: 1
};
/** `<fgColor>`・`<bgColor>`（要素名 name）の色。要素が無い・自動・解けない色なら fallback */
function patternColor$1(body, name, fallback, theme, palette) {
	const tag = body === null ? null : findOpenTag(body, name);
	return (tag === null ? null : resolveCellColor(parseAttrs(tag), theme, palette)) ?? fallback;
}
/** `<patternFill>` 1 件（開きタグと中身）→ 塗りの色。塗りなし（無い・none・未知の模様）は undefined */
function readPatternFill(open, body, theme, palette) {
	const type = parseAttrs(open).patternType;
	if (type === void 0) return void 0;
	const fg = patternColor$1(body, "fgColor", BLACK, theme, palette);
	if (type === "solid") return {
		rgb: fg.rgb,
		alpha: 1
	};
	const ratio = PATTERN_RATIO$1.get(type);
	if (ratio === void 0) return void 0;
	const bg = patternColor$1(body, "bgColor", WHITE, theme, palette);
	return {
		rgb: mixColors(bg, fg, ratio).rgb,
		alpha: 1
	};
}
/**
* `<gradientFill>` の中身 → 50% の位置の色。`<stop>`（16 個まで）を `position` の順に並べ、前後の stop の補間。
* 2 stop は線形 RGB・3 stop 以上は sRGB（ADR-0028 §決定 6.4 と同じ）。stop が無ければ undefined。解けない色は黒。
*/
function readGradientFill(body, theme, palette) {
	if (body === null) return void 0;
	const stops = [];
	let seen = 0;
	forEachElement(body, "stop", (open, inner) => {
		if (seen >= 16) return;
		seen++;
		const pos = Math.min(Math.max(numAttr$1(parseAttrs(open).position) ?? 0, 0), 1);
		stops.push({
			pos,
			color: patternColor$1(inner, "color", BLACK, theme, palette)
		});
	});
	if (stops.length === 0) return void 0;
	stops.sort((a, b) => a.pos - b.pos);
	const mid = .5;
	let color = stops[stops.length - 1].color;
	if (stops[0].pos >= mid) color = stops[0].color;
	else for (let i = 1; i < stops.length; i++) {
		const hi = stops[i];
		if (hi.pos < mid) continue;
		const lo = stops[i - 1];
		const t = (mid - lo.pos) / (hi.pos - lo.pos);
		color = stops.length === 2 ? mixColorsLinear(lo.color, hi.color, t) : mixColors(lo.color, hi.color, t);
		break;
	}
	return {
		rgb: color.rgb,
		alpha: 1
	};
}
/** `<fill>` 1 件の中身 → 塗りの色（`patternFill` か `gradientFill`。どちらも無ければ undefined） */
function readFill(body, theme, palette) {
	if (body === null) return void 0;
	const pattern = findElement(body, "patternFill");
	if (pattern !== null) return readPatternFill(pattern.open, pattern.body, theme, palette);
	const gradient = findElement(body, "gradientFill");
	return gradient === null ? void 0 : readGradientFill(gradient.body, theme, palette);
}
/** 真偽の子要素（`<b>`・`<i>`・`<strike>`）。要素が無ければ偽、`val` が無ければ真、`val="0"`・`false` は偽 */
function boolChild(body, name) {
	const tag = findOpenTag(body, name);
	if (tag === null) return false;
	const v = parseAttrs(tag).val;
	return v !== "0" && v !== "false";
}
/** 下線の種類（`<u val>`）の対応。`none`・未知は含めない（下線なし）。会計用は普通の下線・二重下線に寄せる（F8） */
var UNDERLINE = /* @__PURE__ */ new Map([
	["single", "single"],
	["singleAccounting", "single"],
	["double", "double"],
	["doubleAccounting", "double"]
]);
/** 文字サイズの比の範囲（ADR-0030 §決定 4） */
var MIN_SIZE_RATIO = .5;
var MAX_SIZE_RATIO = 4;
/** `<font>` の `<sz val>`（pt）。無い・数でない・0 以下は null */
function fontSize(body) {
	const tag = findOpenTag(body, "sz");
	const v = tag === null ? null : numAttr$1(parseAttrs(tag).val);
	return v !== null && v > 0 ? v : null;
}
/**
* `<font>` 1 件の中身 → 文字の装飾。
* @param baseSize 先頭の `<font>` の `sz`（pt。無ければ null＝比を持たない）
*/
function readFont(body, baseSize, theme, palette) {
	const font = {};
	if (body === null) return font;
	if (boolChild(body, "b")) font.bold = true;
	if (boolChild(body, "i")) font.italic = true;
	if (boolChild(body, "strike")) font.strike = true;
	const u = findOpenTag(body, "u");
	if (u !== null) {
		const kind = UNDERLINE.get(parseAttrs(u).val ?? "single");
		if (kind !== void 0) font.underline = kind;
	}
	const size = fontSize(body);
	if (size !== null && baseSize !== null) {
		const ratio = Math.round(Math.min(Math.max(size / baseSize, MIN_SIZE_RATIO), MAX_SIZE_RATIO) * 100) / 100;
		if (ratio !== 1) font.sizeRatio = ratio;
	}
	const colorTag = findOpenTag(body, "color");
	const color = colorTag === null ? null : resolveCellColor(parseAttrs(colorTag), theme, palette);
	if (color !== null && color.rgb !== 0) font.color = color;
	return font;
}
/**
* styles.xml から書式の表を作る（ブックで 1 回）。
* @param xml styles.xml（無ければ null＝書式なし）
* @param theme ブックのテーマ（色の `theme` を解く）
* @param limits `maxFonts`・`maxFills` を超えた要素は読まない（それを参照するセルは装飾なし・塗りなし）。
*   重複を除いた書式の数が `maxCellFormats` を超えた書式のセルは書式なし
*/
function parseCellFormats(xml, theme, limits) {
	if (xml === null) return EMPTY_FORMAT_TABLE;
	const cellXfs = findElement(xml, "cellXfs");
	if (cellXfs === null || cellXfs.body === null) return EMPTY_FORMAT_TABLE;
	const palette = parseIndexedColors(xml);
	const fills = [];
	const fillBlock = findElement(xml, "fills");
	if (fillBlock?.body) forEachElement(fillBlock.body, "fill", (_open, body) => {
		if (fills.length >= limits.maxFills) return;
		fills.push(readFill(body, theme, palette));
	});
	const fonts = [];
	const fontBlock = findElement(xml, "fonts");
	if (fontBlock?.body) {
		let baseSize = null;
		forEachElement(fontBlock.body, "font", (_open, body) => {
			if (fonts.length >= limits.maxFonts) return;
			if (fonts.length === 0) baseSize = body === null ? null : fontSize(body);
			fonts.push(readFont(body, baseSize, theme, palette));
		});
	}
	const formats = [];
	const keyToIndex = /* @__PURE__ */ new Map();
	const xfFormat = [];
	for (const m of cellXfs.body.matchAll(new RegExp(String.raw`<${NS}xf\b[^<>]*>`, "g"))) {
		const a = parseAttrs(m[0]);
		const fontId = intAttr$1(a.fontId);
		const fillId = intAttr$1(a.fillId);
		const font = fontId !== null && fontId >= 0 ? fonts[fontId] : void 0;
		const fill = fillId !== null && fillId >= 0 ? fills[fillId] : void 0;
		const format = {};
		if (fill) format.fill = fill;
		if (font?.color) format.color = font.color;
		if (font?.bold) format.bold = true;
		if (font?.italic) format.italic = true;
		if (font?.strike) format.strike = true;
		if (font?.underline) format.underline = font.underline;
		if (font?.sizeRatio !== void 0) format.sizeRatio = font.sizeRatio;
		if (Object.keys(format).length === 0) {
			xfFormat.push(-1);
			continue;
		}
		const key = formatKey(format);
		let idx = keyToIndex.get(key);
		if (idx === void 0) {
			if (formats.length >= limits.maxCellFormats) {
				xfFormat.push(-1);
				continue;
			}
			idx = formats.length;
			formats.push(format);
			keyToIndex.set(key, idx);
		}
		xfFormat.push(idx);
	}
	return {
		formats,
		xfFormat: Int32Array.from(xfFormat)
	};
}
/** 書式の重複の除去のキー（数値だけで組む：塗り｜文字の色｜装飾のビット｜サイズ比×100） */
function formatKey(f) {
	const bits = (f.bold ? 1 : 0) | (f.italic ? 2 : 0) | (f.strike ? 4 : 0) | (f.underline === "single" ? 8 : 0) | (f.underline === "double" ? 16 : 0);
	return `${f.fill?.rgb ?? -1}|${f.color?.rgb ?? -1}|${bits}|${Math.round((f.sizeRatio ?? 1) * 100)}`;
}
/**
* 「紙」の塗りか：解決後の塗りが不透明な `#FFFFFF`（`rgb`・`theme`・`indexed` など出どころは問わない）。
* 紙は塗りなしと同じに描く（ADR-0030 1.1）ため、表の範囲を広げる判定（sheet.ts）と描き分け（viewer）の両方がこれを使う。
* `#FEFEFE`・`#F2F2F2`・半透明の白は紙ではない（alpha < 1）。
*/
function isPaperFill(fill) {
	return fill !== void 0 && fill.rgb === 16777215 && fill.alpha >= 1;
}
/**
* 書式の色に `theme` の指定があるか（`<fonts>`・`<fills>`・`<borders>` の中だけを見る）。無ければテーマを読まなくても
* 色は決まるので、呼び出し側（index.ts）はテーマの読み取りを省ける（図の無いブックでテーマの読み取りを増やさない。ISSUE-94）。
*/
function stylesUseTheme(xml) {
	if (xml === null) return false;
	for (const name of [
		"fonts",
		"fills",
		"borders"
	]) {
		const body = findElement(xml, name)?.body;
		if (body !== null && body !== void 0 && /\stheme\s*=/.test(body)) return true;
	}
	return false;
}
// src/core/drawing-shape.ts
/** 線として扱う preset（`xdr:sp` の `prstGeom prst` がこれなら種類 `line`）。`xdr:cxnSp` はこれ以外の形でも線 */
var LINE_PRESETS = /* @__PURE__ */ new Set([
	"line",
	"straightConnector1",
	"bentConnector2",
	"bentConnector3",
	"bentConnector4",
	"bentConnector5",
	"curvedConnector2",
	"curvedConnector3",
	"curvedConnector4",
	"curvedConnector5"
]);
var DASHES = /* @__PURE__ */ new Set([
	"solid",
	"dot",
	"dash",
	"lgDash",
	"dashDot",
	"lgDashDot",
	"lgDashDotDot",
	"sysDash",
	"sysDot",
	"sysDashDot",
	"sysDashDotDot"
]);
var ARROWS = /* @__PURE__ */ new Set([
	"triangle",
	"stealth",
	"diamond",
	"oval",
	"arrow"
]);
/** 回転の 1 周（1/60000 度） */
var FULL_TURN = 216e5;
/** 調整値・自由図形の整数の引数の上限（±2^31-1。値を有限に保つためだけ。Excel の U 字線は端の揃った形で adj1=1,800,000 を書く。異常値の px 対策は表示側） */
var ADJ_LIMIT = 2147483647;
/** 線の太さの上限（EMU。ECMA-376 の `ST_LineWidth` の最大値 1,584pt） */
var MAX_LINE_WIDTH = 20116800;
/** `lnRef idx` 1・2・3 の太さ（EMU。Office 2013 以降の既定テーマの値。テーマは読まない） */
var LN_REF_WIDTH = {
	1: 6350,
	2: 12700,
	3: 19050
};
/** 塗りの要素名（`spPr` 直下・`a:ln` 直下） */
var FILL_RE = new RegExp(String.raw`<${NS}(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)(?=[\s>/])`);
/** 開き・閉じタグ 1 つ（名前空間の接頭辞可。`[^<>]*>` で打ち切る）。`spPr` 直下の `extLst` を探す深さの計算用 */
var TAG_RE$2 = new RegExp(String.raw`<(/?)${NS}([A-Za-z_][\w.-]*)(?=[\s>/])[^<>]*>`, "g");
/** 整数の文字（前後の空白可）を数にする。整数でない・桁が多すぎるなら null */
function intAttr(s) {
	if (s === void 0) return null;
	const t = s.trim();
	return /^-?\d{1,15}$/.test(t) ? Number(t) : null;
}
/** XML の真偽（`1`・`true`）。無ければ false */
function boolAttr$1(s) {
	return s === "1" || s === "true";
}
/**
* `spPr` 直下（深さ 0）で最初に出る `extLst` の位置。無ければ -1。
* Office 拡張（`a14:hiddenFill` などに本物でない塗り・線が入る）はここより後ろを読まないための位置。
* `a:ln`・`a:blip` の中の `extLst` では切らない（`a:ln` が途中で切れるため）。深さは開き・閉じタグの数で数え、線形時間。
*/
function topLevelExtLst(xml) {
	let depth = 0;
	TAG_RE$2.lastIndex = 0;
	for (let m = TAG_RE$2.exec(xml); m !== null; m = TAG_RE$2.exec(xml)) {
		if (m[1] === "/") {
			if (depth > 0) depth--;
			continue;
		}
		if (depth === 0 && m[2] === "extLst") return m.index;
		if (!m[0].endsWith("/>")) depth++;
	}
	return -1;
}
/** 回転を 0〜21,599,999 に正規化する。整数でなければ 0 */
function normalizeRot(s) {
	const n = intAttr(s);
	if (n === null) return 0;
	return (n % FULL_TURN + FULL_TURN) % FULL_TURN;
}
/** 開きタグの `idx`（`lnRef`・`fillRef`）を整数で返す。タグが無い・不正なら null */
function refIdx(spBody, name) {
	const tag = findOpenTag(spBody, name);
	return tag === null ? null : intAttr(parseAttrs(tag).idx);
}
/** 塗りの要素の最初の名前（`noFill` など）。無ければ null */
function firstFill(xml) {
	return FILL_RE.exec(xml)?.[1] ?? null;
}
/** 塗りの要素の最初の 1 つを開きタグと中身ごと切り出す。無ければ null */
function firstFillEl(xml) {
	if (xml === null) return null;
	const m = FILL_RE.exec(xml);
	if (m === null) return null;
	const name = m[1];
	const el = findElement(xml, name, m.index);
	return el === null ? null : {
		name,
		open: el.open,
		body: el.body
	};
}
/** `xdr:style` の参照（`fillRef`・`lnRef`）を `idx` と中身で読む。無ければ null */
function styleRef(spBody, name) {
	const el = findElement(spBody, name);
	return el === null ? null : {
		idx: intAttr(parseAttrs(el.open).idx),
		body: el.body
	};
}
/**
* `xdr:sp`・`xdr:cxnSp` の中身から、回転・反転・枠線と塗りの有無・`a:ln`・形を読む。
* 色の解釈は `drawing-color.ts`（ここは塗りの要素と `xdr:style` の参照の切り出しまで）。
* - 枠線が無い＝`a:ln` に `a:noFill`、または `a:ln` に線の塗りの指定が無く `lnRef idx="0"`。それ以外は「あり」（`xdr:style` も線の塗りも無ければあり）
* - 塗りが無い＝`spPr` 直下に `a:noFill`、または塗りの指定が無く `fillRef idx="0"`。それ以外は「あり」
* 閉じの無い `spPr`・`a:ln` でも例外にしない。
*/
function readShapeProps(spBody) {
	let pr = findElement(spBody, "spPr")?.body ?? "";
	const ext = topLevelExtLst(pr);
	if (ext >= 0) pr = pr.slice(0, ext);
	const xfrm = findOpenTag(pr, "xfrm");
	const xa = xfrm === null ? null : parseAttrs(xfrm);
	const geom = findElement(pr, "prstGeom");
	const prst = geom === null ? null : parseAttrs(geom.open).prst ?? null;
	const cust = geom === null ? findElement(pr, "custGeom") : null;
	const lnEl = findElement(pr, "ln");
	const ln = lnEl === null ? null : {
		open: lnEl.open,
		body: lnEl.body
	};
	let direct = pr;
	if (lnEl !== null) {
		const start = pr.indexOf(lnEl.open);
		direct = pr.slice(0, start) + (lnEl.end >= pr.length ? "" : pr.slice(lnEl.end));
	}
	const lnRefIdx = refIdx(spBody, "lnRef");
	const lnFill = ln?.body == null ? null : firstFill(ln.body);
	const hasOutline = lnFill !== null ? lnFill !== "noFill" : lnRefIdx !== 0;
	const fill = firstFill(direct);
	const hasFill = fill !== null ? fill !== "noFill" : refIdx(spBody, "fillRef") !== 0;
	return {
		prst,
		rot: normalizeRot(xa?.rot),
		flipH: boolAttr$1(xa?.flipH),
		flipV: boolAttr$1(xa?.flipV),
		hasOutline,
		hasFill,
		ln,
		lnRefIdx,
		fillEl: firstFillEl(direct),
		lnFillEl: firstFillEl(ln?.body ?? null),
		fillRef: styleRef(spBody, "fillRef"),
		lnRef: styleRef(spBody, "lnRef"),
		geomBody: geom?.body ?? null,
		custGeom: cust === null ? null : { body: cust.body }
	};
}
/** `prst` が線として扱う preset か（`xdr:sp` を線にするかの判定） */
function isLinePreset(prst) {
	return prst !== null && LINE_PRESETS.has(prst);
}
/** 矢じりの大きさ（`sm`・`med`・`lg`）。無い・未知の値は `med` */
function arrowSize(s) {
	return s === "sm" || s === "lg" ? s : "med";
}
/** 線の端（`a:headEnd`・`a:tailEnd`）。`type` が無い・`none` は矢じりなし（null）。未知の種類は `triangle` */
function lineEnd(lnBody, name) {
	const tag = findOpenTag(lnBody, name);
	if (tag === null) return null;
	const a = parseAttrs(tag);
	if (a.type === void 0 || a.type === "none") return null;
	return {
		type: ARROWS.has(a.type) ? a.type : "triangle",
		w: arrowSize(a.w),
		len: arrowSize(a.len)
	};
}
/** 調整値 adj1〜adj3 を読む（`a:gd name="adjN" fmla="val N"` だけ。式は読まない。無ければ 50000。範囲外は丸める） */
function readAdj(geomBody) {
	const adj = [
		5e4,
		5e4,
		5e4
	];
	if (geomBody === null) return adj;
	forEachElement(geomBody, "gd", (open) => {
		const a = parseAttrs(open);
		const m = /^adj([123])$/.exec(a.name ?? "");
		const v = /^val\s+(-?\d{1,15})$/.exec((a.fmla ?? "").trim());
		if (m === null || v === null) return;
		adj[Number(m[1]) - 1] = Math.min(Math.max(Number(v[1]), -2147483647), ADJ_LIMIT);
	});
	return adj;
}
/**
* 枠線（線）の太さと線種。線・矢印（`readLine`）と図形の枠線で同じ規則を使う。
* 太さは `a:ln w`（0〜20,116,800 に丸める）、無ければ `lnRef idx` 1・2・3 → 6350・12700・19050 EMU、どちらも無ければ 0。
* 線種は `a:prstDash val`（無い・未知は実線）。
*/
function readStroke(props) {
	const w = intAttr((props.ln === null ? null : parseAttrs(props.ln.open))?.w);
	const widthEmu = w !== null ? Math.min(Math.max(w, 0), MAX_LINE_WIDTH) : LN_REF_WIDTH[props.lnRefIdx ?? -1] ?? 0;
	const lnBody = props.ln?.body ?? null;
	const dashTag = lnBody === null ? null : findOpenTag(lnBody, "prstDash");
	const dashVal = dashTag === null ? void 0 : parseAttrs(dashTag).val;
	return {
		widthEmu,
		dash: dashVal !== void 0 && DASHES.has(dashVal) ? dashVal : "solid"
	};
}
/**
* 線・矢印の形と見た目を読む。形は `prst`（線の preset 以外・`custGeom` は直線＝`line`）、
* 太さ・線種は `readStroke`、矢じりは `a:headEnd`（始点）・`a:tailEnd`（終点）。
*/
function readLine(props) {
	const geom = isLinePreset(props.prst) ? props.prst : "line";
	const { widthEmu, dash } = readStroke(props);
	const lnBody = props.ln?.body ?? null;
	return {
		geom,
		adj: readAdj(props.geomBody),
		widthEmu,
		dash,
		head: lnBody === null ? null : lineEnd(lnBody, "headEnd"),
		tail: lnBody === null ? null : lineEnd(lnBody, "tailEnd")
	};
}
// src/core/drawing-color.ts
/**
* 図形・線・文字の色の読み取り（画面に依存しない）：色の出どころの順と、塗りの代表色。
* - 塗り：`spPr` 直下の塗り → `a:fillRef` の子の色 → テーマの `lt1`。
* - 枠線・線：`a:ln` の塗り → `a:lnRef` の子の色 → テーマの `tx1`。
* - 文字：ランの `a:rPr` の塗り → `a:fontRef` の子の色 → テーマの `tx1`。
* - 解けない色（不正な値・`phClr`・色の要素が無い）はその出どころを飛ばして次へ進む。`a:grpFill` は未指定として扱う。
* - グラデーション・模様・画像の塗りは代表色の 1 色にする（`a:gradFill`＝50% の位置・`a:pattFill`＝前景と背景の混色・`a:blipFill`＝#D9D9D9）。
* 塗り・枠線の「有無」は `drawing-shape.ts`（ADR-0027）が決める。ここは「ある」と決まった塗り・線の色を返す。
* 色の要素そのものの解釈（変換・テーマ）は `color.ts`。結果は RGB の整数と不透明度だけで、ファイルの文字列を持たない。
* 閉じの無い要素・大量の stop・変換の列でも例外にせず、走査は線形時間（上限は `limits.ts`）。
*/
/** 画像の塗り（`a:blipFill`）の代表色（画像は読まない） */
var BLIP_COLOR = 14277081;
/** グラデーションの代表色にする位置（1/1000 %＝50%） */
var GRADIENT_MID = 5e4;
/** 模様の名前（`pct5`〜`pct90`）→ 前景色の割合。この表に無い模様は 50% */
var PATTERN_RATIO = new Map([
	5,
	10,
	20,
	25,
	30,
	40,
	50,
	60,
	70,
	75,
	80,
	90
].map((n) => [`pct${n}`, n / 100]));
/** テーマの色の位置の色（テーマに無ければ既定のテーマ）。毎回新しい色を返す */
function themeColor(theme, slot) {
	return {
		rgb: theme.get(slot) ?? DEFAULT_THEME.get(slot) ?? 0,
		alpha: 1
	};
}
/**
* `a:gradFill` の代表色：`a:gsLst` の stop（16 個まで）を `pos` の順に並べ、50% の位置の色（前後の stop の補間）。解ける stop が無ければ null。
* 補間は、解けた stop が 2 つなら線形 RGB、3 つ以上なら sRGB の値（Excel の画素に合わせた。ISSUE-95・U7：2 stop の赤→青の中央が #BC00B8、3 stop の黒→赤→白の 50% の G が #46＝sRGB の補間）。
*/
function gradientColor(body, theme) {
	const list = findElement(body, "gsLst")?.body;
	if (list === null || list === void 0) return null;
	const stops = [];
	let seen = 0;
	forEachElement(list, "gs", (open, b) => {
		if (seen >= 16) return;
		seen++;
		if (b === null) return;
		const color = resolveColor(b, theme);
		if (color === null) return;
		const pos = intAttr(parseAttrs(open).pos) ?? 0;
		stops.push({
			pos: Math.min(Math.max(pos, 0), 1e5),
			color
		});
	});
	if (stops.length === 0) return null;
	stops.sort((a, b) => a.pos - b.pos);
	if (stops[0].pos >= GRADIENT_MID) return stops[0].color;
	for (let i = 1; i < stops.length; i++) {
		const hi = stops[i];
		if (hi.pos < GRADIENT_MID) continue;
		const lo = stops[i - 1];
		const t = (GRADIENT_MID - lo.pos) / (hi.pos - lo.pos);
		return stops.length === 2 ? mixColorsLinear(lo.color, hi.color, t) : mixColors(lo.color, hi.color, t);
	}
	return stops[stops.length - 1].color;
}
/** `a:pattFill` の代表色：前景（`a:fgClr`）と背景（`a:bgClr`）を、前景の割合で混ぜる。片方しか解けなければその色、どちらも解けなければ null */
function patternColor(open, body, theme) {
	const side = (name) => {
		const inner = findElement(body, name)?.body;
		return inner === null || inner === void 0 ? null : resolveColor(inner, theme);
	};
	const fg = side("fgClr");
	const bg = side("bgClr");
	if (fg === null || bg === null) return fg ?? bg;
	return mixColors(bg, fg, PATTERN_RATIO.get(parseAttrs(open).prst ?? "") ?? .5);
}
/**
* 塗りの要素 1 つの色（ADR-0028 §決定 6.4）。`a:solidFill`＝その色・`a:gradFill`・`a:pattFill`＝代表色・`a:blipFill`＝#D9D9D9。
* `a:noFill`・`a:grpFill`・色の要素が無い・解けないときは null（次の出どころへ進む）。
*/
function fillElementColor(el, theme) {
	if (el.name === "blipFill") return {
		rgb: BLIP_COLOR,
		alpha: 1
	};
	if (el.body === null) return null;
	switch (el.name) {
		case "solidFill": return resolveColor(el.body, theme);
		case "gradFill": return gradientColor(el.body, theme);
		case "pattFill": return patternColor(el.open, el.body, theme);
		default: return null;
	}
}
/**
* `xdr:style` の参照の子の色。idx が 1 以上（1001 以上の背景の塗りも同じ）の参照を見る。idx が 0 なら見ない。
* idx が数値でない・欠けている（null。Excel はほぼ書かないが安全側）ときは、子の色があればそれを使う。参照が無い・子が無い・解けないときは null
*/
function refColor(ref, theme) {
	if (ref === null || ref.body === null || ref.idx !== null && ref.idx < 1) return null;
	return resolveColor(ref.body, theme);
}
/**
* 図形の塗りの色：`spPr` 直下の塗り → `a:fillRef` の子の色 → テーマの `lt1`。
* 塗りがあると決まった図形（`props.hasFill`）にだけ呼ぶ（有無はここで判定しない）。
* `xdr:style` も `spPr` の塗りも無い図形は、有無の規則（ADR-0027）で「塗り・線あり」なので、色は既定の `lt1`（枠線・線は `tx1`）になる。ISSUE-95 の「既定」どおり。
*/
function readFillColor(props, theme) {
	return (props.fillEl === null ? null : fillElementColor(props.fillEl, theme)) ?? refColor(props.fillRef, theme) ?? themeColor(theme, "lt1");
}
/**
* 図形の枠線・線の色：`a:ln` の `a:solidFill`・`a:gradFill`・`a:pattFill` → `a:lnRef` の子の色 → テーマの `tx1`。
* 枠線が見えると決まったもの（`props.hasOutline`）にだけ呼ぶ。矢じりも同じ色。
*/
function readOutlineColor(props, theme) {
	const el = props.lnFillEl;
	return (el === null || el.name === "blipFill" ? null : fillElementColor(el, theme)) ?? refColor(props.lnRef, theme) ?? themeColor(theme, "dk1");
}
/** 図形の文字の既定の色：`xdr:style` の `a:fontRef` の子の色 → テーマの `tx1`（`spBody` は `xdr:sp` の中身） */
function readFontColor(spBody, theme) {
	const body = findElement(spBody, "fontRef")?.body;
	return (body === null || body === void 0 ? null : resolveColor(body, theme)) ?? themeColor(theme, "dk1");
}
/** 塗りの要素の名前（`a:rPr` の直下で最初に出るもの） */
var TEXT_FILL_NAMES = /* @__PURE__ */ new Set([
	"noFill",
	"solidFill",
	"gradFill",
	"blipFill",
	"pattFill",
	"grpFill"
]);
/** 開き・閉じタグ 1 つ（名前空間の接頭辞可。`[^<>]*>` で打ち切る）。`a:rPr` の直下の塗りを探す深さの計算用 */
var TAG_RE$1 = new RegExp(String.raw`<(/?)${NS}([A-Za-z_][\w.-]*)(?=[\s>/])[^<>]*>`, "g");
/**
* `a:rPr` の中身の直下（深さ 0）で最初に出る塗りの要素。無ければ null。
* `a:ln`（文字の輪郭）・`a:uFill`（下線の色）の中の塗りを文字の色にしないよう、深さを開き・閉じタグの数で数える（線形時間）。
*/
function firstDirectFill(rPrBody) {
	let depth = 0;
	TAG_RE$1.lastIndex = 0;
	for (let m = TAG_RE$1.exec(rPrBody); m !== null; m = TAG_RE$1.exec(rPrBody)) {
		if (m[1] === "/") {
			if (depth > 0) depth--;
			continue;
		}
		const selfClosed = m[0].endsWith("/>");
		if (depth === 0 && TEXT_FILL_NAMES.has(m[2])) {
			const el = findElement(rPrBody, m[2], m.index);
			TAG_RE$1.lastIndex = 0;
			return el === null ? null : {
				name: m[2],
				open: el.open,
				body: el.body
			};
		}
		if (!selfClosed) depth++;
	}
	TAG_RE$1.lastIndex = 0;
	return null;
}
/**
* 文字のラン（`a:r`・`a:fld`・`a:br`）の色：`a:rPr` の `a:solidFill`・`a:gradFill`（代表色）→ `fallback`（図形の `fontRef` の色。`readFontColor`）。
* `a:noFill` は透明（不透明度 0）。解けない色・`a:pattFill` などは `fallback`。`a:ln`（文字の輪郭）・`a:highlight` は読まない。
* @param rPrBody `a:rPr` の中身（`a:rPr` が無い・自己閉じのときは null）
*/
function readRunColor(rPrBody, fallback, theme) {
	const el = rPrBody === null ? null : firstDirectFill(rPrBody);
	if (el === null) return { ...fallback };
	if (el.name === "noFill") return {
		rgb: fallback.rgb,
		alpha: 0
	};
	if (el.name === "solidFill" || el.name === "gradFill") return fillElementColor(el, theme) ?? { ...fallback };
	return { ...fallback };
}
// src/core/drawing-group.ts
/**
* グループ（`xdr:grpSp`）の座標系と、子の `a:xfrm` の読み取り（ISSUE-89）。
* DOM に依存せず、`findElement`・`findOpenTag` だけで線形時間に走査する。値は EMU の整数（位置は負でもよい）。
* グループ自身の回転・反転は読まない（ADR-0027 §決定 5）。図の名前・文字はログに出さない（console を使わない）。
*/
/** グループの中身の先頭にある `nvGrpSpPr` の開きタグ（先頭の空白可） */
var NV_RE = new RegExp(String.raw`^\s*<${NS}nvGrpSpPr\b[^<>]*>`);
/** `nvGrpSpPr` の直後にある `grpSpPr` の開きタグ（空白可。`lastIndex` を設定して `y` で照合する） */
var PR_RE = new RegExp(String.raw`\s*<${NS}grpSpPr\b[^<>]*>`, "y");
/**
* `xfrm` の中身から、`offName`（`off`・`chOff`）と `extName`（`ext`・`chExt`）の 2 つの開きタグで矩形を作る。
* どちらかが無い・整数でない・大きさが負なら null。
*/
function rectOf(body, offName, extName) {
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
	return {
		x,
		y,
		cx,
		cy
	};
}
/**
* 最初の `a:xfrm`（グラフの `xdr:xfrm` も同じ）の位置と大きさ（`a:off`・`a:ext`）を読む。
* `xml` は図 1 つ分（`sp`・`cxnSp`・`pic`・`graphicFrame`）か、グループの `grpSpPr` の中身。無い・読めないなら null。
*/
function readXfrm(xml) {
	const body = findElement(xml, "xfrm")?.body;
	return body === null || body === void 0 ? null : rectOf(body, "off", "ext");
}
/**
* グループ（`xdr:grpSp` の中身）の頭にある、グループ自身の `nvGrpSpPr`（の `cNvPr`）と、その直後の `grpSpPr` を取り出す。
* 先頭（`nvGrpSpPr`）と直後（`grpSpPr`）だけを見る：壊れたグループ（`nvGrpSpPr`・`grpSpPr` が無い）で、
* 子（入れ子のグループ）の `cNvPr`・`grpSpPr` を自分のものとして拾わないため。
*/
function readGroupHead(grpBody) {
	if (NV_RE.exec(grpBody) === null) return {
		cNvPr: null,
		pr: null
	};
	const nv = findElement(grpBody, "nvGrpSpPr", 0);
	if (nv === null) return {
		cNvPr: null,
		pr: null
	};
	const cNvPr = nv.body === null ? null : findOpenTag(nv.body, "cNvPr");
	PR_RE.lastIndex = nv.end;
	const m = PR_RE.exec(grpBody);
	return {
		cNvPr,
		pr: m === null ? null : findElement(grpBody, "grpSpPr", m.index)
	};
}
/**
* グループ（`xdr:grpSp` の中身）の座標系を読む。展開できないグループは null：`grpSpPr`・`a:xfrm`・`a:chOff`・`a:chExt` が無い、
* 数が整数でない、`chExt` の幅か高さが 0 以下（割り算ができない）、`needOwn` なのに `a:off`・`a:ext` が無い。
* @param pr グループ自身の `grpSpPr`（`readGroupHead` の結果）
* @param needOwn 入れ子のグループか（親の座標系での位置が要る。最も外側のグループはアンカーの矩形を使うので要らない）
*/
function readGroupFrame(pr, needOwn) {
	if (pr === null || pr.body === null) return null;
	const xfrm = findElement(pr.body, "xfrm")?.body;
	if (xfrm === null || xfrm === void 0) return null;
	const ch = rectOf(xfrm, "chOff", "chExt");
	if (ch === null || ch.cx <= 0 || ch.cy <= 0) return null;
	const own = needOwn ? rectOf(xfrm, "off", "ext") : null;
	if (needOwn && own === null) return null;
	return {
		ch,
		own,
		end: pr.end
	};
}
// src/core/drawing-geom.ts
/**
* 図形（`xdr:sp`）の形の読み取り：`a:prstGeom` の調整値（`a:avLst`）と、自由図形（`a:custGeom`）の定義。
* 自由図形の式・path はファイルの値で非信頼の入力。字句（演算子の表・引数の数・整数か名前）を確かめ、上限（guide・命令の数）で守る。
* 式の評価と経路は geom の `evaluate.ts`（InterHub ISSUE-90）。ここは評価器がそのまま使える形（`CustGeomDef`）に読むところまで。
* DOM に依存せず、`findElement`・`forEachElement`・`findOpenTag` と線形時間の正規表現だけで走査する。
* 名前（guide・調整値）はファイルの文字列。`Map` のキーにだけ使い、DOM・CSS・ログに出さない（console を使わない）。
*/
/** guide・調整値の名前：英字で始まる英数字 16 文字まで */
var NAME_RE = /^[A-Za-z][A-Za-z0-9]{0,15}$/;
/** 角度の組み込み変数（`3cd4`・`3cd8`・`5cd8`・`7cd8`。数字で始まるので `NAME_RE` に入らない） */
var ANGLE_VAR_RE = /^[357]cd[48]$/;
/** `fmla` の長さの上限（文字数。最長の式でも数十文字）。超えたら字句の誤り */
var MAX_FMLA_LENGTH = 200;
/** path 命令の要素名 → `[評価器の命令名, 子の `a:pt` の数]`（`arcTo` は属性 4 つ・`pt` なし） */
var COMMANDS = /* @__PURE__ */ new Map([
	["moveTo", ["M", 1]],
	["lnTo", ["L", 1]],
	["arcTo", ["A", 0]],
	["quadBezTo", ["Q", 2]],
	["cubicBezTo", ["C", 3]],
	["close", ["Z", 0]]
]);
/** `a:path fill` の値（`none` 以外は塗りの濃淡の種類） */
var PATH_FILLS = /* @__PURE__ */ new Set([
	"none",
	"norm",
	"lighten",
	"lightenLess",
	"darken",
	"darkenLess"
]);
/** 開きタグ 1 つ（名前空間の接頭辞可。`[^<>]*>` で打ち切る）。path の中の命令・`a:pt` を順に拾う */
var TAG_RE = new RegExp(String.raw`<${NS}([A-Za-z_][\w.-]*)(?=[\s>/])[^<>]*>`, "g");
/** 整数を ±2^31-1 に丸める（調整値と同じ。値を有限に保つためだけ） */
var clampInt = (n) => Math.min(Math.max(n, -ADJ_LIMIT), ADJ_LIMIT);
/** 引数の字句：整数（15 桁まで。±2^31-1 に丸める）か名前（組み込み変数・guide）。誤りなら null */
function parseArg(s) {
	if (s === void 0) return null;
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
function readPresetAdjustments(geomBody, max) {
	const adj = [];
	if (geomBody === null) return adj;
	forEachElement(geomBody, "gd", (open) => {
		if (adj.length >= max) return;
		const a = parseAttrs(open);
		const name = a.name ?? "";
		const v = /^val\s+(-?\d{1,15})$/.exec((a.fmla ?? "").trim());
		if (!NAME_RE.test(name) || v === null) return;
		adj.push([name, clampInt(Number(v[1]))]);
	});
	return adj;
}
/** guide 1 つの式（`fmla`）を `[演算子, 引数…]` に読む。演算子が未知・引数の数が違う・引数の字句が誤りなら null */
function parseFmla(fmla) {
	const t = (fmla ?? "").trim();
	if (t.length === 0 || t.length > MAX_FMLA_LENGTH) return null;
	const parts = t.split(/\s+/);
	const op = parts[0];
	if (GUIDE_ARITY.get(op) !== parts.length - 1) return null;
	const out = [op];
	for (let i = 1; i < parts.length; i++) {
		const a = parseArg(parts[i]);
		if (a === null) return null;
		out.push(a);
	}
	return out;
}
/** 命令（または命令の無い path）を 1 つ数える。1 図形の上限・シートの残りを超えたら `counter.over` を立てて false */
function charge(counter, budget) {
	counter.n++;
	if (counter.n > budget.maxCommands) counter.over = "limit";
	else if (counter.base + counter.n > budget.sheetRoom) counter.over = "sheet-limit";
	return counter.over === null;
}
/** `a:path 1 つ` の中身（命令と `a:pt` の並び）を命令の列にする。誤りなら null。`counter.n` に命令の数を足し、上限を超えたら `counter.over` を立てて打ち切る */
function parseCommands(body, counter, budget) {
	const cmds = [];
	if (body === null) return cmds;
	/** 今の命令。`need` 個の `a:pt` を集める（`arcTo` は属性から引数を取る） */
	let cur = null;
	const finish = () => {
		if (cur === null) return true;
		const ok = cur.args.length === (cur.op === "A" ? 4 : cur.need * 2);
		if (ok) cmds.push([cur.op, ...cur.args]);
		cur = null;
		return ok;
	};
	TAG_RE.lastIndex = 0;
	for (let m = TAG_RE.exec(body); m !== null; m = TAG_RE.exec(body)) {
		const name = m[1];
		if (name === "pt") {
			if (cur === null || cur.op === "A" || cur.args.length >= cur.need * 2) return null;
			const a = parseAttrs(m[0]);
			const x = parseArg(a.x);
			const y = parseArg(a.y);
			if (x === null || y === null) return null;
			cur.args.push(x, y);
			continue;
		}
		const def = COMMANDS.get(name);
		if (def === void 0) return null;
		if (!finish()) return null;
		if (!charge(counter, budget)) return null;
		cur = {
			op: def[0],
			need: def[1],
			args: []
		};
		if (def[0] === "A") {
			const a = parseAttrs(m[0]);
			for (const k of [
				"wR",
				"hR",
				"stAng",
				"swAng"
			]) {
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
function readCustGeom(body, budget) {
	if (body === null) return {
		ok: false,
		reason: "invalid"
	};
	let guides = 0;
	let bad = false;
	const av = [];
	const avBody = findElement(body, "avLst")?.body ?? null;
	if (avBody !== null) forEachElement(avBody, "gd", (open) => {
		if (bad || ++guides > budget.maxGuides) return;
		const a = parseAttrs(open);
		const v = /^val\s+(-?\d{1,15})$/.exec((a.fmla ?? "").trim());
		if (!NAME_RE.test(a.name ?? "") || v === null) {
			bad = true;
			return;
		}
		av.push([a.name, clampInt(Number(v[1]))]);
	});
	const gd = [];
	const gdBody = findElement(body, "gdLst")?.body ?? null;
	if (gdBody !== null) forEachElement(gdBody, "gd", (open) => {
		if (bad || ++guides > budget.maxGuides) return;
		const a = parseAttrs(open);
		const f = parseFmla(a.fmla);
		if (!NAME_RE.test(a.name ?? "") || f === null) {
			bad = true;
			return;
		}
		gd.push([a.name, ...f]);
	});
	if (guides > budget.maxGuides) return {
		ok: false,
		reason: "limit"
	};
	if (guides > budget.sheetRoom) return {
		ok: false,
		reason: "sheet-limit"
	};
	if (bad) return {
		ok: false,
		reason: "invalid"
	};
	let rect;
	const rectTag = findOpenTag(body, "rect");
	if (rectTag !== null) {
		const a = parseAttrs(rectTag);
		const r = [
			parseArg(a.l),
			parseArg(a.t),
			parseArg(a.r),
			parseArg(a.b)
		];
		if (r.some((v) => v === null)) return {
			ok: false,
			reason: "invalid"
		};
		rect = r;
	}
	const paths = [];
	const counter = {
		n: 0,
		over: null,
		base: guides
	};
	const pathLst = findElement(body, "pathLst")?.body ?? null;
	if (pathLst === null) return {
		ok: false,
		reason: "invalid"
	};
	forEachElement(pathLst, "path", (open, pathBody) => {
		if (bad || counter.over !== null) return;
		if (paths.length >= budget.maxCommands) {
			counter.over = "limit";
			return;
		}
		const c = parseCommands(pathBody, counter, budget);
		if (c === null) {
			if (counter.over === null) bad = true;
			return;
		}
		if (c.length === 0 && !charge(counter, budget)) return;
		const a = parseAttrs(open);
		const path = { c };
		const w = intAttr(a.w);
		const h = intAttr(a.h);
		if (w !== null && w > 0) path.w = clampInt(w);
		if (h !== null && h > 0) path.h = clampInt(h);
		if (a.fill !== void 0 && PATH_FILLS.has(a.fill)) path.fill = a.fill;
		if (a.stroke === "0" || a.stroke === "false") path.stroke = false;
		paths.push(path);
	});
	if (counter.over !== null) return {
		ok: false,
		reason: counter.over
	};
	if (bad || paths.length === 0) return {
		ok: false,
		reason: "invalid"
	};
	return {
		ok: true,
		def: rect === void 0 ? {
			av,
			gd,
			paths
		} : {
			av,
			gd,
			rect,
			paths
		},
		units: guides + counter.n
	};
}
// src/core/drawing-text.ts
/** 文字の大きさの既定（pt。`sz` が無いとき） */
var DEFAULT_SIZE_PT = 11;
/** 文字の大きさの範囲（pt） */
var MIN_SIZE_PT = 1;
var MAX_SIZE_PT = 400;
/** 内側の余白の既定（EMU。左・上・右・下） */
var DEFAULT_INSETS = {
	l: 91440,
	t: 45720,
	r: 91440,
	b: 45720
};
/** 内側の余白の上限（EMU。2^31-1） */
var MAX_INSET = 2147483647;
var ALIGNS = /* @__PURE__ */ new Set([
	"l",
	"ctr",
	"r",
	"just",
	"dist"
]);
var VERTS = /* @__PURE__ */ new Set([
	"horz",
	"vert",
	"vert270",
	"wordArtVert",
	"eaVert",
	"mongolianVert",
	"wordArtVertRtl"
]);
/** 段落の直下の要素（ここに当たったら中を飛ばして次へ進む。`a:fld` は中に `a:pPr`・`a:rPr` を持つ） */
var PARA_TOKEN_RE = new RegExp(String.raw`<(${NS})(pPr|r|br|fld|endParaRPr)(?=[\s>/])[^<>]*>`, "g");
/** `a:t` の開きタグ */
var T_RE = new RegExp(String.raw`<(${NS})t(?=[\s>/])[^<>]*>`, "g");
/** タブ・改行以外の制御文字（C0・DEL・C1） */
var CONTROL_RE = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F]/g;
/** `a:rPr`・`a:endParaRPr` の `sz`（1/100pt）を pt にする。無い・不正は 11。1〜400 に丸める */
function sizeOf(tag) {
	const n = intAttr(tag === null ? void 0 : parseAttrs(tag).sz);
	if (n === null) return DEFAULT_SIZE_PT;
	return Math.min(Math.max(n / 100, MIN_SIZE_PT), MAX_SIZE_PT);
}
/** 余白（EMU）。無い・不正は既定値。0〜2^31-1 に丸める */
function insetOf(s, dflt) {
	const n = intAttr(s);
	return n === null ? dflt : Math.min(Math.max(n, 0), MAX_INSET);
}
/** 本文の末尾が高位サロゲートで切れていたら、その 1 文字を外す（文字を途中で切って壊さない） */
function trimLoneSurrogate(s) {
	const c = s.length === 0 ? 0 : s.charCodeAt(s.length - 1);
	return c >= 55296 && c <= 56319 ? s.slice(0, -1) : s;
}
/** 文字の枠（`a:bodyPr`）を読む。無ければすべて既定値 */
function readBodyPr(txBody) {
	const el = findElement(txBody, "bodyPr");
	const a = el === null ? parseAttrs("") : parseAttrs(el.open);
	const anchor = a.anchor === "ctr" || a.anchor === "b" ? a.anchor : "t";
	const vert = a.vert !== void 0 && VERTS.has(a.vert) ? a.vert : "horz";
	let fontScale = 1;
	const fit = el?.body == null ? null : findOpenTag(el.body, "normAutofit");
	if (fit !== null) {
		const f = intAttr(parseAttrs(fit).fontScale);
		if (f !== null) fontScale = Math.min(Math.max(f / 1e5, .01), 1);
	}
	return {
		anchor,
		insets: {
			l: insetOf(a.lIns, DEFAULT_INSETS.l),
			t: insetOf(a.tIns, DEFAULT_INSETS.t),
			r: insetOf(a.rIns, DEFAULT_INSETS.r),
			b: insetOf(a.bIns, DEFAULT_INSETS.b)
		},
		noWrap: a.wrap === "none",
		fontScale,
		vert
	};
}
/** 段落の書式（`a:pPr`）：揃えと、箇条書きの文字（番号の続きは `num` を更新して数える） */
function readParaProps(open, body, num) {
	const algn = parseAttrs(open).algn;
	const align = algn !== void 0 && ALIGNS.has(algn) ? algn : "l";
	let bullet = null;
	const none = body !== null && findOpenTag(body, "buNone") !== null;
	const auto = body === null || none ? null : findOpenTag(body, "buAutoNum");
	const chr = body === null || none || auto !== null ? null : findOpenTag(body, "buChar");
	if (auto !== null) {
		const a = parseAttrs(auto);
		const type = a.type ?? "arabicPeriod";
		const startAt = intAttr(a.startAt);
		const cur = num.cur;
		if (cur !== null && cur.type === type && (startAt === null || startAt === cur.start)) {
			bullet = `${cur.next}.`;
			cur.next++;
		} else {
			const start = startAt ?? 1;
			bullet = `${start}.`;
			num.cur = {
				type,
				start,
				next: start + 1
			};
		}
		return {
			align,
			bullet
		};
	}
	num.cur = null;
	if (chr !== null) bullet = Array.from((parseAttrs(chr).char ?? "").replace(CONTROL_RE, ""))[0] ?? null;
	return {
		align,
		bullet
	};
}
/** ラン・フィールド（`a:r`・`a:fld`）の中身から `a:t` の本文（実体参照・制御文字の処理前）を取り出す。無ければ空 */
function rawText(body) {
	if (body === null) return "";
	T_RE.lastIndex = 0;
	const m = T_RE.exec(body);
	if (m === null || m[0].endsWith("/>")) return "";
	const start = m.index + m[0].length;
	const close = body.indexOf(`</${m[1] ?? ""}t>`, start);
	return close < 0 ? "" : body.slice(start, close);
}
/** 本文に空でない `a:t` があるか（文字の無い図形で、上限のために文字を失ったかの判定。線形時間） */
function hasNonEmptyText(txBody) {
	T_RE.lastIndex = 0;
	for (let m = T_RE.exec(txBody); m !== null; m = T_RE.exec(txBody)) {
		if (m[0].endsWith("/>")) continue;
		const next = txBody[T_RE.lastIndex];
		if (next !== void 0 && next !== "<") {
			T_RE.lastIndex = 0;
			return true;
		}
	}
	return false;
}
/** 書式（`a:rPr`）を読んでランにする。`rPrBody` は `a:rPr` の中身（色の塗り）、`fallback` は図形の文字の既定の色（`fontRef`） */
function makeRun(rPr, rPrBody, text, br, fallback, theme) {
	const a = rPr === null ? parseAttrs("") : parseAttrs(rPr);
	return {
		br,
		text,
		sizePt: sizeOf(rPr),
		bold: boolAttr$1(a.b),
		italic: boolAttr$1(a.i),
		underline: a.u !== void 0 && a.u !== "none",
		strike: a.strike !== void 0 && a.strike !== "noStrike",
		color: readRunColor(rPrBody, fallback, theme)
	};
}
/**
* 図形（`xdr:sp`）の中身から文字（`xdr:txBody`）を読む。
* - `a:t` は `decodeEntities` で戻し、タブと改行以外の制御文字を除く。
* - 上限（文字数 `maxChars`・段落 `maxParagraphs`・ラン `maxRuns`・1 シートの残り `maxUnits`）を超えたら、文字は切って末尾に「…」を付け、段落とランは以降を捨てて `truncated` にする。
* - 空のラン（文字も `a:br` も無いもの）は一覧に入れず、ランの数にも数えない。
* - ランの文字（長さ 1 以上）が 1 つも無い図形は `text` を null にする（文字の無い図形）。このとき `truncated` は、本文に空でない `a:t` があって上限で失ったときだけ立てる。
* - ランの色（ISSUE-95）は `a:rPr` の塗り → `xdr:style` の `a:fontRef` の子の色 → テーマの `tx1`（`drawing-color.ts`）。
* @param spBody `xdr:sp` の中身
* @param theme ブックのテーマ（色の解決に使う。省くと既定の 12 色）
*/
function readShapeText(spBody, caps, theme = DEFAULT_THEME) {
	const tx = findElement(spBody, "txBody");
	if (tx?.body == null) return {
		text: null,
		chars: 0,
		units: 0,
		truncated: false
	};
	const txBody = tx.body;
	const fontColor = readFontColor(spBody, theme);
	const paragraphs = [];
	const num = { cur: null };
	let chars = 0;
	let runs = 0;
	let units = 0;
	let hasText = false;
	let truncated = false;
	let done = false;
	forEachElement(txBody, "p", (_open, pBody) => {
		if (done) return;
		if (paragraphs.length >= caps.maxParagraphs || units + 1 > caps.maxUnits) {
			truncated = true;
			done = true;
			return;
		}
		const para = {
			runs: [],
			align: "l",
			bullet: null,
			endSizePt: DEFAULT_SIZE_PT
		};
		paragraphs.push(para);
		units++;
		if (pBody === null) {
			num.cur = null;
			return;
		}
		let sawProps = false;
		for (let pos = 0; !done;) {
			PARA_TOKEN_RE.lastIndex = pos;
			const m = PARA_TOKEN_RE.exec(pBody);
			if (m === null) break;
			const open = m[0];
			const name = m[2];
			pos = m.index + open.length;
			let body = null;
			if (!open.endsWith("/>")) {
				const closeTag = `</${m[1] ?? ""}${name}>`;
				const close = pBody.indexOf(closeTag, pos);
				if (close < 0) break;
				body = pBody.slice(pos, close);
				pos = close + closeTag.length;
			}
			if (name === "pPr") {
				if (!sawProps) {
					sawProps = true;
					const p = readParaProps(open, body, num);
					para.align = p.align;
					para.bullet = p.bullet;
				}
				continue;
			}
			if (name === "endParaRPr") {
				para.endSizePt = sizeOf(open);
				continue;
			}
			const rEl = body === null ? null : findElement(body, "rPr");
			const rPr = rEl?.open ?? null;
			const rPrBody = rEl?.body ?? null;
			let text = "";
			if (name !== "br") {
				text = decodeEntities(rawText(body)).replace(CONTROL_RE, "");
				if (text.length === 0) continue;
			}
			if (runs >= caps.maxRuns || units + 1 > caps.maxUnits) {
				truncated = true;
				done = true;
				break;
			}
			if (name === "br") {
				runs++;
				units++;
				para.runs.push(makeRun(rPr, rPrBody, "", true, fontColor, theme));
				continue;
			}
			const room = Math.min(caps.maxChars - chars, caps.maxUnits - units - 1);
			if (text.length > room) {
				truncated = true;
				done = true;
				text = trimLoneSurrogate(text.slice(0, Math.max(room, 0)));
				if (chars + text.length === 0) break;
				text += "…";
			}
			chars += text.length;
			units += 1 + text.length;
			hasText = true;
			runs++;
			para.runs.push(makeRun(rPr, rPrBody, text, false, fontColor, theme));
		}
		if (!sawProps) num.cur = null;
	});
	if (!hasText) return {
		text: null,
		chars: 0,
		units: 0,
		truncated: truncated && hasNonEmptyText(txBody)
	};
	return {
		text: {
			paragraphs,
			...readBodyPr(txBody)
		},
		chars,
		units,
		truncated
	};
}
// src/core/workbook.ts
/** `rels` を Id → 関係 の Map にする（キーを Map に入れて `__proto__` を防ぐ） */
function parseRels(xml) {
	const map = /* @__PURE__ */ new Map();
	for (const m of xml.matchAll(new RegExp(String.raw`<${NS}Relationship\b[^<>]*>`, "g"))) {
		const a = parseAttrs(m[0]);
		if (a.Id !== void 0 && a.Target !== void 0) map.set(a.Id, {
			type: a.Type ?? "",
			target: a.Target,
			external: a.TargetMode?.toLowerCase() === "external"
		});
	}
	return map;
}
/** 関係の Target を zip のパスにする。`/` 始まりは絶対、それ以外は baseDir からの相対（`..` を解く） */
function resolveTarget(baseDir, target) {
	const full = target.startsWith("/") ? target.slice(1) : baseDir + target;
	const out = [];
	for (const seg of full.split("/")) {
		if (seg === "" || seg === ".") continue;
		if (seg === "..") out.pop();
		else out.push(seg);
	}
	return out.join("/");
}
/** パスのフォルダ部分（末尾の / を含む。無ければ空） */
function dirOf(path) {
	const i = path.lastIndexOf("/");
	return i < 0 ? "" : path.slice(0, i + 1);
}
/** パスのファイル名の部分 */
function baseOf(path) {
	const i = path.lastIndexOf("/");
	return i < 0 ? path : path.slice(i + 1);
}
/** 型（URL の最後の語）が一致するか */
var typeIs = (type, last) => type === last || type.endsWith("/" + last);
/**
* workbook.xml を読む。
* - 本体の場所は `_rels/.rels` の officeDocument から引く（無ければ `xl/workbook.xml`）。
* - 読めない形式（Strict・xlsb）は unsupported。workbook が無ければ broken。
*/
async function loadWorkbook(pkg, limits) {
	const small = limits.maxSmallPartBytes;
	let wbPath = "xl/workbook.xml";
	const rootRels = await pkg.readText("_rels/.rels", small);
	if (rootRels !== null) {
		for (const rel of parseRels(rootRels).values()) if (typeIs(rel.type, "officeDocument")) wbPath = resolveTarget("", rel.target);
	}
	const wbXml = await pkg.readText(wbPath, small);
	if (wbXml === null) throw new XlsxError(pkg.has("xl/workbook.bin") ? "unsupported" : "broken", "workbook");
	if (looksStrict(wbXml)) throw new XlsxError("unsupported", "strict");
	const wbDir = dirOf(wbPath);
	const relsXml = await pkg.readText(`${wbDir}_rels/${baseOf(wbPath)}.rels`, small);
	const rels = relsXml === null ? /* @__PURE__ */ new Map() : parseRels(relsXml);
	const date1904 = (() => {
		const m = new RegExp(String.raw`<${NS}workbookPr\b[^<>]*>`).exec(wbXml);
		if (!m) return false;
		const v = parseAttrs(m[0]).date1904;
		return v === "1" || v === "true";
	})();
	const sheets = [];
	for (const m of wbXml.matchAll(new RegExp(String.raw`<${NS}sheet\b[^<>]*>`, "g"))) {
		if (sheets.length >= limits.maxSheets) throw new XlsxError("too-large", "sheets");
		const a = parseAttrs(m[0]);
		const rel = rels.get(attrByLocalName(a, "id") ?? "");
		sheets.push({
			index: sheets.length + 1,
			name: a.name ?? "",
			hidden: a.state === "hidden" || a.state === "veryHidden",
			path: rel ? resolveTarget(wbDir, rel.target) : null,
			isWorksheet: rel ? typeIs(rel.type, "worksheet") : true
		});
	}
	let sharedStringsPath = null;
	let stylesPath = null;
	let themePath = null;
	for (const rel of rels.values()) if (typeIs(rel.type, "sharedStrings")) sharedStringsPath = resolveTarget(wbDir, rel.target);
	else if (typeIs(rel.type, "styles")) stylesPath = resolveTarget(wbDir, rel.target);
	else if (typeIs(rel.type, "theme") && !rel.external) themePath = resolveTarget(wbDir, rel.target);
	return {
		sheets,
		date1904,
		sharedStringsPath,
		stylesPath,
		themePath
	};
}
// src/core/drawing.ts
/**
* シートの図（画像・グラフ・図形）の読み取り：`<drawing r:id>` → シートの rels → drawingN.xml → drawing の rels → media。
* DOMParser に依存せず、正規表現と `indexOf` で線形時間に走査する。
* 開きタグは `[^<>]*>`（`>` が無くても次の `<` で打ち切る）、閉じタグは `indexOf`（`findElement`）か、入れ子を数える一度きりの走査（`findBalanced`）で探す。
* 閉じの無い要素や `>` の無い開きタグを大量に並べた細工でも、時間は XML の大きさに比例する。
* 外部取得をしない：`TargetMode="External"` の関係と `a:blip` の `r:link` は辿らない。
* 図形・テキストボックスの文字（`drawing-text.ts`）と、線・矢印（`xdr:cxnSp`・線の preset の `xdr:sp`。種類 `line`）の形・回転・反転・枠線と塗りの有無（`drawing-shape.ts`）も読む（ADR-0027）。
* 線でない図形の形（`a:prstGeom` の調整値・`a:custGeom` の定義。`drawing-geom.ts`）と枠線の太さ・線種も、文字の有無を問わず読む（ADR-0028）。
* 図形の塗り・枠線、線、文字のランの色（出どころの順・代表色）は `drawing-color.ts` で解き、ブックのテーマ（`loadSheetDrawings` の `theme`）を使う（ISSUE-95）。
* 線は図とは別に数える。画像のトリミング（`srcRect`）・回転・反転は読まない（第 1 段）。
* グループ（`xdr:grpSp`）は中の子（`sp`・`cxnSp`・`pic`・`graphicFrame`・入れ子のグループ）を展開して一覧に足し、グループ自身は足さない。
* 子はトップレベルと同じ `addDrawing` で読み、グループの子の座標系（`SheetDrawing.group`）を持たせる。写像（px）は表示側が行う（ISSUE-89）。
* 図の名前・画像のパスはログに出さない（console を使わない）。
*/
/** EMU（ずれ・大きさ）の上限（2^31-1）。負・これを超える値は丸める */
var MAX_EMU = 2147483647;
/** mc:AlternateContent を辿る深さの上限（入れ子の細工で再帰が深くならないように） */
var MAX_AC_DEPTH = 4;
/** 図の無い結果（共有しない：呼び出し側が書き換えても他に響かないよう毎回作る） */
var empty = () => ({
	drawings: [],
	notices: []
});
/** アンカー（位置の入れ物）の要素名 */
var ANCHOR_RE = new RegExp(String.raw`<(${NS})(twoCellAnchor|oneCellAnchor|absoluteAnchor|AlternateContent)\b[^<>]*>`, "g");
/** アンカーの中の図の要素（最初に現れるものがその図） */
var CONTENT_RE = new RegExp(String.raw`<(${NS})(sp|grpSp|graphicFrame|cxnSp|pic|contentPart|AlternateContent)\b[^<>]*>`);
/** グループの直下の子（図の要素。入れ子のグループ・`mc:AlternateContent` を含む）の開きタグ */
var CHILD_RE = new RegExp(String.raw`<(${NS})(sp|grpSp|graphicFrame|cxnSp|pic|AlternateContent)\b[^<>]*>`, "g");
/** `xdr:pic` の開きタグ */
var PIC_RE = new RegExp(String.raw`<${NS}pic\b`);
/**
* シートの XML（`<sheetData>` の後ろの部分）から `<drawing r:id>` の関係の Id を取り出す。無ければ null。
* `legacyDrawing`（コメントの吹き出し・フォームコントロール）は別の要素なので拾わない。
*/
function findDrawingRelId(sheetXml) {
	const tail = sheetXml.lastIndexOf("sheetData>");
	const el = findElement(sheetXml, "drawing", tail < 0 ? 0 : tail);
	if (el === null) return null;
	return attrByLocalName(parseAttrs(el.open), "id") ?? null;
}
/**
* `<name>…</name>` を、同じ名前の入れ子を数えて線形時間で切り出す（`xml.ts` の `findElement` は最初の閉じで切るため、入れ子がありうる要素用）。
* `from` は開きタグの先頭の位置。閉じが無いときは body を null・end を xml の長さにして走査を打ち切らせる。
*/
function findBalanced(xml, name, from) {
	const re = new RegExp(String.raw`<(/?)(${NS})${name}\b[^<>]*>`, "g");
	re.lastIndex = from;
	const first = re.exec(xml);
	if (first === null) return null;
	if (first[0].endsWith("/>")) return {
		body: null,
		end: re.lastIndex
	};
	const bodyStart = re.lastIndex;
	let depth = 1;
	for (let m = re.exec(xml); m !== null; m = re.exec(xml)) if (m[1] === "/") {
		depth--;
		if (depth === 0) return {
			body: xml.slice(bodyStart, m.index),
			end: re.lastIndex
		};
	} else if (!m[0].endsWith("/>")) depth++;
	return {
		body: null,
		end: xml.length
	};
}
/**
* `mc:AlternateContent` の中身から、採る枝を決める：`xdr:pic` を含む枝（Choice・Fallback を問わず）、
* 無ければ（図の要素を包むアンカーの中の AlternateContent のとき）`graphicFrame` を持つ Choice（`graphicData uri` は問わない。
* 新しい種類のグラフ・スライサー・タイムラインの Fallback にある「お使いのバージョンでは利用できません」という説明文入りの図形を出さない）、
* 無ければ Fallback、Fallback も無ければ最初の Choice。採る枝が無ければ null。
* @param inAnchor アンカーの中の AlternateContent か（アンカー自身を枝に持つ外側のものは false。スライサーの Fallback のアンカーを今のまま採る）
*/
function pickBranch(body, inAnchor) {
	const branches = [];
	let fallback = null;
	const re = new RegExp(String.raw`<${NS}(Choice|Fallback)\b[^<>]*>`, "g");
	for (let pos = 0;;) {
		re.lastIndex = pos;
		const m = re.exec(body);
		if (m === null) break;
		const el = findBalanced(body, m[1], m.index);
		if (el === null) break;
		if (el.body !== null) {
			if (m[1] === "Choice") branches.push(el.body);
			else fallback ??= el.body;
		}
		pos = el.end;
	}
	const pic = (fallback === null ? branches : [...branches, fallback]).find((b) => PIC_RE.test(b));
	if (pic !== void 0) return pic;
	if (inAnchor) {
		const frame = branches.find((b) => hasGraphicFrame(b));
		if (frame !== void 0) return frame;
	}
	return fallback ?? branches[0] ?? null;
}
/** 枝が `graphicFrame`（グラフ・新しい種類のグラフ・スライサー・タイムラインなど。`graphicData uri` は問わない）を含むか */
function hasGraphicFrame(branch) {
	return findOpenTag(branch, "graphicFrame") !== null;
}
/** 整数の文字（前後の空白可）を数にする。整数でない・桁が多すぎるなら null */
function toInt(s) {
	if (s === void 0) return null;
	const t = s.trim();
	if (!/^-?\d{1,15}$/.test(t)) return null;
	return Number(t);
}
/** EMU を 0〜MAX_EMU に丸める（負は 0） */
var clampEmu = (n) => Math.min(Math.max(n, 0), MAX_EMU);
/** `<xdr:from>` / `<xdr:to>` の中身から 1 点を読む。番号が無い・不正・範囲外なら null（ずれが無ければ 0） */
function readPoint(head, tag) {
	const body = findElement(head, tag)?.body;
	if (body === null || body === void 0) return null;
	const num = (name) => findElement(body, name)?.body?.trim();
	const col = toInt(num("col"));
	const row = toInt(num("row"));
	if (col === null || row === null || col < 0 || col >= 16384 || row < 0 || row >= 1048576) return null;
	const off = (name) => {
		const s = num(name);
		return s === void 0 ? 0 : toInt(s) ?? null;
	};
	const colOff = off("colOff");
	const rowOff = off("rowOff");
	if (colOff === null || rowOff === null) return null;
	return {
		col,
		colOff: clampEmu(colOff),
		row,
		rowOff: clampEmu(rowOff)
	};
}
/** `<xdr:ext cx cy>` / `<xdr:pos x y>` の 2 つの整数の属性を読む。無い・不正なら null */
function readPair(head, tag, a, b) {
	const open = findElement(head, tag)?.open;
	if (open === void 0) return null;
	const attrs = parseAttrs(open);
	const x = toInt(attrs[a]);
	const y = toInt(attrs[b]);
	return x === null || y === null ? null : [clampEmu(x), clampEmu(y)];
}
/** アンカーの種類と、図の要素より前の部分（from・to・pos・ext）から位置を読む。読めなければ null */
function readAnchor(kind, head) {
	if (kind === "twoCellAnchor") {
		const from = readPoint(head, "from");
		const to = readPoint(head, "to");
		return from && to ? {
			type: "twoCell",
			from,
			to
		} : null;
	}
	if (kind === "oneCellAnchor") {
		const from = readPoint(head, "from");
		const ext = readPair(head, "ext", "cx", "cy");
		return from && ext ? {
			type: "oneCell",
			from,
			cx: ext[0],
			cy: ext[1]
		} : null;
	}
	const pos = readPair(head, "pos", "x", "y");
	const ext = readPair(head, "ext", "cx", "cy");
	return pos && ext ? {
		type: "absolute",
		x: pos[0],
		y: pos[1],
		cx: ext[0],
		cy: ext[1]
	} : null;
}
/** 画像の参照（`a:blip r:embed`）を zip のパスにする。外部・r:embed 無し・zip に無いものは null（`r:link` は読まない） */
function resolveMedia(ctx, blip) {
	const id = blip === null ? void 0 : attrByLocalName(blip, "embed");
	if (id === void 0 || id === "") return null;
	const rel = ctx.rels.get(id);
	if (rel === void 0 || rel.external) return null;
	const path = resolveTarget(ctx.dir, rel.target);
	return ctx.pkg.has(path) ? path : null;
}
/**
* 図形（`xdr:sp`）・線（`xdr:cxnSp`）の中身を切り出し、見た目（回転・反転・枠線と塗り・形）を読む。
* `at` は要素の開きタグの位置。閉じが無いなどで中身が取れなければ null。
*/
function readShape(body, at, name) {
	const spBody = findElement(body, name, at)?.body ?? null;
	return spBody === null ? null : {
		spBody,
		props: readShapeProps(spBody)
	};
}
/**
* アンカーの中の図の要素 1 つを一覧に足す。`at` は図の要素の開きタグの位置、`name` はその要素名。
* 隠した図（`cNvPr hidden`）と、見えない線は足さない。
* 線・矢印（`xdr:cxnSp`・線の preset の `xdr:sp`）は種類 `line` で、図とは別の上限（`maxLines`）で数える。
* 数が上限を超えたら、その種類だけ足さず（`figuresFull`・`linesFull`）、もう一方は足し続ける。
* 図形（`textbox`・`shape`）は、文字の有無を問わず回転・反転・枠線と塗りの有無・枠線の太さと線種・形を持つ。文字は文字のある図形だけが持つ。
*/
function addDrawing(ctx, anchor, body, at, name, depth, levels = null) {
	if (name === "AlternateContent") {
		const acBody = findBalanced(body, "AlternateContent", at)?.body ?? null;
		const branch = acBody === null || depth >= MAX_AC_DEPTH ? null : pickBranch(acBody, true);
		if (branch === null) return;
		const m = CONTENT_RE.exec(branch);
		if (m !== null) addDrawing(ctx, anchor, branch, m.index, m[2], depth + 1, levels);
		return;
	}
	const grpBody = name === "grpSp" ? findBalanced(body, "grpSp", at)?.body ?? null : null;
	const head = grpBody === null ? null : readGroupHead(grpBody);
	const cNvPr = head !== null ? head.cNvPr : findOpenTag(body, "cNvPr", at);
	const attrs = cNvPr === null ? null : parseAttrs(cNvPr);
	if (attrs?.hidden === "1" || attrs?.hidden === "true") return;
	if (grpBody !== null && head !== null && expandGroup(ctx, anchor, grpBody, head, depth, levels)) return;
	let kind;
	let mediaPath = null;
	let shape = null;
	if (name === "pic") {
		kind = "image";
		const blip = findOpenTag(body, "blip", at);
		mediaPath = resolveMedia(ctx, blip === null ? null : parseAttrs(blip));
	} else if (name === "graphicFrame") kind = (parseAttrs(findOpenTag(body, "graphicData", at) ?? "").uri ?? "").endsWith("/diagram") ? "smartart" : "chart";
	else if (name === "grpSp") kind = "group";
	else if (name === "cxnSp") {
		shape = readShape(body, at, "cxnSp");
		if (shape === null) return;
		kind = "line";
	} else if (name === "sp") {
		shape = readShape(body, at, "sp");
		if (shape !== null && isLinePreset(shape.props.prst)) kind = "line";
		else {
			const tx = parseAttrs(findOpenTag(body, "cNvSpPr", at) ?? "").txBox;
			kind = tx === "1" || tx === "true" ? "textbox" : "shape";
		}
	} else kind = "shape";
	let group;
	if (levels !== null) {
		const xfrm = childXfrm(body, at, name, shape, head);
		if (xfrm === null) return;
		group = {
			levels,
			xfrm
		};
	}
	const drawing = {
		anchor,
		kind,
		name: (attrs?.name ?? "").slice(0, MAX_DRAWING_TEXT),
		descr: (attrs?.descr ?? "").slice(0, MAX_DRAWING_TEXT),
		mediaPath
	};
	if (group !== void 0) drawing.group = group;
	if (kind === "line") {
		const props = shape.props;
		if (!props.hasOutline) return;
		if (ctx.lines >= ctx.limits.maxLines) {
			ctx.linesFull = true;
			return;
		}
		ctx.lines++;
		drawing.rot = props.rot;
		drawing.flipH = props.flipH;
		drawing.flipV = props.flipV;
		drawing.line = readLine(props);
		drawing.line.color = readOutlineColor(props, ctx.theme);
		ctx.out.push(drawing);
		return;
	}
	if (ctx.figures >= ctx.limits.maxDrawings) {
		ctx.figuresFull = true;
		return;
	}
	ctx.figures++;
	if (shape !== null) {
		const props = shape.props;
		drawing.rot = props.rot;
		drawing.flipH = props.flipH;
		drawing.flipV = props.flipV;
		if (group === void 0 && props.rot !== 0) {
			const xf = readXfrm(shape.spBody);
			if (xf !== null && xf.cx > 0 && xf.cy > 0) drawing.xfrmExt = {
				cx: xf.cx,
				cy: xf.cy
			};
		}
		drawing.hasOutline = props.hasOutline;
		drawing.hasFill = props.hasFill;
		drawing.outline = readStroke(props);
		if (props.hasFill) drawing.fillColor = readFillColor(props, ctx.theme);
		if (props.hasOutline) drawing.outlineColor = readOutlineColor(props, ctx.theme);
		const geom = readShapeGeom(ctx, props);
		if (geom !== void 0) drawing.geom = geom;
		const room = ctx.limits.maxSheetTextChars - ctx.textUnits;
		const read = readShapeText(shape.spBody, {
			maxChars: ctx.limits.maxShapeTextChars,
			maxUnits: Math.max(room, 0),
			maxParagraphs: ctx.limits.maxShapeParagraphs,
			maxRuns: ctx.limits.maxShapeRuns
		}, ctx.theme);
		if (read.truncated) ctx.textTruncated = true;
		if (read.text !== null) {
			ctx.textUnits += read.units;
			drawing.text = read.text;
		}
	}
	ctx.out.push(drawing);
}
/**
* 図形の形（`a:prstGeom`＋調整値、または `a:custGeom` の定義）を読む。形が無い・自由図形の字句の誤り・上限超えは undefined。
* 自由図形の上限（1 図形の guide・命令、1 シートの命令の合計）を超えたら `shape-simplified` を立て、シート合計を超えたら以後の自由図形は読まない。
* 図形 1 つ（`sp`）の読み取りで、トップレベルの図形もグループの子も同じ `addDrawing` を通る。
*/
function readShapeGeom(ctx, props) {
	if (props.prst !== null && props.prst !== "") return {
		type: "preset",
		prst: props.prst,
		adj: readPresetAdjustments(props.geomBody, ctx.limits.maxShapeAdjustments)
	};
	if (props.custGeom === null) return void 0;
	if (ctx.geomFull) {
		ctx.shapeSimplified = true;
		return;
	}
	const read = readCustGeom(props.custGeom.body, {
		maxGuides: ctx.limits.maxGeomGuides,
		maxCommands: ctx.limits.maxGeomCommands,
		sheetRoom: Math.max(ctx.limits.maxSheetGeomUnits - ctx.geomUnits, 0)
	});
	if (read.ok) {
		ctx.geomUnits += read.units;
		return {
			type: "custom",
			def: read.def
		};
	}
	if (read.reason !== "invalid") {
		ctx.shapeSimplified = true;
		if (read.reason === "sheet-limit") ctx.geomFull = true;
	}
}
/**
* グループの中の図の `a:xfrm`（グループの子の座標系での位置と大きさ）。`at` は図の要素の開きタグの位置。
* `sp`・`cxnSp` は `readShape` が切り出した中身から、`pic`・`graphicFrame` はその要素の中身から、グループは自分の `grpSpPr`（`head`）から読む。読めなければ null。
*/
function childXfrm(body, at, name, shape, head) {
	if (name === "sp" || name === "cxnSp") return shape === null ? null : readXfrm(shape.spBody);
	if (name === "grpSp") {
		const pr = head?.pr?.body;
		return pr === null || pr === void 0 ? null : readXfrm(pr);
	}
	if (name === "pic" || name === "graphicFrame") {
		const el = findElement(body, name, at)?.body;
		return el === null || el === void 0 ? null : readXfrm(el);
	}
	return null;
}
/**
* グループの中身を展開して、子を一覧に足す。展開したら true。
* 展開できないとき（深さが上限を超える・座標系が読めない・`chExt` が 0）は false で、何も足さない（呼び出し側が破線枠にする）。
* 子は `levels`（外側のグループから順の座標系）に、このグループの座標系を足して `addDrawing` へ渡す。
* @param grp グループの中身（開きタグと閉じタグの間）
* @param head グループ自身の cNvPr・grpSpPr（`readGroupHead`）
* @param levels このグループを囲むグループの座標系（トップレベルのグループは null）
*/
function expandGroup(ctx, anchor, grp, head, depth, levels) {
	if ((levels?.length ?? 0) >= ctx.limits.maxGroupDepth) return false;
	const frame = readGroupFrame(head.pr, levels !== null);
	if (frame === null) return false;
	scanGroupChildren(ctx, anchor, grp, frame.end, depth, [...levels ?? [], {
		ch: frame.ch,
		own: frame.own
	}]);
	return true;
}
/**
* グループの中身（`xml`）の `from` 以降にある直下の子を順に `addDrawing` へ渡す。
* 子の要素 1 つ分だけを切り出して渡す（図の読み取りが、後ろの兄弟の要素まで読まないように）。
* 子の終わりを探したあと、その中身は走査し直さない（入れ子のグループは `expandGroup` が自分の中身だけを走査する）。
*/
function scanGroupChildren(ctx, anchor, xml, from, depth, levels) {
	let pos = from;
	while (!(ctx.figuresFull && ctx.linesFull)) {
		CHILD_RE.lastIndex = pos;
		const m = CHILD_RE.exec(xml);
		if (m === null) return;
		const name = m[2];
		const el = name === "grpSp" || name === "AlternateContent" ? findBalanced(xml, name, m.index) : findElement(xml, name, m.index);
		if (el === null) return;
		pos = el.end;
		if (el.body === null) continue;
		addDrawing(ctx, anchor, xml.slice(m.index, el.end), 0, name, depth, levels);
	}
}
/** drawing XML（またはその一部）のアンカーを順に読み、図を足す。トップレベルの `mc:AlternateContent` は採る枝の中を読む */
function scanAnchors(ctx, xml, depth) {
	let pos = 0;
	while (!(ctx.figuresFull && ctx.linesFull)) {
		ANCHOR_RE.lastIndex = pos;
		const m = ANCHOR_RE.exec(xml);
		if (m === null) return;
		const name = m[2];
		if (name === "AlternateContent") {
			const ac = findBalanced(xml, name, m.index);
			if (ac === null) return;
			pos = ac.end;
			if (ac.body !== null && depth < MAX_AC_DEPTH) {
				const branch = pickBranch(ac.body, false);
				if (branch !== null) scanAnchors(ctx, branch, depth + 1);
			}
			continue;
		}
		const el = findElement(xml, name, m.index);
		if (el === null) return;
		pos = el.end;
		if (el.body === null) continue;
		const content = CONTENT_RE.exec(el.body);
		if (content === null) continue;
		const anchor = readAnchor(name, el.body.slice(0, content.index));
		if (anchor === null) continue;
		addDrawing(ctx, anchor, el.body, content.index, content[2], depth);
	}
}
/**
* シートの図を読む。`relId` はシートの `<drawing r:id>`。
* 上限超え・読み取りの失敗は例外にせず、通知（`notices`）で返す。図の無いシート・zip に解決できない参照は空の一覧。
* @param sheetPath シート XML の zip の中のパス（rels の場所と Target の基準）
*/
async function loadSheetDrawings(pkg, sheetPath, relId, limits, theme = DEFAULT_THEME) {
	const small = limits.maxSmallPartBytes;
	try {
		const sheetRels = await pkg.readText(`${dirOf(sheetPath)}_rels/${baseOf(sheetPath)}.rels`, small);
		if (sheetRels === null) return empty();
		const rel = parseRels(sheetRels).get(relId);
		if (rel === void 0 || rel.external) return empty();
		const drawingPath = resolveTarget(dirOf(sheetPath), rel.target);
		if (!pkg.has(drawingPath)) return empty();
		const xml = await pkg.readText(drawingPath, small);
		if (xml === null) return empty();
		const relsXml = await pkg.readText(`${dirOf(drawingPath)}_rels/${baseOf(drawingPath)}.rels`, small);
		const ctx = {
			pkg,
			dir: dirOf(drawingPath),
			rels: relsXml === null ? /* @__PURE__ */ new Map() : parseRels(relsXml),
			limits,
			out: [],
			figures: 0,
			lines: 0,
			figuresFull: false,
			linesFull: false,
			textUnits: 0,
			textTruncated: false,
			theme,
			geomUnits: 0,
			geomFull: false,
			shapeSimplified: false
		};
		scanAnchors(ctx, xml, 0);
		const notices = [];
		if (ctx.figuresFull || ctx.linesFull) notices.push("too-many");
		if (ctx.textTruncated) notices.push("text-truncated");
		if (ctx.shapeSimplified) notices.push("shape-simplified");
		return {
			drawings: ctx.out,
			notices
		};
	} catch (e) {
		return {
			drawings: [],
			notices: [e instanceof XlsxError && e.kind === "too-large" ? "part-too-large" : "broken"]
		};
	}
}
// src/core/format.ts
/**
* 数値の表示形式を当てて、Excel で見える文字にする。
* 書式コードの解釈は `ssf` 0.11.2 に任せ、日本語環境に合わせた組み込みの書式の対応表だけ自前で持つ。
* 和暦（`ggge`）は ssf にも無いので、1 次では対応しない（ADR-0025「悪い点」）。
*/
var ssf = ssfModule.default ?? ssfModule;
/**
* 組み込みの表示形式（numFmtId → 書式コード）。
* ssf の既定は米国式（14 は `m/d/yy`・22 は `m/d/yy h:mm`）なので、日本語環境の Excel の表示に合わせて上書きする。
* 27〜36・50〜58 は日本語環境では和暦だが、1 次は西暦で代用する（和暦は範囲外）。
*/
var BUILTIN_FORMATS = /* @__PURE__ */ new Map([
	[1, "0"],
	[2, "0.00"],
	[3, "#,##0"],
	[4, "#,##0.00"],
	[9, "0%"],
	[10, "0.00%"],
	[11, "0.00E+00"],
	[12, "# ?/?"],
	[13, "# ??/??"],
	[14, "yyyy/m/d"],
	[15, "d-mmm-yy"],
	[16, "d-mmm"],
	[17, "mmm-yy"],
	[18, "h:mm AM/PM"],
	[19, "h:mm:ss AM/PM"],
	[20, "h:mm"],
	[21, "h:mm:ss"],
	[22, "yyyy/m/d h:mm"],
	[27, "yyyy/m/d"],
	[28, "yyyy/m/d"],
	[29, "yyyy/m/d"],
	[30, "yyyy/m/d"],
	[31, "yyyy/m/d"],
	[32, "h\"時\"mm\"分\""],
	[33, "h\"時\"mm\"分\"ss\"秒\""],
	[34, "AM/PM h\"時\"mm\"分\""],
	[35, "AM/PM h\"時\"mm\"分\"ss\"秒\""],
	[36, "yyyy/m/d"],
	[37, "#,##0_);(#,##0)"],
	[38, "#,##0_);[Red](#,##0)"],
	[39, "#,##0.00_);(#,##0.00)"],
	[40, "#,##0.00_);[Red](#,##0.00)"],
	[45, "mm:ss"],
	[46, "[h]:mm:ss"],
	[47, "mm:ss.0"],
	[48, "##0.0E+0"],
	[49, "@"],
	[50, "yyyy/m/d"],
	[51, "yyyy/m/d"],
	[52, "yyyy/m/d"],
	[53, "yyyy/m/d"],
	[54, "yyyy/m/d"],
	[55, "yyyy/m/d"],
	[56, "yyyy/m/d"],
	[57, "yyyy/m/d"],
	[58, "yyyy/m/d"]
]);
/** 組み込みの表示形式の書式コード。無い番号（0 を含む）は General */
function builtinFormatCode(id) {
	return BUILTIN_FORMATS.get(id) ?? "General";
}
/** 書式コードが日付・時刻か */
function isDateCode(code) {
	try {
		return ssf.is_date(code);
	} catch {
		return false;
	}
}
/** General の表示。整数は桁を落とさず、それ以外は ssf（Excel と同じ有効桁）に任せる */
function formatGeneral(v, date1904) {
	if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
	try {
		return ssf.format("General", v, { date1904 });
	} catch {
		return String(v);
	}
}
/**
* 表示形式を当てる部品を作る。
* 書式コードは cellXfs の位置ごとに 1 回だけ引く（独自の書式 → 組み込みの対応表 → General の順）。
* ssf が解釈できない書式は General の表示に落とす（壊れずに数字を見せる）。
* 書式コードは非信頼の styles.xml から来るため、長さが上限（MAX_FORMAT_CODE_LENGTH）を超えるものは
* ssf に渡さず General として値だけを出す（ReDoS 対策）。
*/
function createFormatter(styles, date1904) {
	const codes = styles.xfNumFmtIds.map((id) => {
		const code = styles.customFormats.get(id) ?? builtinFormatCode(id);
		return code.length > 255 ? "General" : code;
	});
	const codeOf = (xf) => codes[xf] ?? "General";
	return {
		number(value, xf) {
			const code = codeOf(xf);
			if (code === "General" || code === "@") return formatGeneral(value, date1904);
			try {
				return ssf.format(code, value, { date1904 });
			} catch {
				return formatGeneral(value, date1904);
			}
		},
		date(serial, xf, hasTime) {
			let code = codeOf(xf);
			if (!isDateCode(code)) code = hasTime ? "yyyy/m/d h:mm:ss" : "yyyy/m/d";
			try {
				return ssf.format(code, serial, { date1904 });
			} catch {
				return formatGeneral(serial, date1904);
			}
		}
	};
}
/**
* ISO 8601 の日付・日時（`2023-03-15`・`2023-03-15T18:00:00`）を Excel のシリアル値にする。
* 読めなければ null。時刻の無い日付は hasTime=false。タイムゾーンの印は無視する。
*/
function isoToSerial(iso, date1904) {
	const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?/.exec(iso.trim());
	if (!m) return null;
	const days = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) / 864e5;
	if (!Number.isFinite(days)) return null;
	const hasTime = m[4] !== void 0;
	const frac = hasTime ? (Number(m[4]) * 3600 + Number(m[5]) * 60 + Number(m[6] ?? 0)) / 86400 : 0;
	let serial = days + (date1904 ? 24107 : 25569);
	if (!date1904 && serial < 61) serial -= 1;
	return {
		serial: serial + frac,
		hasTime
	};
}
// src/core/image.ts
/** 描ける形式 → MIME */
var MIME = Object.freeze({
	png: "image/png",
	jpeg: "image/jpeg",
	gif: "image/gif",
	bmp: "image/bmp",
	webp: "image/webp"
});
/** 判定できなかった（描けない）結果 */
var UNKNOWN = Object.freeze({
	ok: false,
	label: null
});
/** 先頭から ascii の文字列が並んでいるか（範囲外は false） */
function startsWith(b, at, s) {
	if (at + s.length > b.length) return false;
	for (let i = 0; i < s.length; i++) if (b[at + i] !== s.charCodeAt(i)) return false;
	return true;
}
/** 縦横が 1 以上の整数なら結果にする（0・負・非整数は不正＝描けない） */
function dims(format, width, height) {
	if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) return UNKNOWN;
	return {
		ok: true,
		format,
		width,
		height
	};
}
/** JPEG のフレーム見出し（SOF0〜SOF15。DHT=C4・JPG=C8・DAC=CC は別物） */
var isSof = (m) => m >= 192 && m <= 207 && m !== 196 && m !== 200 && m !== 204;
/** JPEG：マーカーのセグメントを長さで飛ばして SOF を探す（SOS・EOI まで。ループごとに必ず進む） */
function sniffJpeg(b) {
	const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
	let i = 2;
	while (i + 4 <= b.length) {
		if (b[i] !== 255) return UNKNOWN;
		let m = b[i + 1];
		while (m === 255 && i + 2 < b.length) {
			i++;
			m = b[i + 1];
		}
		if (m === 216 || m === 1 || m >= 208 && m <= 215) {
			i += 2;
			continue;
		}
		if (m === 217 || m === 218) return UNKNOWN;
		if (m === 255 || i + 4 > b.length) return UNKNOWN;
		const segLen = dv.getUint16(i + 2);
		if (segLen < 2) return UNKNOWN;
		if (isSof(m)) {
			if (i + 9 > b.length) return UNKNOWN;
			return dims("jpeg", dv.getUint16(i + 7), dv.getUint16(i + 5));
		}
		i += 2 + segLen;
	}
	return UNKNOWN;
}
/** BMP の DIB 見出しの大きさ（旧形式 12・INFO 40・V2 52・V3 56・OS/2 64・V4 108・V5 124） */
var BMP_DIB_SIZES = /* @__PURE__ */ new Set([
	12,
	40,
	52,
	56,
	64,
	108,
	124
]);
/**
* GIF：論理画面の幅・高さ（リトルエンディアン）と、最初の Image Descriptor（0x2C）の幅・高さを読み、それぞれ大きい方を返す。
* 論理画面を小さく偽ってフレームだけ巨大にした GIF を、上限判定で見逃さないため。
* 拡張ブロック（0x21）はサブブロックの長さで読み飛ばす。どの読み出しも長さを確かめ、ループごとに必ず進む。
*/
function sniffGif(b) {
	if (b.length < 10) return UNKNOWN;
	const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
	let w = dv.getUint16(6, true);
	let h = dv.getUint16(8, true);
	if (b.length >= 13) {
		let i = 13;
		if ((b[10] & 128) !== 0) i += 3 * 2 ** ((b[10] & 7) + 1);
		while (i < b.length) {
			const block = b[i];
			if (block === 33) {
				i += 2;
				while (i < b.length) {
					const n = b[i];
					i += 1 + n;
					if (n === 0) break;
				}
			} else {
				if (block === 44 && i + 9 <= b.length) {
					w = Math.max(w, dv.getUint16(i + 5, true));
					h = Math.max(h, dv.getUint16(i + 7, true));
				}
				break;
			}
		}
	}
	return dims("gif", w, h);
}
/** WebP：RIFF の最初のチャンク（VP8・VP8L・VP8X）の見出しから読む */
function sniffWebp(b) {
	if (!startsWith(b, 0, "RIFF") || !startsWith(b, 8, "WEBP")) return UNKNOWN;
	const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
	if (startsWith(b, 12, "VP8 ")) {
		if (b.length < 30 || b[23] !== 157 || b[24] !== 1 || b[25] !== 42) return UNKNOWN;
		return dims("webp", dv.getUint16(26, true) & 16383, dv.getUint16(28, true) & 16383);
	}
	if (startsWith(b, 12, "VP8L")) {
		if (b.length < 25 || b[20] !== 47) return UNKNOWN;
		const bits = dv.getUint32(21, true);
		return dims("webp", (bits & 16383) + 1, (bits >>> 14 & 16383) + 1);
	}
	if (startsWith(b, 12, "VP8X")) {
		if (b.length < 30) return UNKNOWN;
		const u24 = (at) => b[at] | b[at + 1] << 8 | b[at + 2] << 16;
		return dims("webp", u24(24) + 1, u24(27) + 1);
	}
	return UNKNOWN;
}
/**
* 先頭バイトで画像の形式を判定し、ヘッダから縦横（px）を読む。
* 描けない形式（EMF・WMF・TIFF）は形式名つきで、判定できないもの（SVG・不正・切れた・空）は名前なしで `ok: false` を返す。
* どんなバイト列でも例外を出さない。
*/
function sniffImage(b) {
	if (b.length >= 8 && b[0] === 137 && startsWith(b, 1, "PNG") && b[4] === 13) {
		if (b.length < 24 || !startsWith(b, 12, "IHDR")) return UNKNOWN;
		const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
		const w = dv.getUint32(16);
		const h = dv.getUint32(20);
		return w > 2147483647 || h > 2147483647 ? UNKNOWN : dims("png", w, h);
	}
	if (b.length >= 3 && b[0] === 255 && b[1] === 216 && b[2] === 255) return sniffJpeg(b);
	if (startsWith(b, 0, "GIF87a") || startsWith(b, 0, "GIF89a")) return sniffGif(b);
	if (startsWith(b, 0, "BM") && b.length >= 22) {
		const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
		const dibSize = dv.getUint32(14, true);
		if (!BMP_DIB_SIZES.has(dibSize)) return UNKNOWN;
		if (dibSize === 12) return dims("bmp", dv.getUint16(18, true), dv.getUint16(20, true));
		if (b.length < 26) return UNKNOWN;
		return dims("bmp", dv.getInt32(18, true), Math.abs(dv.getInt32(22, true)));
	}
	if (startsWith(b, 0, "RIFF")) return sniffWebp(b);
	if (b.length >= 44 && b[0] === 1 && b[1] === 0 && b[2] === 0 && b[3] === 0 && startsWith(b, 40, " EMF")) return {
		ok: false,
		label: "EMF"
	};
	if (b.length >= 4 && b[0] === 215 && b[1] === 205 && b[2] === 198 && b[3] === 154 || b.length >= 4 && (b[0] === 1 || b[0] === 2) && b[1] === 0 && b[2] === 9 && b[3] === 0) return {
		ok: false,
		label: "WMF"
	};
	if (b.length >= 4 && (b[0] === 73 && b[1] === 73 && b[2] === 42 && b[3] === 0 || b[0] === 77 && b[1] === 77 && b[2] === 0 && b[3] === 42)) return {
		ok: false,
		label: "TIFF"
	};
	return UNKNOWN;
}
/** 空の ImageBudget */
function createImageBudget() {
	return {
		used: 0,
		exhausted: false,
		counted: /* @__PURE__ */ new Set()
	};
}
/**
* zip の中の画像を読み、形式と縦横を返す。出せないときも例外にせず、理由つきの `ok: false` を返す。
* - 1 枚の上限（maxImageBytes）で展開しながら止める（超えたら too-large。これを先に判定するので、シートの残りが少なくても理由は取り違えない）。
* - 1 枚に収まる画像でも、シートの合計の残りを超えるなら over-budget。一度 over-budget になったら、以後そのシートの新しい画像はすべて over-budget（読まずに断る）。
*   すでに数えた画像（同じ画像の再読み込み）は、合計を使い切った後も読める。
* - 形式を判定したあと、1 辺と画素数の上限を確かめる。
*/
async function readImageFrom(pkg, path, limits, budget) {
	if (!pkg.has(path)) return {
		ok: false,
		reason: "missing",
		label: null
	};
	const charged = !budget.counted.has(path);
	if (charged && budget.exhausted) return {
		ok: false,
		reason: "over-budget",
		label: null
	};
	let bytes;
	try {
		bytes = await pkg.readBytes(path, limits.maxImageBytes);
	} catch (e) {
		if (e instanceof XlsxError && e.kind === "too-large") return {
			ok: false,
			reason: "too-large",
			label: null
		};
		return {
			ok: false,
			reason: "broken",
			label: null
		};
	}
	if (bytes === null) return {
		ok: false,
		reason: "missing",
		label: null
	};
	if (charged) {
		if (budget.used + bytes.length > limits.maxSheetImageBytes) {
			budget.exhausted = true;
			return {
				ok: false,
				reason: "over-budget",
				label: null
			};
		}
		budget.used += bytes.length;
		budget.counted.add(path);
	}
	const sniff = sniffImage(bytes);
	if (!sniff.ok) return {
		ok: false,
		reason: "unsupported",
		label: sniff.label
	};
	const { format, width, height } = sniff;
	if (width > limits.maxImageSidePx || height > limits.maxImageSidePx || width * height > limits.maxImagePixels) return {
		ok: false,
		reason: "oversize-px",
		label: null
	};
	return {
		ok: true,
		format,
		mime: MIME[format],
		width,
		height,
		bytes
	};
}
// src/core/open.ts
/**
* xlsx（zip）を開く。先頭の印の確認・ファイル数と展開後の大きさの上限・部品の文字での読み出し。
* 暗号化されたブック（OLE 形式。先頭 D0 CF 11 E0）と、zip でないものは、zip として読む前に弾く。
*/
/** Uint8Array へそろえる */
function toBytes(data) {
	return data instanceof Uint8Array ? data : new Uint8Array(data);
}
/**
* 先頭の印を確かめる。
* - `PK\x03\x04` 以外は zip でない。
* - OLE（`D0 CF 11 E0`）は、パスワード付きの xlsx（暗号化）か古い形式。どちらにしても読めないので encrypted。
*/
function checkSignature(bytes) {
	if (bytes.length >= 4 && bytes[0] === 208 && bytes[1] === 207 && bytes[2] === 17 && bytes[3] === 224) throw new XlsxError("encrypted");
	if (bytes.length < 4 || bytes[0] !== 80 || bytes[1] !== 75 || bytes[2] !== 3 || bytes[3] !== 4) throw new XlsxError("not-zip");
}
/**
* zip の 1 ファイルを展開しながら、実際に出てきた大きさを数える。maxBytes を超えたらそこで止めて too-large。
* 見出しの大きさ（偽りうる）が maxBytes を超えるものは、展開前に断る。流れてきた塊は onChunk へ渡し、全体の大きさを返す。
*/
function streamEntry(file, maxBytes, onChunk) {
	const declared = file._data?.uncompressedSize;
	if (typeof declared === "number" && declared > maxBytes) return Promise.reject(new XlsxError("too-large", "part"));
	return new Promise((resolve, reject) => {
		let total = 0;
		let done = false;
		const stream = file.internalStream("uint8array");
		stream.on("data", (chunk) => {
			if (done) return;
			total += chunk.length;
			if (total > maxBytes) {
				done = true;
				stream.pause();
				reject(new XlsxError("too-large", "part"));
				return;
			}
			onChunk(chunk);
		});
		stream.on("error", () => {
			if (done) return;
			done = true;
			reject(new XlsxError("broken", "entry"));
		});
		stream.on("end", () => {
			if (done) return;
			done = true;
			resolve(total);
		});
		stream.resume();
	});
}
/** zip を開く。壊れていれば broken、ファイル数が多すぎれば too-large */
async function openPackage(bytes, limits) {
	let zip;
	try {
		zip = await JSZip.loadAsync(bytes);
	} catch {
		throw new XlsxError("broken", "zip");
	}
	if (Object.keys(zip.files).length > limits.maxEntries) throw new XlsxError("too-large", "entries");
	return {
		has: (path) => zip !== null && zip.file(path) !== null,
		async readText(path, maxBytes) {
			if (zip === null) throw new XlsxError("broken", "closed");
			const file = zip.file(path);
			if (file === null) return null;
			const decoder = new TextDecoder("utf-8");
			const parts = [];
			await streamEntry(file, maxBytes, (chunk) => parts.push(decoder.decode(chunk, { stream: true })));
			parts.push(decoder.decode());
			return parts.join("");
		},
		async readBytes(path, maxBytes) {
			if (zip === null) throw new XlsxError("broken", "closed");
			const file = zip.file(path);
			if (file === null) return null;
			const chunks = [];
			const total = await streamEntry(file, maxBytes, (chunk) => chunks.push(chunk));
			const out = new Uint8Array(total);
			let at = 0;
			for (const c of chunks) {
				out.set(c, at);
				at += c.length;
			}
			return out;
		},
		dispose() {
			zip = null;
		}
	};
}
// src/core/shared-strings.ts
/**
* sharedStrings.xml の読み取り。
* - ふりがな（`rPh`）は除く。リッチテキスト（`r` の連なり）は連結して平文にする。
* - 文字の `&amp;` などの実体参照と、Excel の `_xHHHH_` を戻す。
* - 1 件ずつの展開は引かれたときに行う（索引だけ先に作る）。
*/
var RPH_OPEN = new RegExp(String.raw`<(${NS})rPh(?=[\s>/])[^<>]*>`, "g");
/** ふりがな（`<rPh>…</rPh>`）を取り除く。閉じの無い `<rPh>` に当たったら、以後に閉じは無いので残りはそのまま残す */
function stripPhonetic(s) {
	RPH_OPEN.lastIndex = 0;
	let out = "";
	let pos = 0;
	for (;;) {
		const m = RPH_OPEN.exec(s);
		if (m === null) break;
		let end = RPH_OPEN.lastIndex;
		if (!m[0].endsWith("/>")) {
			const closeTag = `</${m[1] ?? ""}rPh>`;
			const close = s.indexOf(closeTag, end);
			if (close < 0) break;
			end = close + closeTag.length;
		}
		out += s.slice(pos, m.index);
		pos = end;
		RPH_OPEN.lastIndex = end;
	}
	return pos === 0 ? s : out + s.slice(pos);
}
/**
* `<si>`・`<is>` の中身（XML の断片）を平文にする。ふりがなを除き、`<t>` を順に連結する。
* @param inner 開きタグと閉じタグの間の文字列
*/
function richTextToPlain(inner) {
	const body = inner.indexOf("rPh") >= 0 ? stripPhonetic(inner) : inner;
	let out = "";
	forEachElement(body, "t", (_open, text) => {
		if (text) out += decodeText(text);
	});
	return out;
}
/** sharedStrings.xml を読む。`<si>` の位置の索引だけを作り、文字への展開は get で行う */
function parseSharedStrings(xml) {
	const inners = [];
	if (xml !== null) forEachElement(xml, "si", (_open, body) => inners.push(body ?? ""));
	return {
		count: inners.length,
		get(i) {
			const inner = inners[i];
			return inner === void 0 ? "" : richTextToPlain(inner);
		}
	};
}
// src/core/sheet.ts
/** 空のシート（グラフシートなど、セルを持たないもの） */
function createEmptySheet() {
	return {
		rowCount: 0,
		colCount: 0,
		layout: emptyLayout(),
		borders: [],
		formats: [],
		rowFormat: () => void 0,
		colFormat: () => -1,
		getRows: () => [],
		release() {}
	};
}
var V_OPEN = new RegExp(String.raw`<(${NS})v(?=[\s>/])[^<>]*>`);
var IS_OPEN = new RegExp(String.raw`<(${NS})is(?=[\s>/])[^<>]*>`);
/**
* 断片の中で最初の要素（開きタグの正規表現 open で探す）の中身を返す。自己閉じ・閉じタグの無い要素・要素が無いときは undefined。
* 閉じは `[\s\S]*?` ではなく `indexOf` で探すので、閉じの無い要素を並べた細工でも線形時間。
*/
function bodyOf(xml, open, name) {
	const m = open.exec(xml);
	if (m === null || m[0].endsWith("/>")) return void 0;
	const start = m.index + m[0].length;
	const close = xml.indexOf(`</${m[1] ?? ""}${name}>`, start);
	return close < 0 ? void 0 : xml.slice(start, close);
}
var MERGE_RE = new RegExp(String.raw`<${NS}mergeCell\b[^<>]*?\sref\s*=\s*["']([^"']+)["']`, "g");
var ROW_NUM_RE = /\sr\s*=\s*["'](\d+)["']/;
var CELL_COL_RE = /\sr\s*=\s*["']([A-Z]{1,3})\d*["']/;
var HT_RE = /\sht\s*=\s*["']([^"']*)["']/;
var HIDDEN_RE = /\shidden\s*=\s*["']([^"']*)["']/;
var FORMAT_PR_RE = new RegExp(String.raw`<${NS}sheetFormatPr\b[^<>]*>`);
var CUSTOM_RE = /\scustomHeight\s*=\s*["']([^"']*)["']/;
var COL_RE = new RegExp(String.raw`<${NS}col\b[^<>]*>`, "g");
var CELL_S_RE = /\ss\s*=\s*["'](\d+)["']/;
var CUSTOM_FORMAT_RE = /\scustomFormat\s*=\s*["']([^"']*)["']/;
var SHEET_DATA_RE = new RegExp(String.raw`<(${NS})sheetData\b([^<>]*?)(/?)>`);
/**
* シートの XML から、シートを開く。
* @throws XlsxError Strict 形式は unsupported、`<sheetData>` の閉じが無ければ broken、行・結合が多すぎれば too-large
*/
function createSheet(xml, ctx) {
	if (looksStrict(xml)) throw new XlsxError("unsupported", "strict");
	const sd = SHEET_DATA_RE.exec(xml);
	let rowNums = /* @__PURE__ */ new Int32Array(0);
	let starts = /* @__PURE__ */ new Int32Array(0);
	let ends = /* @__PURE__ */ new Int32Array(0);
	let rowLayouts = [];
	const layoutBase = readLayoutHead(sd ? xml.slice(0, sd.index) : xml, ctx.mdw, ctx.fontRowPx);
	let dataEnd = 0;
	let prefix = "";
	if (sd) {
		prefix = sd[1] ?? "";
		const selfClosing = sd[3] === "/";
		const dataStart = sd.index + sd[0].length;
		if (selfClosing) dataEnd = dataStart;
		else {
			const close = xml.indexOf(`</${prefix}sheetData>`, dataStart);
			if (close < 0) throw new XlsxError("broken", "sheetData");
			dataEnd = close;
			({rowNums, starts, ends, rowLayouts} = indexRows(xml, prefix, dataStart, dataEnd, layoutBase));
		}
	}
	let source = xml;
	const closeCell = `</${prefix}c>`;
	const openCell = `<${prefix}c`;
	const n = rowNums.length;
	const cellText = (attrs, body) => {
		const t = attrs.t;
		let text;
		if (t === "inlineStr") {
			const is = bodyOf(body, IS_OPEN, "is");
			text = is === void 0 ? "" : richTextToPlain(is);
		} else {
			const raw = bodyOf(body, V_OPEN, "v");
			if (raw === void 0) return "";
			if (t === "s") {
				const idx = Number(raw);
				text = Number.isInteger(idx) ? ctx.strings.get(idx) : "";
			} else if (t === "str") text = decodeText(raw);
			else if (t === "b") {
				const v = raw.trim();
				text = v === "1" || v === "true" ? "TRUE" : "FALSE";
			} else if (t === "e") text = decodeEntities(raw);
			else if (t === "d") {
				const iso = decodeEntities(raw);
				const parsed = isoToSerial(iso, ctx.date1904);
				text = parsed ? ctx.formatter.date(parsed.serial, xfOf(attrs), parsed.hasTime) : iso;
			} else {
				const v = raw.trim();
				if (v === "") return "";
				const num = Number(v);
				text = Number.isFinite(num) ? ctx.formatter.number(num, xfOf(attrs)) : decodeEntities(raw);
			}
		}
		return text.length > 32767 ? text.slice(0, MAX_CELL_TEXT) : text;
	};
	/** `<c t>` からセルの型を決める（t 無し・`n`・未知の値は数値） */
	const kindOf = (t) => {
		switch (t) {
			case "s":
			case "str":
			case "inlineStr": return "s";
			case "b": return "b";
			case "e": return "e";
			case "d": return "d";
			default: return "n";
		}
	};
	/** セルの中身に数式（`<f>`）があるか */
	const hasFormula = (body) => {
		const tag = `<${prefix}f`;
		for (let i = body.indexOf(tag); i >= 0; i = body.indexOf(tag, i + 1)) if (isTagDelimiter(body[i + tag.length])) return true;
		return false;
	};
	/**
	* 文字が空のセルが「値のあるセル」か（Excel は右隣への文字のはみ出しを止める）。次の 2 つだけ：
	* - 数式の結果が空文字（`t="str"` で `<f>` と `<v>` がある）
	* - 空の文字列を指す共有文字列（`t="s"` で番号が範囲内）
	* それ以外（値の要素が無い・数式の無い `t="str"` の空の `<v>`・空の inlineStr・文字列以外）は空のセル。
	*/
	const isEmptyString = (attrs, body) => {
		const t = attrs.t;
		if (t === "s") {
			const raw = bodyOf(body, V_OPEN, "v");
			if (raw === void 0) return false;
			const idx = Number(raw);
			return Number.isInteger(idx) && idx >= 0 && idx < ctx.strings.count;
		}
		if (t === "str") return bodyOf(body, V_OPEN, "v") !== void 0 && hasFormula(body);
		return false;
	};
	/** 行の断片の中のセルを順に渡す。`r` が無い・不正なら、直前のセルの次の列とみなす（列を数えて補う） */
	const eachCell = (seg, cb) => {
		let prevCol = 0;
		forEachElement(seg, "c", (open, body) => {
			const attrs = parseAttrs(open);
			const ref = attrs.r;
			const col = ref !== void 0 ? colOfRef(ref) : prevCol + 1;
			if (col === 0 || col > 16384) return;
			prevCol = col;
			if (body === null) return;
			cb(col, attrs, body);
		});
	};
	/** i 番目の行（索引の位置）を cells へ展開する。cells の長さより右の列は捨てる */
	const expandRow = (i, cells) => {
		if (source === null) return;
		const seg = source.slice(starts[i], ends[i]);
		eachCell(seg, (col, attrs, body) => {
			if (col > cells.length) return;
			const text = cellText(attrs, body);
			if (text !== "") cells[col - 1] = {
				text,
				kind: kindOf(attrs.t),
				colSpan: 1,
				rowSpan: 1
			};
			else if (isEmptyString(attrs, body)) cells[col - 1] = {
				text,
				kind: "s",
				colSpan: 1,
				rowSpan: 1
			};
		});
	};
	const xfBorder = ctx.borderTable.xfBorder;
	const xfFormat = ctx.formatTable.xfFormat;
	const formatList = ctx.formatTable.formats;
	/** 罫線を持つ書式が 1 つでもあるか（無ければ、罫線のための走査をしない） */
	const hasBorders = ctx.borderTable.borders.length > 0;
	/** 見える書式が 1 つでもあるか（無ければ、書式のための走査も行・列の書式の解決もしない） */
	const hasFormats = formatList.length > 0;
	/** 書式の番号 → 紙でない塗りがあるか（表の範囲を広げる対象。装飾だけ・不透明な白の塗り＝紙では広げない） */
	const formatFilled = Uint8Array.from(formatList, (f) => f.fill !== void 0 && !isPaperFill(f.fill) ? 1 : 0);
	const hasFilled = formatFilled.includes(1);
	/**
	* xf の番号 → 表の範囲を広げるか（罫線があるか、紙でない塗りの書式か）。
	* 行の断片にこれらの xf を参照する `s` が 1 つも無ければ、その行はセルの要素単位の走査をしない（lastStyledCol）。
	*/
	const xfExtends = new Uint8Array(Math.max(xfBorder.length, xfFormat.length));
	for (let x = 0; x < xfExtends.length; x++) {
		const f = xfFormat[x] ?? -1;
		if ((xfBorder[x] ?? -1) >= 0 || f >= 0 && formatFilled[f] === 1) xfExtends[x] = 1;
	}
	/** `s` が無いセルは xf 0 の書式を引く。xf 0 が紙でない塗りのときだけ、`s` の有無で行を絞れない */
	const xf0Filled = (xfFormat[0] ?? -1) >= 0 && formatFilled[xfFormat[0]] === 1;
	const sAttrRe = new RegExp(CELL_S_RE.source, "g");
	/** 断片に、表の範囲を広げる xf を参照する `s` があるか（無ければ lastStyledCol の走査を省ける） */
	const mayExtend = (seg) => {
		if (xf0Filled) return true;
		sAttrRe.lastIndex = 0;
		for (let m = sAttrRe.exec(seg); m !== null; m = sAttrRe.exec(seg)) if (xfExtends[Number(m[1])] === 1) return true;
		return false;
	};
	const cellTagRe = new RegExp(String.raw`<${prefix}c(?=[\s>/])[^<>]*>`, "g");
	/**
	* 行の断片の中のセルの開きタグを順に渡す（値のあるセルも自己閉じの書式だけのセルも。閉じは探さない）。
	* 列は `r` 属性、無い・不正なら直前の次（eachCell と同じ。範囲外の列は渡さない）。
	* 渡すのは、セルの要素 1 つにつき 1 回。border は罫線の番号（`s` が無い・cellXfs の範囲外・罫線なしは -1）、
	* format は書式の番号（`s` が無ければ xf 0。cellXfs の範囲外・書式なしは -1）。1 回の走査で両方を引く。
	*/
	const eachStyledCell = (seg, cb) => {
		let prevCol = 0;
		cellTagRe.lastIndex = 0;
		for (let m = cellTagRe.exec(seg); m !== null; m = cellTagRe.exec(seg)) {
			const tag = m[0];
			const ref = CELL_COL_RE.exec(tag)?.[1];
			const col = ref !== void 0 ? colOfRef(ref) : prevCol + 1;
			if (col === 0 || col > 16384) continue;
			prevCol = col;
			const sm = CELL_S_RE.exec(tag);
			const xf = sm === null ? 0 : Number(sm[1]);
			cb(col, sm === null ? -1 : xfBorder[xf] ?? -1, xfFormat[xf] ?? -1);
		}
	};
	/** i 番目の行の、罫線か塗りを持つ最後のセルの列（無ければ 0） */
	const lastStyledCol = (i) => {
		let max = 0;
		const seg = source.slice(starts[i], ends[i]);
		if (!mayExtend(seg)) return 0;
		eachStyledCell(seg, (col, b, f) => {
			if (col > max && (b >= 0 || f >= 0 && formatFilled[f] === 1)) max = col;
		});
		return max;
	};
	/**
	* 行の断片の開きタグから、行の書式（`customFormat="1"` の行の `s`）の書式の番号を返す。
	* `customFormat` が無い・偽なら undefined（行の書式なし）。`s` が無い・範囲外・書式なしの xf は -1（セルの要素の無い位置を書式なしにする）。
	*/
	const rowFormatOf = (seg) => {
		const gt = seg.indexOf(">");
		const tag = gt < 0 ? seg : seg.slice(0, gt + 1);
		if (!boolAttr(CUSTOM_FORMAT_RE.exec(tag)?.[1])) return void 0;
		const sm = CELL_S_RE.exec(tag);
		return xfFormat[sm === null ? 0 : Number(sm[1])] ?? -1;
	};
	/** i 番目の行に値のセルがあるか（まず `</c>` の有無で安く確かめ、あるときだけ実際に読む） */
	const rowHasText = (i) => {
		const seg = source.slice(starts[i], ends[i]);
		if (seg.indexOf(closeCell) < 0) return false;
		let found = false;
		eachCell(seg, (_c, attrs, body) => {
			if (!found && cellText(attrs, body) !== "") found = true;
		});
		return found;
	};
	/** i 番目の行の、値のある最後のセルの列（無ければ 0）。多くの行は最後のセルの r 属性だけで決まる */
	const lastValueCol = (i) => {
		const seg = source.slice(starts[i], ends[i]);
		const closeAt = seg.lastIndexOf(closeCell);
		if (closeAt < 0) return 0;
		let openAt = seg.lastIndexOf(openCell, closeAt);
		while (openAt >= 0 && !isTagDelimiter(seg[openAt + openCell.length])) openAt = seg.lastIndexOf(openCell, openAt - 1);
		if (openAt >= 0) {
			const gt = seg.indexOf(">", openAt);
			const m = CELL_COL_RE.exec(seg.slice(openAt, gt < 0 ? void 0 : gt));
			if (m) {
				const col = colOfRef(m[1]);
				if (col > 0) return col;
			}
		}
		let max = 0;
		eachCell(seg, (col, attrs, body) => {
			if (col > max && cellText(attrs, body) !== "") max = col;
		});
		return max;
	};
	let lastRowIdx = -1;
	for (let i = n - 1; i >= 0; i--) if (rowHasText(i)) {
		lastRowIdx = i;
		break;
	}
	let rows = lastRowIdx >= 0 ? rowNums[lastRowIdx] : 0;
	let cols = 0;
	for (let i = 0; i <= lastRowIdx; i++) {
		const c = lastValueCol(i);
		if (c > cols) cols = c;
	}
	if (hasBorders || hasFilled) for (let i = n - 1; i >= 0; i--) {
		const c = lastStyledCol(i);
		if (c === 0) continue;
		if (rowNums[i] > rows) rows = rowNums[i];
		if (c > cols) cols = c;
	}
	let merges = [];
	if (sd && source.length > dataEnd) {
		const tail = source.slice(dataEnd);
		for (const m of tail.matchAll(MERGE_RE)) {
			const [a, b] = (m[1] ?? "").split(":");
			const p = parseCellRef(a ?? "");
			const q = b === void 0 ? p : parseCellRef(b);
			if (!p || !q) continue;
			const r1 = Math.min(p.r, q.r);
			const r2 = Math.max(p.r, q.r);
			const c1 = Math.min(p.c, q.c);
			const c2 = Math.max(p.c, q.c);
			if (r1 === r2 && c1 === c2) continue;
			if (merges.length >= ctx.limits.maxMerges) throw new XlsxError("too-large", "merges");
			merges.push({
				r1,
				c1,
				r2,
				c2
			});
			if (r2 > rows) rows = r2;
			if (c2 > cols) cols = c2;
		}
	}
	const rowCount = rows;
	const colCount = Math.min(cols, MAX_COLS);
	/**
	* 列の書式（`<col style>`）→ 表の列ごとの書式の番号（長さ colCount。指定の無い列・書式なしは -1）。
	* 見える書式の列が 1 つも無ければ null。重なった指定は文書で先のものを採る。
	* 範囲は、まだ決まっていない列だけを数える（重なりが多くても列数にほぼ比例）。
	*/
	const colFormats = (() => {
		if (!hasFormats || colCount === 0) return null;
		const out = new Int32Array(colCount).fill(-1);
		const next = Int32Array.from({ length: colCount + 2 }, (_v, i) => i);
		const find = (c0) => {
			let r = c0;
			while (next[r] !== r) r = next[r];
			for (let c = c0; next[c] !== r;) {
				const t = next[c];
				next[c] = r;
				c = t;
			}
			return r;
		};
		let any = false;
		for (const cl of layoutBase.cols) {
			if (cl.xf === void 0) continue;
			const f = xfFormat[cl.xf] ?? -1;
			for (let c = find(cl.min); c <= Math.min(cl.max, colCount); c = find(c)) {
				out[c - 1] = f;
				next[c] = c + 1;
				if (f >= 0) any = true;
			}
		}
		return any ? out : null;
	})();
	/** rowNums の中で num 以上の最初の位置（無ければ n） */
	const lowerBound = (num) => {
		let lo = 0;
		let hi = n;
		while (lo < hi) {
			const mid = lo + hi >> 1;
			if (rowNums[mid] < num) lo = mid + 1;
			else hi = mid;
		}
		return lo;
	};
	/** 範囲の外にある結合の左上の文字を取る（その行 1 本だけを読む） */
	const textAt = (r, c) => {
		const tmp = new Array(colCount).fill(void 0);
		for (let i = lowerBound(r); i < n && rowNums[i] === r; i++) expandRow(i, tmp);
		return tmp[c - 1]?.text ?? "";
	};
	return {
		rowCount,
		colCount,
		layout: {
			...layoutBase,
			rows: rowLayouts
		},
		borders: ctx.borderTable.borders,
		formats: formatList,
		rowFormat(r) {
			if (!hasFormats || source === null || !Number.isInteger(r) || r < 1) return void 0;
			const i = lowerBound(r);
			if (i >= n || rowNums[i] !== r) return void 0;
			return rowFormatOf(source.slice(starts[i], ends[i]));
		},
		colFormat(c) {
			if (!hasFormats || !(c >= 1) || c > 16384) return -1;
			for (const cl of layoutBase.cols) if (cl.xf !== void 0 && cl.min <= c && c <= cl.max) return xfFormat[cl.xf] ?? -1;
			return -1;
		},
		getRows(first, last) {
			const from = Math.max(1, Math.trunc(first));
			const to = Math.min(rowCount, Math.trunc(last));
			if (source === null || !(from <= to)) return [];
			const out = [];
			for (let r = from; r <= to; r++) out.push({
				r,
				cells: new Array(colCount).fill(void 0)
			});
			for (let i = lowerBound(from); i < n && rowNums[i] <= to; i++) expandRow(i, out[rowNums[i] - from].cells);
			if (hasBorders || hasFormats) {
				const rowFormat = [];
				const present = [];
				for (let i = lowerBound(from); i < n && rowNums[i] <= to; i++) {
					const oi = rowNums[i] - from;
					const row = out[oi];
					const seg = source.slice(starts[i], ends[i]);
					let pres = present[oi];
					if (hasFormats) {
						const rf = rowFormat[oi] ?? rowFormatOf(seg);
						if (rf !== void 0) rowFormat[oi] = rf;
						if (pres === void 0 && (rf !== void 0 ? rf >= 0 : colFormats !== null)) pres = present[oi] = new Uint8Array(colCount);
					}
					eachStyledCell(seg, (col, b, f) => {
						if (col > colCount) return;
						if (pres !== void 0) pres[col - 1] = 1;
						if (b >= 0) (row.borders ??= new Array(colCount).fill(void 0))[col - 1] = b;
						if (f >= 0) (row.formats ??= new Array(colCount).fill(void 0))[col - 1] = f;
					});
				}
				if (hasFormats) for (let oi = 0; oi < out.length; oi++) {
					const rf = rowFormat[oi];
					if (rf !== void 0 && rf < 0) continue;
					if (rf === void 0 && colFormats === null) continue;
					const row = out[oi];
					const pres = present[oi];
					for (let c = 0; c < colCount; c++) {
						if (pres !== void 0 && pres[c] === 1) continue;
						const f = rf ?? colFormats[c];
						if (f >= 0) (row.formats ??= new Array(colCount).fill(void 0))[c] = f;
					}
				}
			}
			for (const mg of merges) {
				if (mg.r1 > to || mg.r2 < from) continue;
				const top = Math.max(mg.r1, from);
				const bottom = Math.min(mg.r2, to);
				const right = Math.min(mg.c2, colCount);
				const text = mg.r1 >= from ? out[mg.r1 - from].cells[mg.c1 - 1]?.text ?? "" : textAt(mg.r1, mg.c1);
				for (let r = top; r <= bottom; r++) {
					const cells = out[r - from].cells;
					for (let c = mg.c1; c <= right; c++) cells[c - 1] = null;
				}
				out[top - from].cells[mg.c1 - 1] = {
					text,
					colSpan: right - mg.c1 + 1,
					rowSpan: bottom - top + 1
				};
			}
			return out;
		},
		release() {
			source = null;
			rowNums = /* @__PURE__ */ new Int32Array(0);
			starts = /* @__PURE__ */ new Int32Array(0);
			ends = /* @__PURE__ */ new Int32Array(0);
			merges = [];
		}
	};
}
/** 数として読める属性値（空・数でない・無限は undefined） */
function numAttr(v) {
	if (v === void 0 || v.trim() === "") return void 0;
	const n = Number(v);
	return Number.isFinite(n) ? n : void 0;
}
/** 真偽の属性（`1`・`true` が真） */
var boolAttr = (v) => v === "1" || v === "true";
/**
* シート XML の `<sheetData>` より前の部分から、`<sheetFormatPr>` と `<cols>` を読む（行の寸法はここでは空）。
* 不正な値（負・数でない）は無いものとして既定にし、極端な値は Excel の上限（列幅 255 文字・行高 409.5pt）に丸める。
* `<col>` は `min`〜`max` を展開せず範囲のまま持つ。`<cols>` が複数あっても順にすべて読む。
*/
function readLayoutHead(head, mdw, fontRowPx) {
	const layout = { ...emptyLayout(mdw, fontRowPx) };
	const fp = FORMAT_PR_RE.exec(head);
	if (fp) {
		const a = parseAttrs(fp[0]);
		const dcw = numAttr(a.defaultColWidth);
		if (dcw !== void 0 && dcw >= 0) layout.defaultColWidth = Math.min(dcw, 255);
		const bcw = numAttr(a.baseColWidth);
		if (bcw !== void 0 && bcw >= 0) layout.baseColWidth = Math.min(bcw, 255);
		const drh = numAttr(a.defaultRowHeight);
		if (drh !== void 0 && drh >= 0) layout.defaultRowHeight = Math.min(drh, MAX_ROW_PT);
		layout.customHeight = boolAttr(a.customHeight);
		layout.zeroHeight = boolAttr(a.zeroHeight);
	}
	const cols = [];
	scan: for (let pos = 0; pos < head.length;) {
		const el = findElement(head, "cols", pos);
		if (el === null) break;
		pos = el.end;
		if (el.body === null) continue;
		for (const m of el.body.matchAll(COL_RE)) {
			if (cols.length >= 16384) break scan;
			const a = parseAttrs(m[0]);
			const min = numAttr(a.min);
			const max = numAttr(a.max);
			if (min === void 0 || max === void 0) continue;
			if (!Number.isInteger(min) || !Number.isInteger(max)) continue;
			if (min < 1 || min > 16384 || max < min) continue;
			const w = numAttr(a.width);
			cols.push({
				min,
				max: Math.min(max, MAX_COLS),
				width: w !== void 0 && w >= 0 ? Math.min(w, 255) : null,
				hidden: boolAttr(a.hidden),
				...styleOf(a.style)
			});
		}
	}
	layout.cols = cols;
	return layout;
}
/** `<col style>`（cellXfs の位置）を `ColLayout` に足す形にする。無い・負・整数でないなら何も足さない */
function styleOf(v) {
	const n = numAttr(v);
	return n !== void 0 && Number.isInteger(n) && n >= 0 ? { xf: n } : {};
}
/** セルの `s` 属性（cellXfs の位置）。無い・不正なら 0 */
function xfOf(attrs) {
	const s = Number(attrs.s);
	return Number.isInteger(s) && s >= 0 ? s : 0;
}
/**
* `<row>` の位置の索引を作る（セルは見ない）。
* 行番号は `r` 属性、無ければ直前の行の次。行番号が昇順でない文書は、番号で並べ替える。
* 範囲外の行番号（Excel の上限 1,048,576 より大きい）の行は索引に入れない（ただし前の行の終わりの目印には使う）。
* 同じ走査で、行の `ht`（pt）・`hidden`・`customHeight` も拾う（XML を 2 度読まない）。
* 寸法の指定のある行だけ `rowLayouts` に入れる（既定と同じ行は持たない。`zeroHeight` のときは、`<row>` のあるすべての行。
* 無い行を既定で非表示にするため）。
* `rowLayouts` は行番号の昇順・重複なしで、連続する同じ指定の行は 1 件にまとめる（全行に同じ `ht` があっても 1 件）。
* 同じ行番号が複数あるときは、文書で後の行を採る（後の行に指定が無ければ、前の行の指定も取り除く）。
*/
function indexRows(xml, prefix, dataStart, dataEnd, base) {
	const keepAll = base.zeroHeight;
	const rowOpen = `<${prefix}row`;
	const allNums = [];
	const allPos = [];
	const allFlag = [];
	const allHt = [];
	let prev = 0;
	for (let p = xml.indexOf(rowOpen, dataStart); p >= 0 && p < dataEnd; p = xml.indexOf(rowOpen, p + rowOpen.length)) {
		if (!isTagDelimiter(xml[p + rowOpen.length])) continue;
		const gt = xml.indexOf(">", p);
		const tag = xml.slice(p, gt < 0 || gt > dataEnd ? dataEnd : gt);
		const m = ROW_NUM_RE.exec(tag);
		const num = m ? parseInt(m[1], 10) : prev + 1;
		prev = num;
		if (allPos.length >= 2097152) throw new XlsxError("too-large", "rows");
		const inRange = num >= 1 && num <= 1048576;
		allNums.push(inRange ? num : 0);
		allPos.push(p);
		let flag = 0;
		let ht = NaN;
		if (inRange) {
			const raw = numAttr(HT_RE.exec(tag)?.[1]);
			const hasHt = raw !== void 0 && raw >= 0;
			if (hasHt) ht = Math.min(raw, MAX_ROW_PT);
			const hidden = boolAttr(HIDDEN_RE.exec(tag)?.[1]);
			const custom = hasHt && boolAttr(CUSTOM_RE.exec(tag)?.[1]);
			const plain = !hidden && (!hasHt || !custom && Math.abs(ht - base.defaultRowHeight) <= .01);
			if (keepAll || !plain) flag = 1 | (hidden ? 2 : 0) | (custom ? 4 : 0);
		}
		allFlag.push(flag);
		allHt.push(ht);
	}
	const order = [];
	let sorted = true;
	let last = 0;
	for (let i = 0; i < allNums.length; i++) {
		const num = allNums[i];
		if (num === 0) continue;
		if (num < last) sorted = false;
		last = num;
		order.push(i);
	}
	if (order.length > 1048576) throw new XlsxError("too-large", "rows");
	if (!sorted) order.sort((a, b) => allNums[a] - allNums[b] || a - b);
	const rowNums = new Int32Array(order.length);
	const starts = new Int32Array(order.length);
	const ends = new Int32Array(order.length);
	const rowLayouts = [];
	let lastR = 0;
	order.forEach((src, i) => {
		const r = allNums[src];
		rowNums[i] = r;
		starts[i] = allPos[src];
		ends[i] = src + 1 < allPos.length ? allPos[src + 1] : dataEnd;
		const tail = rowLayouts[rowLayouts.length - 1];
		if (r === lastR && tail !== void 0 && tail.max === r) {
			if (tail.min === r) rowLayouts.pop();
			else tail.max = r - 1;
		}
		lastR = r;
		const flag = allFlag[src];
		if (flag === 0) return;
		const ht = Number.isNaN(allHt[src]) ? null : allHt[src];
		const hidden = (flag & 2) !== 0;
		const custom = (flag & 4) !== 0;
		const t = rowLayouts[rowLayouts.length - 1];
		if (t !== void 0 && t.max === r - 1 && t.ht === ht && t.hidden === hidden && t.custom === custom) t.max = r;
		else rowLayouts.push({
			min: r,
			max: r,
			ht,
			hidden,
			custom
		});
	});
	return {
		rowNums,
		starts,
		ends,
		rowLayouts
	};
}
// src/core/font-metrics.ts
/** 表に無いフォントに使う行の名前 */
var FALLBACK = "calibri";
var CALIBRI = [
	[6, 4],
	[8, 6],
	[9, 6],
	[10, 7],
	[11, 7],
	[12, 8],
	[14, 10],
	[16, 11],
	[18, 12],
	[20, 14],
	[24, 16]
];
var ARIAL = [
	[6, 4],
	[8, 6],
	[9, 7],
	[10, 8],
	[11, 8],
	[12, 9],
	[14, 11],
	[16, 12],
	[18, 13],
	[20, 15],
	[24, 18]
];
var TIMES = [
	[6, 4],
	[8, 6],
	[9, 6],
	[10, 7],
	[11, 7],
	[12, 8],
	[14, 9],
	[16, 11],
	[18, 12],
	[20, 14],
	[24, 16]
];
var MS_GOTHIC = [
	[6, 4],
	[8, 6],
	[9, 6],
	[10, 7],
	[11, 8],
	[12, 8],
	[14, 10],
	[16, 11],
	[18, 12],
	[20, 14],
	[24, 16]
];
var MEIRYO = [
	[6, 5],
	[8, 7],
	[9, 7],
	[10, 9],
	[11, 9],
	[12, 10],
	[14, 12],
	[16, 14],
	[18, 15],
	[20, 17],
	[24, 20]
];
/** 正規化したフォント名 → 実測表（Yu Gothic は Arial と同じ値） */
var MDW_TABLE = {
	calibri: CALIBRI,
	arial: ARIAL,
	timesnewroman: TIMES,
	yugothic: ARIAL,
	mspgothic: MS_GOTHIC,
	msgothic: MS_GOTHIC,
	meiryo: MEIRYO,
	meiryoui: MEIRYO,
	malgungothic: [
		[6, 4],
		[8, 6],
		[9, 7],
		[10, 8],
		[11, 8],
		[12, 9],
		[14, 10],
		[16, 12],
		[18, 13],
		[20, 15],
		[24, 18]
	],
	gulim: [
		[6, 5],
		[8, 6],
		[9, 6],
		[10, 8],
		[11, 8],
		[12, 9],
		[14, 10],
		[16, 11],
		[18, 13],
		[20, 16],
		[24, 18]
	],
	dotum: [
		[6, 5],
		[8, 6],
		[9, 6],
		[10, 9],
		[11, 9],
		[12, 8],
		[14, 10],
		[16, 12],
		[18, 13],
		[20, 16],
		[24, 19]
	],
	segoeui: [
		[6, 4],
		[8, 6],
		[9, 6],
		[10, 8],
		[11, 8],
		[12, 9],
		[14, 10],
		[16, 12],
		[18, 13],
		[20, 15],
		[24, 17]
	]
};
/** 現地語名・全角表記 → 表の名前（正規化したあとの名前で引く） */
var ALIASES = {
	游ゴシック: "yugothic",
	mspゴシック: "mspgothic",
	msゴシック: "msgothic",
	メイリオ: "meiryo",
	メイリオui: "meiryoui",
	맑은고딕: "malgungothic",
	굴림: "gulim",
	돋움: "dotum"
};
/** 既定の行高（px・96dpi）。キーは「正規化した名前:サイズ」 */
var DEFAULT_ROW_PX = {
	"calibri:11": 20,
	"yugothic:11": 25,
	"malgungothic:11": 22,
	"mspgothic:11": 18,
	"arial:10": 17
};
/**
* フォント名を表の名前にそろえる。NFKC 正規化（全角英字を半角に）・小文字・空白除去のあと、現地語名の別名を引く。
* 表に無い名前は正規化しただけの名前を返す。
*/
function canonicalFontName(name) {
	const n = name.normalize("NFKC").toLowerCase().replace(/\s+/g, "");
	return ALIASES[n] ?? n;
}
/**
* フォントの名前・サイズ（pt）→ 最大桁幅 MDW（px・96dpi）。
* 実測表のサイズの間は、近い 2 点から線形補間して四捨五入する。表の範囲外は、端の点に比例させる。
* 表に無いフォントは Calibri の行を使う。サイズが不正（数でない・0 以下）なら null。
*/
function lookupMdw(name, size) {
	if (!Number.isFinite(size) || size <= 0) return null;
	const s = Math.min(size, 409);
	const row = MDW_TABLE[canonicalFontName(name)] ?? MDW_TABLE[FALLBACK];
	const first = row[0];
	const last = row[row.length - 1];
	if (s <= first[0]) return Math.max(1, Math.round(s * first[1] / first[0]));
	if (s >= last[0]) return Math.max(1, Math.round(s * last[1] / last[0]));
	for (let i = 1; i < row.length; i++) {
		const [s1, m1] = row[i];
		if (s > s1) continue;
		const [s0, m0] = row[i - 1];
		return Math.max(1, Math.round(m0 + (m1 - m0) * (s - s0) / (s1 - s0)));
	}
	return last[1];
}
/**
* フォントの既定の行高（px・96dpi）。実測した（名前・サイズ）の組だけを持ち、無ければ null。
* Calibri 11＝20、游ゴシック 11＝25、맑은 고딕 11＝22、MS Pゴシック 11＝18、Arial 10＝17。
*/
function fontDefaultRowPx(name, size) {
	return DEFAULT_ROW_PX[`${canonicalFontName(name)}:${size}`] ?? null;
}
// src/core/styles.ts
/**
* styles.xml の読み取り：セルの書式（cellXfs）の番号 → 数値の表示形式の番号、独自の表示形式（numFmts）。
* 書式のうち塗り・文字の装飾は読まない（塗り・太字・色などは cell-format.ts、罫線は borders.ts、色の解釈は cell-color.ts）。ただし寸法の換算に要る、先頭の font（名前・サイズ）から
* 最大桁幅 MDW と、フォントの既定の行高を出す（実測表は font-metrics.ts）。
*/
/**
* フォントの名前・サイズ（pt）→ 最大桁幅 MDW（px）。
* 実測表（font-metrics.ts）から、サイズの間は線形補間・範囲外は比例で求め、表に無いフォントは Calibri の行を使う。
* Calibri 11＝7px、游ゴシック 11＝8px、맑은 고딕 11＝8px。サイズが不正なら 7px。
*/
function maxDigitWidth(name, size) {
	return lookupMdw(name, size) ?? 7;
}
/** styles.xml を読む（無ければ空）。書式コードは実体参照を戻した文字列 */
function parseStyles(xml) {
	const styles = {
		xfNumFmtIds: [],
		customFormats: /* @__PURE__ */ new Map(),
		mdw: 7,
		fontRowPx: null
	};
	if (xml === null) return styles;
	for (const m of xml.matchAll(new RegExp(String.raw`<${NS}numFmt\b[^<>]*>`, "g"))) {
		const a = parseAttrs(m[0]);
		const id = Number(a.numFmtId);
		if (Number.isInteger(id) && a.formatCode !== void 0) styles.customFormats.set(id, a.formatCode);
	}
	const fonts = findElement(xml, "fonts");
	const font = fonts?.body ? findElement(fonts.body, "font") : null;
	if (font?.body) {
		const sz = new RegExp(String.raw`<${NS}sz\b[^<>]*>`).exec(font.body);
		const nm = new RegExp(String.raw`<${NS}name\b[^<>]*>`).exec(font.body);
		if (sz) {
			const raw = Number(parseAttrs(sz[0]).val);
			const size = Number.isFinite(raw) ? Math.min(raw, 409) : raw;
			const name = nm ? parseAttrs(nm[0]).val ?? "" : "";
			styles.mdw = maxDigitWidth(name, size);
			styles.fontRowPx = fontDefaultRowPx(name, size);
		}
	}
	const block = findElement(xml, "cellXfs");
	if (block?.body) for (const m of block.body.matchAll(new RegExp(String.raw`<${NS}xf\b[^<>]*>`, "g"))) {
		const id = Number(parseAttrs(m[0]).numFmtId);
		styles.xfNumFmtIds.push(Number.isInteger(id) ? id : 0);
	}
	return styles;
}
// src/core/geometry.ts
/**
* シートの寸法（layout）を px の座標へ換算する純関数（DOM 非依存）。
* - 表（colgroup・行高）・仮想化の上下余白・画像レイヤーが、同じ座標（Excel の寸法）を唯一の幾何データとして使う。
* - 累積は「既定値×番号＋疎な差分の累積和」で出す。前計算は指定のある数に比例し、
*   1,048,576 行・16,384 列でも全行・全列の配列は作らない。`rowAt`／`colAt` は二分探索。
* - 非表示の行・列は寸法 0。
*/
/** 行高 pt → px。`round(pt × 96 / 72)`（15pt＝20px、18.75pt＝25px） */
function ptToPx(pt) {
	return Math.round(pt * 96 / 72);
}
/**
* 列幅（`<col width>` の文字数・余白を含む値）→ px。
* ECMA-376 Part 1 の `col` 要素の換算式 `trunc(((256·w + trunc(128/MDW)) / 256) · MDW)`。
* @param w 列幅（文字数）
* @param mdw 最大桁幅 MDW（px）
*/
function colWidthToPx(w, mdw) {
	return Math.trunc((256 * w + Math.trunc(128 / mdw)) / 256 * mdw);
}
/**
* 既定の列幅（`defaultColWidth` が無いとき）→ px。`ceil((baseColWidth·MDW + 余白) / 8) · 8`
* （MDW 7 で 64px、MDW 8 で 72px）。余白は `2·ceil(MDW/4) + 1`（MDW 5〜8 では 5px）。
* 根拠: Excel 16 の実ファイルの `<col width>` の端数から余白を逆算（2026-10-27）。
*/
function defaultColWidthToPx(baseColWidth, mdw) {
	const padding = 2 * Math.ceil(mdw / 4) + 1;
	return Math.ceil((baseColWidth * mdw + padding) / 8) * 8;
}
/**
* 軸を作る。`spans` は並べ替え・重なりの除去（min の小さい範囲を優先し、後の範囲は重なりを削る）を行う。
* 前計算（累積和）は spans の数に比例し、番号の上限までの配列は作らない。
*/
function createAxis(def, input) {
	const sorted = input.filter((s) => s.min <= s.max).sort((a, b) => a.min - b.min || a.max - b.max);
	const spans = [];
	let lastMax = 0;
	for (const s of sorted) {
		const min = Math.max(s.min, lastMax + 1);
		if (min > s.max) continue;
		const prev = spans[spans.length - 1];
		if (prev !== void 0 && prev.max + 1 === min && prev.size === s.size) prev.max = s.max;
		else spans.push({
			min,
			max: s.max,
			size: s.size
		});
		lastMax = s.max;
	}
	const cum = new Array(spans.length + 1).fill(0);
	for (let k = 0; k < spans.length; k++) {
		const s = spans[k];
		cum[k + 1] = cum[k] + (s.max - s.min + 1) * (s.size - def);
	}
	/** min ≤ i を満たす範囲の数（二分探索） */
	const countStarting = (i) => {
		let lo = 0;
		let hi = spans.length;
		while (lo < hi) {
			const mid = lo + hi >> 1;
			if (spans[mid].min <= i) lo = mid + 1;
			else hi = mid;
		}
		return lo;
	};
	const offset = (i) => {
		if (!(i > 1)) return 0;
		const last = i - 1;
		const k = countStarting(last);
		let delta = 0;
		if (k > 0) {
			const s = spans[k - 1];
			delta = cum[k - 1] + (Math.min(s.max, last) - s.min + 1) * (s.size - def);
		}
		return def * last + delta;
	};
	const size = (i) => {
		if (!(i >= 1)) return 0;
		const k = countStarting(i);
		if (k > 0) {
			const s = spans[k - 1];
			if (i <= s.max) return s.size;
		}
		return def;
	};
	const at = (pos, n) => {
		if (!(n >= 1)) return 0;
		const p = pos > 0 ? pos : 0;
		let lo = 1;
		let hi = n;
		while (lo < hi) {
			const mid = Math.ceil((lo + hi) / 2);
			if (offset(mid) <= p) lo = mid;
			else hi = mid - 1;
		}
		return lo;
	};
	return {
		offset,
		size,
		at
	};
}
/** 数が有限で 0 以上のときだけ採り、上限 max に丸める（それ以外は既定へ。巨大な値で座標が壊れないように） */
var clampNum = (v, max, fallback) => Number.isFinite(v) && v >= 0 ? Math.min(v, max) : fallback;
/**
* layout から座標（px）の計算器を作る。
* - 列幅は `<col width>` を `colWidthToPx` で換算し、指定の無い列は既定の列幅（`defaultColWidth` があれば同じ式で換算、
*   無ければ `defaultColWidthToPx`）。非表示の列は 0。
* - 行高：シートの既定の行高は、`customHeight="1"` なら `defaultRowHeight`（pt）の換算値、
*   そうでなければフォントの既定の行高（`fontRowPx`。実測表）を優先し、無ければ `defaultRowHeight` の換算値。
*   （Excel は `customHeight` の無い `defaultRowHeight`・`<row ht>` を、その PC のフォントから計算した行高で置き換えるため）
*   行は、`customHeight="1"` の `ht` は `ptToPx(ht)`。`customHeight` の無い `ht` は、ファイルの
*   `defaultRowHeight` と ±0.01 で等しければシートの既定、違えば保存時の自動調整値として `ptToPx(ht)` とする。
*   `ht` が無ければシートの既定。非表示の行は 0。
*   `zeroHeight` のときは、`<row>` の無い行を既定で高さ 0 とする（`<row>` のある行は表示）。
* - 列幅・行高は Excel の上限（255 文字・409.5pt）に丸める。
* @param layout `XlsxSheet.layout`
* @param rows 表の行数（`XlsxSheet.rowCount`）。`totalHeight`・`rowAt` の範囲に使う
* @param cols 表の列数（`XlsxSheet.colCount`）。`totalWidth`・`colAt` の範囲に使う
*/
function createGeometry(layout, rows, cols) {
	const mdw = Number.isFinite(layout.mdw) && layout.mdw >= 1 ? Math.min(layout.mdw, 100) : 7;
	const defColPx = layout.defaultColWidth !== null ? colWidthToPx(clampNum(layout.defaultColWidth, 255, 0), mdw) : defaultColWidthToPx(clampNum(layout.baseColWidth, 255, 8), mdw);
	const colSpans = [];
	for (const c of layout.cols) {
		const min = Math.max(1, Math.trunc(c.min));
		const max = Math.min(MAX_COLS, Math.trunc(c.max));
		const size = c.hidden ? 0 : c.width === null ? defColPx : colWidthToPx(clampNum(c.width, 255, 0), mdw);
		colSpans.push({
			min,
			max,
			size
		});
	}
	const colAxis = createAxis(defColPx, colSpans);
	const defaultPt = clampNum(layout.defaultRowHeight, MAX_ROW_PT, 15);
	const fontRowPx = layout.fontRowPx !== null && Number.isFinite(layout.fontRowPx) && layout.fontRowPx >= 0 ? layout.fontRowPx : null;
	/** シートの既定の行高（px）。zeroHeight のとき `<row>` のある行（表示）の高さにも使う */
	const sheetRowPx = layout.customHeight ? ptToPx(defaultPt) : fontRowPx ?? ptToPx(defaultPt);
	const defRowPx = layout.zeroHeight ? 0 : sheetRowPx;
	const rowSpans = [];
	for (const r of layout.rows) {
		const min = Math.max(1, Math.trunc(r.min));
		const max = Math.min(MAX_ROWS, Math.trunc(r.max));
		let size;
		if (r.hidden) size = 0;
		else if (r.ht === null) size = sheetRowPx;
		else {
			const ht = clampNum(r.ht, MAX_ROW_PT, defaultPt);
			size = r.custom || Math.abs(ht - defaultPt) > .01 ? ptToPx(ht) : sheetRowPx;
		}
		if (size === defRowPx) continue;
		rowSpans.push({
			min,
			max,
			size
		});
	}
	const rowAxis = createAxis(defRowPx, rowSpans);
	const nRows = Math.max(0, Math.trunc(rows));
	const nCols = Math.max(0, Math.trunc(cols));
	return {
		colLeft: (c) => colAxis.offset(c),
		colWidth: (c) => colAxis.size(c),
		rowTop: (r) => rowAxis.offset(r),
		rowHeight: (r) => rowAxis.size(r),
		rowAt: (y) => rowAxis.at(y, nRows),
		colAt: (x) => colAxis.at(x, nCols),
		totalWidth: colAxis.offset(nCols + 1),
		totalHeight: rowAxis.offset(nRows + 1)
	};
}
// src/core/index.ts
/**
* xlsx の読み取り部品の公開の入口（画面に依存しない。Node の vitest でも動く）。
* jszip で zip を開き、workbook・共有文字列・styles・シートの XML を自前で読み、
* 数値の表示形式だけ `ssf` に任せる。シートは 1 枚ずつ開き、見えている行だけを展開する。
* セルの値・シート名はログに出さない（console を使わない）。
*/
/**
* xlsx を開く（シートの中身は openSheet で開くときに初めて展開する）。
* @throws XlsxError zip でない（not-zip）・暗号化（encrypted）・壊れている（broken）・大きすぎる（too-large）・読めない形式（unsupported）
*/
async function openXlsx(data, options = {}) {
	const limits = {
		...DEFAULT_LIMITS,
		...options.limits
	};
	const bytes = toBytes(data);
	checkSignature(bytes);
	const pkg = await openPackage(bytes, limits);
	let workbook;
	try {
		workbook = await loadWorkbook(pkg, limits);
	} catch (e) {
		pkg.dispose();
		throw e;
	}
	let closed = false;
	const opened = /* @__PURE__ */ new Set();
	let stringsPromise = null;
	let formatterPromise = null;
	let themePromise = null;
	const drawingRelIds = /* @__PURE__ */ new Map();
	const imageBudgets = /* @__PURE__ */ new Map();
	/** readImage を 1 つずつ順に走らせる（合計の上限を、呼んだ順に確実に数えるため） */
	let imageChain = Promise.resolve();
	const loadStrings = (p) => {
		stringsPromise ??= (async () => {
			const path = workbook.sharedStringsPath;
			return parseSharedStrings(path !== null && p.has(path) ? await p.readText(path, limits.maxSharedStringsBytes) : null);
		})();
		return stringsPromise;
	};
	/** テーマの 12 色（xl/theme/theme1.xml。無い・読めないときは既定の 12 色。ブックで 1 回だけ読む） */
	const loadThemeOnce = (p) => {
		themePromise ??= loadTheme(p, workbook.themePath, limits);
		return themePromise;
	};
	/** 表示形式の整形器・寸法の換算に使う最大桁幅 MDW・罫線の表・書式の表（いずれも styles.xml から。ブックで 1 回だけ読む） */
	const loadFormatter = (p) => {
		formatterPromise ??= (async () => {
			const path = workbook.stylesPath;
			const xml = path !== null && p.has(path) ? await p.readText(path, limits.maxSmallPartBytes) : null;
			const styles = parseStyles(xml);
			const theme = stylesUseTheme(xml) ? await loadThemeOnce(p) : DEFAULT_THEME;
			return {
				formatter: createFormatter(styles, workbook.date1904),
				mdw: styles.mdw,
				fontRowPx: styles.fontRowPx,
				borderTable: parseBorders(xml, theme, limits),
				formatTable: parseCellFormats(xml, theme, limits)
			};
		})();
		return formatterPromise;
	};
	return {
		sheets: workbook.sheets.map((s) => ({
			index: s.index,
			name: s.name,
			hidden: s.hidden
		})),
		async openSheet(index) {
			if (closed) throw new XlsxError("broken", "closed");
			const info = workbook.sheets[index - 1];
			if (!Number.isInteger(index) || info === void 0) throw new RangeError("シートの番号が範囲外");
			if (!info.isWorksheet) return createEmptySheet();
			if (info.path === null) throw new XlsxError("broken", "sheet path");
			const xml = await pkg.readText(info.path, limits.maxSheetXmlBytes);
			if (xml === null) throw new XlsxError("broken", "sheet part");
			drawingRelIds.set(index, findDrawingRelId(xml));
			const [strings, { formatter, mdw, fontRowPx, borderTable, formatTable }] = await Promise.all([loadStrings(pkg), loadFormatter(pkg)]);
			if (closed) throw new XlsxError("broken", "closed");
			const sheet = createSheet(xml, {
				strings,
				formatter,
				date1904: workbook.date1904,
				limits,
				mdw,
				fontRowPx,
				borderTable,
				formatTable
			});
			for (const ref of opened) if (ref.deref() === void 0) opened.delete(ref);
			opened.add(new WeakRef(sheet));
			return sheet;
		},
		async openDrawings(index) {
			if (closed) throw new XlsxError("broken", "closed");
			const info = workbook.sheets[index - 1];
			if (!Number.isInteger(index) || info === void 0) throw new RangeError("シートの番号が範囲外");
			if (!info.isWorksheet || info.path === null) return {
				drawings: [],
				notices: []
			};
			let relId = drawingRelIds.get(index);
			if (relId === void 0) {
				try {
					const xml = await pkg.readText(info.path, limits.maxSheetXmlBytes);
					relId = xml === null ? null : findDrawingRelId(xml);
				} catch (e) {
					if (closed) throw new XlsxError("broken", "closed");
					return {
						drawings: [],
						notices: [e instanceof XlsxError && e.kind === "too-large" ? "part-too-large" : "broken"]
					};
				}
				drawingRelIds.set(index, relId);
			}
			if (relId === null) return {
				drawings: [],
				notices: []
			};
			const theme = await loadThemeOnce(pkg);
			const result = await loadSheetDrawings(pkg, info.path, relId, limits, theme);
			if (closed) throw new XlsxError("broken", "closed");
			return result;
		},
		readImage(index, path) {
			if (closed) return Promise.reject(new XlsxError("broken", "closed"));
			if (!Number.isInteger(index) || workbook.sheets[index - 1] === void 0) return Promise.reject(/* @__PURE__ */ new RangeError("シートの番号が範囲外"));
			const budget = imageBudgets.get(index) ?? createImageBudget();
			imageBudgets.set(index, budget);
			const run = imageChain.then(async () => {
				const r = await readImageFrom(pkg, path, limits, budget);
				if (closed) throw new XlsxError("broken", "closed");
				return r;
			});
			imageChain = run.catch(() => void 0);
			return run;
		},
		close() {
			closed = true;
			for (const ref of opened) ref.deref()?.release();
			opened.clear();
			stringsPromise = null;
			formatterPromise = null;
			themePromise = null;
			drawingRelIds.clear();
			imageBudgets.clear();
			pkg.dispose();
		}
	};
}
export { ptToPx as a, isPaperFill as c, defaultColWidthToPx as i, colWidthToPx as n, fontDefaultRowPx as o, createGeometry as r, lookupMdw as s, openXlsx as t };
