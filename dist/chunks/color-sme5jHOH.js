import { n as MAX_GEOM_COMMANDS } from "./defs-BMhuSHU6.js";
// src/core/limits.ts
/**
* xlsx の読み取りの上限（定数の置き場は本ファイルの 1 か所だけ。図形の guide・path 命令の上限だけは geom の `defs.ts`）。
* 読み込むファイルは信頼できない入力として扱う（Excel の仕様上限と XML の大きさの見積もりに基づく値）。
* zip の展開（zip 爆弾）・巨大な添字・大量の結合セルで、画面が固まったりメモリが尽きたりしないようにする。
* 超えたら `XlsxError('too-large')`。
*/
/**
* 1 シートの XML の上限 = 160 MiB。
* セル 1 つの XML は `<c r="AB1234" s="12" t="s"><v>12345</v></c>` で 40〜80 バイト。
* 200 万セル級のシートでも 80 バイト × 200 万 = 160 MB に収まる見積もり（強制するセル数の上限ではない）。
* これより大きいと JS の文字列（2 バイト文字を含むと 2 倍）でブラウザのメモリを圧迫するので読まない。
*/
var MAX_SHEET_XML_BYTES = 167772160;
/** 共有文字列の上限 = 64 MiB（同じ文字は 1 回しか入らないので、シート本体より小さく見積もる） */
var MAX_SHARED_STRINGS_BYTES = 67108864;
/** 小さな部品の上限 = 16 MiB（styles.xml は書式が数万あっても数 MB。それ以上は異常） */
var MAX_SMALL_PART_BYTES = 16777216;
/** zip の中のファイル数の上限（Excel のブックは数十個。画像が多くても数千） */
var MAX_ENTRIES = 1e4;
/** シート数の上限（実データは 26 枚。Excel も実用上 1,000 枚で足りる） */
var MAX_SHEETS = 1e3;
/** 1 シートの結合セル数の上限（200 万セルのシートで全セルが 4 セル結合でも 50 万） */
var MAX_MERGES = 5e5;
/** 1 シートの図の数の上限（線・矢印は数えない。設計書の図は数十個。細工したファイルで DOM が膨らまない水準） */
var MAX_DRAWINGS = 1e3;
/** 1 シートの線・矢印の数の上限（図とは別に数える。線の多い構成図で画像・図形が押し出されないように） */
var MAX_LINES = 1e3;
/** 図形 1 つの文字数の上限（長い文章を入れた図形でも収まる水準。超えた分は切る） */
var MAX_SHAPE_TEXT_CHARS = 1e4;
/** 図形 1 つのラン（`a:br` を含む）の数の上限 */
var MAX_SHAPE_RUNS = 2e3;
/** 1 シートの図形の文字の合計の上限（文字数＋ラン＋段落。図形 1,000 個×1,000 文字の水準。細工したファイルで DOM が膨らまない） */
var MAX_SHEET_TEXT_CHARS = 1e6;
/** 画像 1 枚の上限 = 20 MiB（画面写しは数百 KB〜数 MB） */
var MAX_IMAGE_BYTES = 20971520;
/** 1 シートの画像の合計の上限 = 128 MiB（blob にして持つ量の上限。細工した zip 爆弾でメモリが尽きない水準） */
var MAX_SHEET_IMAGE_BYTES = 134217728;
/** 画像の 1 辺の上限 = 16,384 px */
var MAX_IMAGE_SIDE_PX = 16384;
/** 画像 1 枚の画素数の上限 = 64 MP（64 × 1,048,576。8,192 × 8,192 まで） */
var MAX_IMAGE_PIXELS = 67108864;
/** 図の名前・説明（`cNvPr` の name・descr）の文字数の上限。超えた分は切る */
var MAX_DRAWING_TEXT = 1e3;
/** 既定の上限の組 */
var DEFAULT_LIMITS = Object.freeze({
	maxSheetXmlBytes: MAX_SHEET_XML_BYTES,
	maxSharedStringsBytes: MAX_SHARED_STRINGS_BYTES,
	maxSmallPartBytes: MAX_SMALL_PART_BYTES,
	maxEntries: MAX_ENTRIES,
	maxSheets: MAX_SHEETS,
	maxMerges: MAX_MERGES,
	maxDrawings: MAX_DRAWINGS,
	maxLines: MAX_LINES,
	maxShapeTextChars: MAX_SHAPE_TEXT_CHARS,
	maxShapeParagraphs: 500,
	maxShapeRuns: MAX_SHAPE_RUNS,
	maxSheetTextChars: MAX_SHEET_TEXT_CHARS,
	maxGroupDepth: 8,
	maxImageBytes: MAX_IMAGE_BYTES,
	maxSheetImageBytes: MAX_SHEET_IMAGE_BYTES,
	maxImageSidePx: MAX_IMAGE_SIDE_PX,
	maxImagePixels: MAX_IMAGE_PIXELS,
	maxGeomGuides: 512,
	maxGeomCommands: MAX_GEOM_COMMANDS,
	maxSheetGeomUnits: 2e5,
	maxShapeAdjustments: 16,
	maxBorders: 65536,
	maxFonts: 65536,
	maxFills: 65536,
	maxCellFormats: 4096
});
/** Excel の行数の上限（1,048,576）。これを超える行番号は不正として読まない */
var MAX_ROWS = 1048576;
/** Excel の列数の上限（XFD = 16,384）。これを超える列は不正として読まない */
var MAX_COLS = 16384;
/** 1 つのセルの文字数の上限（Excel の上限 32,767）。超えた分は切る */
var MAX_CELL_TEXT = 32767;
/** 行高（`<row ht>`・`defaultRowHeight`、pt）の上限。Excel の上限 409.5。超えた値は丸める */
var MAX_ROW_PT = 409.5;
// src/core/xml.ts
/**
* XML を正規表現で走査する小さな道具と、セルの番地の計算。
* DOMParser に依存せず Node でも動かすため、タグは名前空間の接頭辞を許して照合する
* （`<c>` も `<x:c>` も読む。実験 05 §6）。
*/
/** 要素名の前に付いてよい名前空間の接頭辞（正規表現の断片。例 `x:`） */
var NS = String.raw`(?:[A-Za-z_][\w.-]*:)?`;
/** コード単位 1 つの文字参照を戻す。範囲外・不正な値は U+FFFD */
function codePointToString(n) {
	if (!Number.isFinite(n) || n < 0 || n > 1114111 || n >= 55296 && n <= 57343) return "�";
	return String.fromCodePoint(n);
}
var NAMED = Object.freeze({
	amp: "&",
	lt: "<",
	gt: ">",
	quot: "\"",
	apos: "'"
});
/** XML の実体参照（`&amp;` `&#x41;` `&#65;`）を戻す。属性値と、`_xHHHH_` を使わない値に使う */
function decodeEntities(s) {
	if (s.indexOf("&") < 0) return s;
	return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_m, e) => {
		if (e[0] === "#") return codePointToString(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
		return NAMED[e] ?? "";
	});
}
/**
* 文字列セルの本文を戻す。実体参照のあと、Excel の `_xHHHH_`（制御文字などの退避）を戻す。
* `_x005F_` は `_` を表す（本物の `_x0041_` という文字列を守るため）。
*/
function decodeText(s) {
	const e = decodeEntities(s);
	if (e.indexOf("_x") < 0) return e;
	return e.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h) => String.fromCharCode(parseInt(h, 16)));
}
/**
* `from` 以降で最初の `<name ...>…</name>`（名前空間の接頭辞可）を、線形時間で切り出す。
* 開きタグを探し、閉じタグは `indexOf` で探す（`[\s\S]*?` の正規表現は、閉じが無いと入れ子の走査で二乗になる）。
*/
function findElement(xml, name, from = 0) {
	const re = new RegExp(String.raw`<(${NS})${name}\b[^<>]*>`, "g");
	re.lastIndex = from;
	const m = re.exec(xml);
	if (!m) return null;
	const openEnd = m.index + m[0].length;
	if (m[0].endsWith("/>")) return {
		open: m[0],
		body: null,
		end: openEnd
	};
	const closeTag = `</${m[1] ?? ""}${name}>`;
	const close = xml.indexOf(closeTag, openEnd);
	if (close < 0) return {
		open: m[0],
		body: null,
		end: xml.length
	};
	return {
		open: m[0],
		body: xml.slice(openEnd, close),
		end: close + closeTag.length
	};
}
/** forEachElement の開きタグの正規表現（名前ごとに 1 つ。lastIndex は呼び出しのたびに設定し直すので、入れ子の呼び出しで共有しても壊れない） */
var openTagCache = /* @__PURE__ */ new Map();
/**
* `<name …>…</name>`（名前空間の接頭辞可）を出現順に callback へ渡す。同じ名前の入れ子は扱わない（`c`・`si`・`t` など入れ子にならない要素用）。
* 開きタグは `[^<>]*>`、閉じは `indexOf` で探すので、`>` や閉じの無い要素を大量に並べても時間は線形。
* 自己閉じ（`<c/>`）は body を null で渡す。閉じタグの無い要素に当たったら、以後に閉じが無いので打ち切る。
* 正規表現は名前ごとに 1 度だけ作る（セル 1 つごとに作らない）。callback の中で同じ `xml` を再び走査してよい。
*/
function forEachElement(xml, name, cb) {
	let re = openTagCache.get(name);
	if (re === void 0) {
		re = new RegExp(String.raw`<(${NS})${name}(?=[\s>/])[^<>]*>`, "g");
		openTagCache.set(name, re);
	}
	for (let pos = 0;;) {
		re.lastIndex = pos;
		const m = re.exec(xml);
		if (m === null) return;
		pos = m.index + m[0].length;
		if (m[0].endsWith("/>")) {
			cb(m[0], null);
			continue;
		}
		const closeTag = `</${m[1] ?? ""}${name}>`;
		const close = xml.indexOf(closeTag, pos);
		if (close < 0) return;
		cb(m[0], xml.slice(pos, close));
		pos = close + closeTag.length;
	}
}
/** findOpenTag の正規表現（名前ごとに 1 つ。lastIndex は呼び出しのたびに設定し直す） */
var findOpenTagCache = /* @__PURE__ */ new Map();
/**
* `from` 以降で最初の `<name …>` の開きタグだけを返す（閉じは探さない。名前空間の接頭辞可。無ければ null）。
* 開きタグは `[^<>]*>` で打ち切るので、`>` の無いタグを大量に並べても時間は線形。正規表現は名前ごとに 1 度だけ作る。
*/
function findOpenTag(xml, name, from = 0) {
	let re = findOpenTagCache.get(name);
	if (re === void 0) {
		re = new RegExp(String.raw`<${NS}${name}(?=[\s>/])[^<>]*>`, "g");
		findOpenTagCache.set(name, re);
	}
	re.lastIndex = from;
	return re.exec(xml)?.[0] ?? null;
}
/**
* タグ（`<c r="A1" s="2">` の属性の部分）から属性を取り出す。値は実体参照を戻す。
* 名前の先頭は後読みで「前が名前の文字でない位置」に限る（長い名前の途中から照合し直して二乗時間になるのを防ぐ。線形時間）。
*/
function parseAttrs(tag) {
	const o = Object.create(null);
	for (const m of tag.matchAll(/(?<![\w:.-])([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) o[m[1]] = decodeEntities(m[2] ?? m[3] ?? "");
	return o;
}
/** 名前空間の接頭辞を除いた属性の値（`r:id` と `rel:id` を同じに扱う）。`id` を探すとき `sheetId` は拾わない */
function attrByLocalName(attrs, local) {
	for (const k in attrs) if (k === local || k.endsWith(":" + local)) return attrs[k];
}
/** `A`→1、`Z`→26、`AA`→27、`XFD`→16384 */
function colNumber(letters) {
	let c = 0;
	for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
	return c;
}
/** 列番号（1 始まり）を列の記号にする。1→A、27→AA。1 未満・整数でなければ空文字 */
function colName(c) {
	if (!Number.isInteger(c) || c < 1) return "";
	let s = "";
	for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
	return s;
}
/** 番地 `AB12` の位置（1 始まり）。範囲外・不正なら null */
function parseCellRef(ref) {
	const m = /^([A-Z]{1,3})(\d{1,7})$/.exec(ref);
	if (!m) return null;
	const c = colNumber(m[1]);
	const r = parseInt(m[2], 10);
	if (c < 1 || c > 16384 || r < 1 || r > 1048576) return null;
	return {
		r,
		c
	};
}
/** 列だけの番地（`AB`）。セルの r 属性（`AB12`）から列だけ欲しいとき用。範囲外・不正なら 0 */
function colOfRef(ref) {
	const m = /^([A-Z]{1,3})\d*$/.exec(ref);
	if (!m) return 0;
	const c = colNumber(m[1]);
	return c >= 1 && c <= 16384 ? c : 0;
}
/** 開きタグの名前の直後の文字として許す文字（`<c` の後が `o`（`<cols`）などでないことの確認用） */
function isTagDelimiter(ch) {
	return ch === " " || ch === ">" || ch === "/" || ch === "	" || ch === "\n" || ch === "\r";
}
/** Strict 形式（ISO 29500 Strict）の名前空間の目印。1 次では読まない（unsupported） */
var STRICT_NS = "purl.oclc.org/ooxml";
/** ルート要素の宣言（先頭の 2,000 文字）に Strict の名前空間があるか */
function looksStrict(xml) {
	return xml.slice(0, 2e3).includes(STRICT_NS);
}
// src/core/types.ts
/**
* xlsx を読めないときの例外。種類は `kind` で決まる。
* メッセージには、原本の中身（セルの値・シート名）を入れない。
*/
var XlsxError = class extends Error {
	kind;
	constructor(kind, detail) {
		super(detail ? `xlsx: ${kind} (${detail})` : `xlsx: ${kind}`);
		this.name = "XlsxError";
		this.kind = kind;
	}
};
/** 既定の最大桁幅 MDW（px。Calibri 11 相当） */
var DEFAULT_MDW = 7;
/** 寸法の指定が何も無いシートの layout（既定値のみ） */
function emptyLayout(mdw = 7, fontRowPx = null) {
	return {
		cols: [],
		defaultColWidth: null,
		baseColWidth: 8,
		defaultRowHeight: 15,
		customHeight: false,
		zeroHeight: false,
		rows: [],
		mdw,
		fontRowPx
	};
}
// src/core/theme.ts
/** テーマの色の位置（`a:clrScheme` の子の要素名）。この順で 12 色 */
var THEME_SLOTS = [
	"dk1",
	"lt1",
	"dk2",
	"lt2",
	"accent1",
	"accent2",
	"accent3",
	"accent4",
	"accent5",
	"accent6",
	"hlink",
	"folHlink"
];
/** 既定の 12 色（Office 2013〜2022。Excel 16 が書く theme1.xml と同じ値。`THEME_SLOTS` の順） */
var DEFAULT_VALUES = [
	0,
	16777215,
	4478058,
	15197926,
	4485828,
	15564081,
	10855845,
	16760832,
	6003669,
	7384391,
	353217,
	9785202
];
/** 既定のテーマ（共有する。変更しない） */
var DEFAULT_THEME = new Map(THEME_SLOTS.map((s, i) => [s, DEFAULT_VALUES[i]]));
/** 16 進 6 桁（`[0-9A-Fa-f]{6}`）だけを整数にする。それ以外・未指定は null */
function parseHex6(s) {
	return s !== void 0 && /^[0-9A-Fa-f]{6}$/.test(s) ? parseInt(s, 16) : null;
}
/** `sysClr` の色：`lastClr`（16 進 6 桁）。無い・不正なら `windowText`＝黒・`window`＝白、その他は null */
function sysClrValue(attrs) {
	const last = parseHex6(attrs.lastClr);
	if (last !== null) return last;
	if (attrs.val === "windowText") return 0;
	if (attrs.val === "window") return 16777215;
	return null;
}
/** 位置の中の色の要素（`srgbClr`・`sysClr`）の開きタグ。`[^<>]*>` で打ち切るので線形時間 */
var SLOT_COLOR_RE = new RegExp(String.raw`<${NS}(srgbClr|sysClr)(?=[\s>/])[^<>]*>`);
/** 位置の中身（`<a:srgbClr val="…"/>` など）から色を取る。解けなければ null */
function slotColor(body) {
	const m = SLOT_COLOR_RE.exec(body);
	if (m === null) return null;
	const a = parseAttrs(m[0]);
	return m[1] === "srgbClr" ? parseHex6(a.val) : sysClrValue(a);
}
/**
* `theme1.xml` の `a:clrScheme` から 12 色を読む。欠けた色・不正な色はその色だけ既定で補う。
* xml が null・`clrScheme` が無い・閉じが無いときは既定の 12 色（新しい `Map`）。
*/
function parseTheme(xml) {
	const out = new Map(DEFAULT_THEME);
	if (xml === null) return out;
	const scheme = findElement(xml, "clrScheme")?.body;
	if (scheme === null || scheme === void 0) return out;
	for (const slot of THEME_SLOTS) {
		const body = findElement(scheme, slot)?.body;
		if (body === null || body === void 0) continue;
		const rgb = slotColor(body);
		if (rgb !== null) out.set(slot, rgb);
	}
	return out;
}
/**
* テーマを読む。`path` が無い・部品が zip に無い・上限（小さな部品）を超える・読めないときは既定の 12 色。
* 画面を閉じた後の読み取り（`broken`・`closed`）だけは呼び出し側へ投げる（index.ts が閉じたかを確かめる）。
*/
async function loadTheme(pkg, path, limits) {
	if (path === null || !pkg.has(path)) return parseTheme(null);
	try {
		return parseTheme(await pkg.readText(path, limits.maxSmallPartBytes));
	} catch (e) {
		if (e instanceof XlsxError && e.kind === "broken" && /closed/.test(e.message)) throw e;
		return parseTheme(null);
	}
}
var clamp01 = (x) => x > 0 ? x < 1 ? x : 1 : 0;
/** sRGB の成分（0〜1）→ 線形 RGB（0〜1）。範囲外は収め、NaN は 0 */
function toLinear(c, gamma = null) {
	const v = clamp01(c);
	if (gamma !== null) return Math.pow(v, gamma);
	return v <= .04045 ? v / 12.92 : Math.pow((v + .055) / 1.055, 2.4);
}
/** 線形 RGB（0〜1）→ sRGB の成分（0〜1）。範囲外は収め、NaN は 0 */
function fromLinear(l, gamma = null) {
	const v = clamp01(l);
	if (gamma !== null) return Math.pow(v, 1 / gamma);
	return v <= .0031308 ? v * 12.92 : 1.055 * Math.pow(v, 1 / 2.4) - .055;
}
/**
* path の `fill` の濃淡：`shade`（黒へ寄せる）か `tint`（白へ寄せる）の係数（0〜1）。
* Excel の画素から決めた値：darken＝0.6・lighten＝0.6 は成分ごとの比がちょうど 0.6、darkenLess・lightenLess は 0.802〜0.809 に揃うので 0.805。
* 濃淡は `applyFillMode` が **sRGB の値のまま**掛ける（線形 RGB にすると Excel の面より暗く／明るくなりすぎる）。
*/
var FILL_MODE_FACTORS = /* @__PURE__ */ new Map([
	["darken", {
		kind: "shade",
		factor: .6
	}],
	["darkenLess", {
		kind: "shade",
		factor: .805
	}],
	["lighten", {
		kind: "tint",
		factor: .6
	}],
	["lightenLess", {
		kind: "tint",
		factor: .805
	}]
]);
/** 整数だけの文字列（`-` 可）を数にする。整数でない・巨大で有限でない・未指定は null */
function parseIntAttr(s) {
	if (s === void 0 || !/^-?\d+$/.test(s)) return null;
	const n = Number(s);
	return Number.isFinite(n) ? n : null;
}
/** 成分（0〜1）を 0〜255 の整数にする */
var to255 = (c) => Math.round(clamp01(c) * 255);
/** RGB の整数 → 成分 3 つ（0〜1） */
var unpack = (rgb) => [
	(rgb >> 16 & 255) / 255,
	(rgb >> 8 & 255) / 255,
	(rgb & 255) / 255
];
/** 成分 3 つ（0〜1）→ RGB の整数 */
var pack = (r, g, b) => to255(r) << 16 | to255(g) << 8 | to255(b);
/** RGB（0〜1）→ HSL（色相は度 0〜360・彩度・明るさは 0〜1） */
function rgbToHsl(r, g, b) {
	const max = Math.max(r, g, b);
	const min = Math.min(r, g, b);
	const l = (max + min) / 2;
	const d = max - min;
	if (d === 0) return [
		0,
		0,
		l
	];
	const s = d / (1 - Math.abs(2 * l - 1));
	let hue;
	if (max === r) hue = (g - b) / d % 6;
	else if (max === g) hue = (b - r) / d + 2;
	else hue = (r - g) / d + 4;
	return [
		wrapHue(hue * 60),
		clamp01(s),
		l
	];
}
/** 色相を 0 以上 360 未満に回す（有限でなければ 0） */
function wrapHue(deg) {
	if (!Number.isFinite(deg)) return 0;
	const m = deg % 360;
	return m < 0 ? m + 360 : m;
}
/** HSL → RGB（0〜1）。色相は 360° で回し、彩度・明るさは収める */
function hslToRgb(hue, sat, lum) {
	const s = clamp01(sat);
	const l = clamp01(lum);
	const c = (1 - Math.abs(2 * l - 1)) * s;
	const hp = wrapHue(hue) / 60;
	const x = c * (1 - Math.abs(hp % 2 - 1));
	let r;
	let g;
	let b;
	if (hp < 1) [r, g, b] = [
		c,
		x,
		0
	];
	else if (hp < 2) [r, g, b] = [
		x,
		c,
		0
	];
	else if (hp < 3) [r, g, b] = [
		0,
		c,
		x
	];
	else if (hp < 4) [r, g, b] = [
		0,
		x,
		c
	];
	else if (hp < 5) [r, g, b] = [
		x,
		0,
		c
	];
	else [r, g, b] = [
		c,
		0,
		x
	];
	const m = l - c / 2;
	return [
		r + m,
		g + m,
		b + m
	];
}
/** `ST_PresetColorVal` の色（名前 → RGB）。dk/lt/med の別名（dark/light/medium）と Gray の別綴り（Grey）は下で足す */
var PRESET_BASE = [
	["aliceBlue", 15792383],
	["antiqueWhite", 16444375],
	["aqua", 65535],
	["aquamarine", 8388564],
	["azure", 15794175],
	["beige", 16119260],
	["bisque", 16770244],
	["black", 0],
	["blanchedAlmond", 16772045],
	["blue", 255],
	["blueViolet", 9055202],
	["brown", 10824234],
	["burlyWood", 14596231],
	["cadetBlue", 6266528],
	["chartreuse", 8388352],
	["chocolate", 13789470],
	["coral", 16744272],
	["cornflowerBlue", 6591981],
	["cornsilk", 16775388],
	["crimson", 14423100],
	["cyan", 65535],
	["deepPink", 16716947],
	["deepSkyBlue", 49151],
	["dimGray", 6908265],
	["dkBlue", 139],
	["dkCyan", 35723],
	["dkGoldenrod", 12092939],
	["dkGray", 11119017],
	["dkGreen", 25600],
	["dkKhaki", 12433259],
	["dkMagenta", 9109643],
	["dkOliveGreen", 5597999],
	["dkOrange", 16747520],
	["dkOrchid", 10040012],
	["dkRed", 9109504],
	["dkSalmon", 15308410],
	["dkSeaGreen", 9419915],
	["dkSlateBlue", 4734347],
	["dkSlateGray", 3100495],
	["dkTurquoise", 52945],
	["dkViolet", 9699539],
	["dodgerBlue", 2003199],
	["firebrick", 11674146],
	["floralWhite", 16775920],
	["forestGreen", 2263842],
	["fuchsia", 16711935],
	["gainsboro", 14474460],
	["ghostWhite", 16316671],
	["gold", 16766720],
	["goldenrod", 14329120],
	["gray", 8421504],
	["green", 32768],
	["greenYellow", 11403055],
	["honeydew", 15794160],
	["hotPink", 16738740],
	["indianRed", 13458524],
	["indigo", 4915330],
	["ivory", 16777200],
	["khaki", 15787660],
	["lavender", 15132410],
	["lavenderBlush", 16773365],
	["lawnGreen", 8190976],
	["lemonChiffon", 16775885],
	["lime", 65280],
	["limeGreen", 3329330],
	["linen", 16445670],
	["ltBlue", 11393254],
	["ltCoral", 15761536],
	["ltCyan", 14745599],
	["ltGoldenrodYellow", 16448210],
	["ltGray", 13882323],
	["ltGreen", 9498256],
	["ltPink", 16758465],
	["ltSalmon", 16752762],
	["ltSeaGreen", 2142890],
	["ltSkyBlue", 8900346],
	["ltSlateGray", 7833753],
	["ltSteelBlue", 11584734],
	["ltYellow", 16777184],
	["magenta", 16711935],
	["maroon", 8388608],
	["medAquamarine", 6737322],
	["medBlue", 205],
	["medOrchid", 12211667],
	["medPurple", 9662683],
	["medSeaGreen", 3978097],
	["medSlateBlue", 8087790],
	["medSpringGreen", 64154],
	["medTurquoise", 4772300],
	["medVioletRed", 13047173],
	["midnightBlue", 1644912],
	["mintCream", 16121850],
	["mistyRose", 16770273],
	["moccasin", 16770229],
	["navajoWhite", 16768685],
	["navy", 128],
	["oldLace", 16643558],
	["olive", 8421376],
	["oliveDrab", 7048739],
	["orange", 16753920],
	["orangeRed", 16729344],
	["orchid", 14315734],
	["paleGoldenrod", 15657130],
	["paleGreen", 10025880],
	["paleTurquoise", 11529966],
	["paleVioletRed", 14381203],
	["papayaWhip", 16773077],
	["peachPuff", 16767673],
	["peru", 13468991],
	["pink", 16761035],
	["plum", 14524637],
	["powderBlue", 11591910],
	["purple", 8388736],
	["red", 16711680],
	["rosyBrown", 12357519],
	["royalBlue", 4286945],
	["saddleBrown", 9127187],
	["salmon", 16416882],
	["sandyBrown", 16032864],
	["seaGreen", 3050327],
	["seaShell", 16774638],
	["sienna", 10506797],
	["silver", 12632256],
	["skyBlue", 8900331],
	["slateBlue", 6970061],
	["slateGray", 7372944],
	["snow", 16775930],
	["springGreen", 65407],
	["steelBlue", 4620980],
	["tan", 13808780],
	["teal", 32896],
	["thistle", 14204888],
	["tomato", 16737095],
	["turquoise", 4251856],
	["violet", 15631086],
	["wheat", 16113331],
	["white", 16777215],
	["whiteSmoke", 16119285],
	["yellow", 16776960],
	["yellowGreen", 10145074]
];
/** プリセット色の表（別名を足した後）。キーは名前なので Map（`__proto__` などで汚染されない） */
var PRESET_COLORS = (() => {
	const m = /* @__PURE__ */ new Map();
	for (const [name, rgb] of PRESET_BASE) {
		m.set(name, rgb);
		const long = name.replace(/^dk(?=[A-Z])/, "dark").replace(/^lt(?=[A-Z])/, "light").replace(/^med(?=[A-Z])/, "medium");
		m.set(long, rgb);
		for (const n of [name, long]) if (n.endsWith("Gray")) m.set(n.slice(0, -4) + "Grey", rgb);
	}
	return m;
})();
PRESET_COLORS.size;
/** `prstClr` の名前から RGB の整数を引く。未知の名前は null */
function lookupPresetColor(name) {
	return PRESET_COLORS.get(name) ?? null;
}
/** `schemeClr` の名前 → テーマの色の位置（スプレッドシートは `clrMap` を持たないため、既定の対応で固定）。`phClr`・未知の名前は無い */
var SCHEME_SLOTS = /* @__PURE__ */ new Map([
	["dk1", "dk1"],
	["lt1", "lt1"],
	["dk2", "dk2"],
	["lt2", "lt2"],
	["accent1", "accent1"],
	["accent2", "accent2"],
	["accent3", "accent3"],
	["accent4", "accent4"],
	["accent5", "accent5"],
	["accent6", "accent6"],
	["hlink", "hlink"],
	["folHlink", "folHlink"],
	["bg1", "lt1"],
	["tx1", "dk1"],
	["bg2", "lt2"],
	["tx2", "dk2"]
]);
/** 色の要素 6 種の開きタグ（名前を捕まえる）。`[^<>]*>` で打ち切るので、`>` の無いタグを並べても時間は線形 */
var COLOR_OPEN_RE = new RegExp(String.raw`<(${NS})(srgbClr|schemeClr|sysClr|scrgbClr|hslClr|prstClr)(?=[\s>/])[^<>]*>`);
/** 子の要素の開きタグ（名前を捕まえる）。変換の子の列を読む */
var CHILD_OPEN_RE = new RegExp(String.raw`<${NS}([A-Za-z][\w.-]*)(?=[\s>/])[^<>]*>`, "g");
/**
* `xml`（例 `<a:solidFill>` の中身）で最初に見つかる色の要素を切り出す。無ければ null。
* 閉じタグは `indexOf` で探す（閉じが無ければ body は null＝子の変換を読まない）。
*/
function findColorElement(xml) {
	const m = COLOR_OPEN_RE.exec(xml);
	if (m === null) return null;
	const open = m[0];
	const name = m[2];
	if (open.endsWith("/>")) return {
		name,
		open,
		body: null
	};
	const start = m.index + open.length;
	const close = xml.indexOf(`</${m[1] ?? ""}${name}>`, start);
	return {
		name,
		open,
		body: close < 0 ? null : xml.slice(start, close)
	};
}
/** 色の要素の基本の色（変換の前）。解けなければ null */
function baseColor(el, theme) {
	const a = parseAttrs(el.open);
	switch (el.name) {
		case "srgbClr": return parseHex6(a.val);
		case "schemeClr": {
			const slot = a.val === void 0 ? void 0 : SCHEME_SLOTS.get(a.val);
			return slot === void 0 ? null : theme.get(slot) ?? DEFAULT_THEME.get(slot) ?? null;
		}
		case "sysClr": return sysClrValue(a);
		case "scrgbClr": {
			const r = parseIntAttr(a.r);
			const g = parseIntAttr(a.g);
			const b = parseIntAttr(a.b);
			if (r === null || g === null || b === null) return null;
			return pack(fromLinear(r / 1e5), fromLinear(g / 1e5), fromLinear(b / 1e5));
		}
		case "hslClr": {
			const hue = parseIntAttr(a.hue);
			const sat = parseIntAttr(a.sat);
			const lum = parseIntAttr(a.lum);
			if (hue === null || sat === null || lum === null) return null;
			const [r, g, b] = hslToRgb(hue / 6e4, sat / 1e5, lum / 1e5);
			return pack(r, g, b);
		}
		case "prstClr": return a.val === void 0 ? null : lookupPresetColor(a.val);
		default: return null;
	}
}
/** HSL の 1 成分の変換（lum・sat・hue の系統）。値は 1/1000 %（色相は 1/60000 度）の整数 */
function applyHsl(w, kind, op, v) {
	const hsl = rgbToHsl(w.r, w.g, w.b);
	const i = kind === "hue" ? 0 : kind === "sat" ? 1 : 2;
	const x = hsl[i];
	const unit = kind === "hue" ? 6e4 : 1e5;
	if (op === "set") hsl[i] = v / unit;
	else if (op === "mod") hsl[i] = x * v / 1e5;
	else hsl[i] = x + v / unit;
	if (i === 0) hsl[0] = wrapHue(hsl[0]);
	else hsl[i] = clamp01(hsl[i]);
	[w.r, w.g, w.b] = hslToRgb(hsl[0], hsl[1], hsl[2]);
}
/** 線形 RGB で `c′ = c × shade`（黒へ寄せる）。係数は 0 以上 */
function applyShade(w, factor) {
	const k = Math.max(0, factor);
	w.r = fromLinear(toLinear(w.r) * k);
	w.g = fromLinear(toLinear(w.g) * k);
	w.b = fromLinear(toLinear(w.b) * k);
}
/** 線形 RGB で `c′ = 1 − (1 − c) × tint`（白へ寄せる）。係数は 0 以上 */
function applyTint(w, factor) {
	const k = Math.max(0, factor);
	w.r = fromLinear(1 - (1 - toLinear(w.r)) * k);
	w.g = fromLinear(1 - (1 - toLinear(w.g)) * k);
	w.b = fromLinear(1 - (1 - toLinear(w.b)) * k);
}
/**
* 基本の色に、子の変換を出現順に掛ける（ADR-0028 §決定 6.3）。
* 読む子は最初の `MAX_COLOR_TRANSFORMS` 個まで（無視する要素・不正な値も 1 個と数える）。
* 各段の結果は 0〜1 に収める。値が整数でない変換は飛ばす。無視する変換（`comp`・`inv`・`gray`・`gamma`・`invGamma`・`red`/`green`/`blue` の系統）は色を変えない。
*/
function applyTransforms(base, body) {
	if (body === null) return base;
	const [r, g, b] = unpack(base.rgb);
	const w = {
		r,
		g,
		b,
		alpha: base.alpha
	};
	CHILD_OPEN_RE.lastIndex = 0;
	for (let count = 0; count < 16; count++) {
		const m = CHILD_OPEN_RE.exec(body);
		if (m === null) break;
		const v = parseIntAttr(parseAttrs(m[0]).val);
		if (v === null) continue;
		switch (m[1]) {
			case "lum":
				applyHsl(w, "lum", "set", v);
				break;
			case "lumMod":
				applyHsl(w, "lum", "mod", v);
				break;
			case "lumOff":
				applyHsl(w, "lum", "off", v);
				break;
			case "sat":
				applyHsl(w, "sat", "set", v);
				break;
			case "satMod":
				applyHsl(w, "sat", "mod", v);
				break;
			case "satOff":
				applyHsl(w, "sat", "off", v);
				break;
			case "hue":
				applyHsl(w, "hue", "set", v);
				break;
			case "hueMod":
				applyHsl(w, "hue", "mod", v);
				break;
			case "hueOff":
				applyHsl(w, "hue", "off", v);
				break;
			case "tint":
				applyTint(w, v / 1e5);
				break;
			case "shade":
				applyShade(w, v / 1e5);
				break;
			case "alpha":
				w.alpha = clamp01(v / 1e5);
				break;
			case "alphaMod":
				w.alpha = clamp01(w.alpha * v / 1e5);
				break;
			case "alphaOff": w.alpha = clamp01(w.alpha + v / 1e5);
		}
	}
	CHILD_OPEN_RE.lastIndex = 0;
	return {
		rgb: pack(w.r, w.g, w.b),
		alpha: clamp01(w.alpha)
	};
}
/**
* 切り出した色の要素 1 つを解く（基本の色 → 子の変換）。解けなければ null（例外にしない）。
* 解けないときに次の出どころへ進むのは呼び出し側（ISSUE-95）。
*/
function resolveColorElement(el, theme) {
	const rgb = baseColor(el, theme);
	if (rgb === null) return null;
	return applyTransforms({
		rgb,
		alpha: 1
	}, el.body);
}
/**
* `xml`（`<a:solidFill>` の中身など）の最初の色の要素を解く。色の要素が無い・解けなければ null。
* 色は RGB の整数と不透明度で返す（ファイルの文字列は返さない）。
*/
function resolveColor(xml, theme) {
	const el = findColorElement(xml);
	return el === null ? null : resolveColorElement(el, theme);
}
/**
* path の `fill`（`norm`・`darken`・`darkenLess`・`lighten`・`lightenLess`）を色に掛ける。
* 係数は `FILL_MODE_FACTORS`（U8）。未知の名前・`norm` は色のまま。不透明度は変えない。
* 掛けるのは sRGB の値のまま（shade＝`c × k`・tint＝`1 − (1 − c) × k`）。変換の `shade`・`tint`（線形 RGB）とは別の式で、Excel の面の画素と合う方を採った。
*/
function applyFillMode(color, mode) {
	const f = FILL_MODE_FACTORS.get(mode);
	if (f === void 0) return color;
	const k = Math.max(0, f.factor);
	const [r, g, b] = unpack(color.rgb);
	const one = (c) => f.kind === "shade" ? c * k : 1 - (1 - c) * k;
	return {
		rgb: pack(one(r), one(g), one(b)),
		alpha: color.alpha
	};
}
/**
* 2 色を **線形 RGB** で割合を混ぜる（`t`＝0 で a、1 で b。0〜1 に収め、NaN は 0）。不透明度は sRGB と同じ線形補間。
* Excel の 2 stop のグラデーションの中央の色がこの混色と合う（ISSUE-95・U7。3 stop 以上は sRGB の混色 `mixColors` と合う）。
*/
function mixColorsLinear(a, b, t) {
	const k = clamp01(t);
	const [ar, ag, ab] = unpack(a.rgb).map((c) => toLinear(c));
	const [br, bg, bb] = unpack(b.rgb).map((c) => toLinear(c));
	return {
		rgb: pack(fromLinear(ar + (br - ar) * k), fromLinear(ag + (bg - ag) * k), fromLinear(ab + (bb - ab) * k)),
		alpha: clamp01(a.alpha + (b.alpha - a.alpha) * k)
	};
}
/**
* 2 色を割合で混ぜる（`t`＝0 で a、1 で b。0〜1 に収め、NaN は 0）。各成分は sRGB の値の線形補間で、不透明度も混ぜる。
* グラデーションの補間・模様の混色に使う（適用は ISSUE-95）。
*/
function mixColors(a, b, t) {
	const k = clamp01(t);
	const [ar, ag, ab] = unpack(a.rgb);
	const [br, bg, bb] = unpack(b.rgb);
	return {
		rgb: pack(ar + (br - ar) * k, ag + (bg - ag) * k, ab + (bb - ab) * k),
		alpha: clamp01(a.alpha + (b.alpha - a.alpha) * k)
	};
}
export { parseCellRef as C, MAX_DRAWING_TEXT as D, MAX_COLS as E, MAX_ROWS as O, parseAttrs as S, MAX_CELL_TEXT as T, findElement as _, DEFAULT_THEME as a, isTagDelimiter as b, DEFAULT_MDW as c, NS as d, attrByLocalName as f, decodeText as g, decodeEntities as h, resolveColor as i, MAX_ROW_PT as k, XlsxError as l, colOfRef as m, mixColors as n, loadTheme as o, colName as p, mixColorsLinear as r, parseHex6 as s, applyFillMode as t, emptyLayout as u, findOpenTag as v, DEFAULT_LIMITS as w, looksStrict as x, forEachElement as y };
