// src/render/selection-copy.ts
/** コピーするセルの数の上限（見える行 × 見える列） */
var MAX_COPY_CELLS = 1e6;
/** 行を読む単位 */
var READ_CHUNK = 512;
/** コピーする範囲。行の全体は値のある最後の列まで、列の全体は値のある最後の行まで（空なら r2 < r1 か c2 < c1） */
function copyRange(sel, used) {
	const { r1, c1, r2, c2 } = sel.range;
	return {
		r1,
		c1,
		r2: sel.type === "column" ? Math.min(r2, used.rows) : r2,
		c2: sel.type === "row" ? Math.min(c2, used.cols) : c2
	};
}
/** セル 1 つの文字（タブ・改行・`"` を含めば `"` で囲む） */
function quoteCell(text) {
	return /[\t\r\n"]/.test(text) ? `"${text.replace(/"/g, "\"\"")}"` : text;
}
/**
* 範囲の文字。見える行 × 見える列が `maxCells` を超えたら null。
* 結合の左上が範囲（と読む区切り）より上にあると、`getRows` は区切りの最初の行に左上の文字を置くので、
* 1 行前から読んで捨てる（範囲の中の続きの行は空のまま）。
*/
function rangeText(sheet, range, geo, maxCells = MAX_COPY_CELLS) {
	const cols = [];
	for (let c = range.c1; c <= range.c2; c++) if (geo.colWidth(c) > 0) cols.push(c);
	let shownRows = 0;
	for (let r = range.r1; r <= range.r2; r++) if (geo.rowHeight(r) > 0 && ++shownRows * cols.length > maxCells) return null;
	if (shownRows === 0 || cols.length === 0) return "";
	const lines = [];
	for (let a = range.r1; a <= range.r2; a += READ_CHUNK) {
		const b = Math.min(range.r2, a + READ_CHUNK - 1);
		const from = Math.max(1, a - 1);
		const data = a <= sheet.rowCount ? sheet.getRows(from, Math.min(b, sheet.rowCount)) : [];
		for (let r = a; r <= b; r++) {
			if (geo.rowHeight(r) === 0) continue;
			const cells = data[r - from]?.cells;
			lines.push(cols.map((c) => quoteCell(cells?.[c - 1]?.text ?? "")).join("	"));
		}
	}
	return lines.join("\r\n") + "\r\n";
}
// src/render/selection-keys.ts
/** 選択を動かすキー（Ctrl・Shift・Alt の扱いは {@link keyTarget}） */
var SELECTION_KEYS = /* @__PURE__ */ new Set([
	"ArrowUp",
	"ArrowDown",
	"ArrowLeft",
	"ArrowRight",
	"PageUp",
	"PageDown",
	"Home",
	"End"
]);
/**
* from の次（dir 方向）の、大きさが 0 でない番号（1〜max）。無ければ null。
* from が範囲の外（max より先）でも、範囲の中へは戻らない（null）。
*/
function nextVisible(from, dir, max, size) {
	for (let i = from + dir; i >= 1 && i <= max; i += dir) if (size(i) > 0) return i;
	return null;
}
/**
* キーで動かした先の今のセル。扱わないキーは undefined、端で動けないときは null。
* @param ctrl Ctrl（Mac の ⌘ も呼ぶ側で渡す）を押しているか
*/
function keyTarget(key, ctrl, sel, ctx) {
	const { geo: g, extent: ext } = ctx;
	const rowH = (r) => g.rowHeight(r);
	const colW = (c) => g.colWidth(c);
	const { type, range, active } = sel;
	const at = (r, c) => r === null || c === null ? null : {
		r,
		c
	};
	switch (key) {
		case "ArrowDown": return at(nextVisible(type === "column" ? active.r : range.r2, 1, ext.rows, rowH), active.c);
		case "ArrowUp": return at(nextVisible(type === "column" ? active.r : range.r1, -1, ext.rows, rowH), active.c);
		case "ArrowRight": return at(active.r, nextVisible(type === "row" ? active.c : range.c2, 1, ext.cols, colW));
		case "ArrowLeft": return at(active.r, nextVisible(type === "row" ? active.c : range.c1, -1, ext.cols, colW));
		case "PageDown":
		case "PageUp": {
			const dir = key === "PageDown" ? 1 : -1;
			const y = g.rowTop(active.r) + dir * Math.max(1, ctx.pageHeight);
			let r = Math.min(ext.rows, Math.max(1, g.rowAt(Math.min(g.totalHeight - 1, Math.max(0, y)))));
			if (rowH(r) === 0) r = nextVisible(r, dir, ext.rows, rowH) ?? nextVisible(r, dir === 1 ? -1 : 1, ext.rows, rowH) ?? r;
			if (r === active.r) return at(nextVisible(active.r, dir, ext.rows, rowH), active.c);
			return {
				r,
				c: active.c
			};
		}
		case "Home": return ctrl ? at(nextVisible(0, 1, ext.rows, rowH), nextVisible(0, 1, ext.cols, colW)) : at(active.r, nextVisible(0, 1, ext.cols, colW));
		case "End": {
			if (!ctrl) return void 0;
			const rows = Math.min(ext.rows, Math.max(1, ctx.used.rows));
			const cols = Math.min(ext.cols, Math.max(1, ctx.used.cols));
			return at(nextVisible(rows + 1, -1, rows, rowH), nextVisible(cols + 1, -1, cols, colW));
		}
		default: return;
	}
}
// src/render/selection.ts
/** Excel の最大の行数・列数（grid.ts の `MAX_GRID_ROWS`・`MAX_GRID_COLS` と同じ） */
var MAX_GRID_ROWS = 1048576;
var MAX_GRID_COLS = 16384;
/** 結合を探すときに上下へ読む行数（これより上から始まる結合はここからの、これより高い結合はここまでの範囲として扱う） */
var MERGE_LOOKBACK = 100;
/** クリックとみなす、押してから離すまでに動いた距離（px） */
var CLICK_SLOP = 4;
/** 1 以上 max 以下の整数か */
var inRange = (v, max) => typeof v === "number" && Number.isInteger(v) && v >= 1 && v <= max;
/**
* 行 r・列 c を含む結合の範囲（無ければそのセル 1 つ）。シートの行を上へ辿って、結合の左上を探す。
* `getRows` は範囲より上で始まる結合の左上を範囲の最初の行に置くので、`MERGE_LOOKBACK` 行より上で始まる結合は、そこからの範囲になる。
*/
function mergeRangeAt(sheet, r, c) {
	const one = {
		r1: r,
		c1: c,
		r2: r,
		c2: c
	};
	if (r > sheet.rowCount) return one;
	const top = Math.max(1, r - MERGE_LOOKBACK + 1);
	const rows = sheet.getRows(top, Math.min(sheet.rowCount, r + MERGE_LOOKBACK - 1));
	for (let i = r - top; i >= 0; i--) {
		const row = rows[i];
		let k = c - 1;
		while (k >= 0 && row.cells[k] === null) k--;
		const cell = k >= 0 ? row.cells[k] : void 0;
		if (cell && (cell.colSpan > 1 || cell.rowSpan > 1) && k + cell.colSpan > c - 1 && row.r + cell.rowSpan > r) return {
			r1: row.r,
			c1: k + 1,
			r2: row.r + cell.rowSpan - 1,
			c2: k + cell.colSpan
		};
		if (row.cells[c - 1] !== null) break;
	}
	return one;
}
/** 2 つの選択が同じか */
function sameSelection(a, b) {
	if (a === null || b === null) return a === b;
	return a.type === b.type && a.range.r1 === b.range.r1 && a.range.c1 === b.range.c1 && a.range.r2 === b.range.r2 && a.range.c2 === b.range.c2 && a.active.r === b.active.r && a.active.c === b.active.c;
}
/**
* 選択を作る。`mode` が無ければクリックでは選ばない（`set` だけが効く）。
* @param mode セルをクリックしたときに選ぶ物
*/
function createSelection(host, mode) {
	let value = null;
	const band = document.createElement("div");
	band.className = "xs-sel";
	band.setAttribute("aria-hidden", "true");
	band.hidden = true;
	/** 印を付けた見出しのセル（付け直す前に外す） */
	let marked = [];
	/** 選択の番地の読み上げ（見えない要素） */
	const live = document.createElement("span");
	live.className = "xs-live";
	live.setAttribute("aria-live", "polite");
	host.root.append(live);
	/** 行 r1〜r2 の全体（結合の行の範囲に広げない） */
	const rowSel = (r1, r2, c) => ({
		type: "row",
		range: {
			r1: Math.min(r1, r2),
			c1: 1,
			r2: Math.max(r1, r2),
			c2: MAX_GRID_COLS
		},
		active: {
			r: Math.min(r1, r2),
			c
		}
	});
	/** 列 c1〜c2 の全体 */
	const colSel = (c1, c2, r) => ({
		type: "column",
		range: {
			r1: 1,
			c1: Math.min(c1, c2),
			r2: MAX_GRID_ROWS,
			c2: Math.max(c1, c2)
		},
		active: {
			r,
			c: Math.min(c1, c2)
		}
	});
	/** セル（結合なら結合の範囲）。asRow なら、その行の範囲の全体 */
	const cellSel = (r, c, asRow) => {
		const sh = host.sheet();
		const m = sh ? mergeRangeAt(sh, r, c) : {
			r1: r,
			c1: c,
			r2: r,
			c2: c
		};
		const active = {
			r: m.r1,
			c: m.c1
		};
		return asRow ? {
			type: "row",
			range: {
				r1: m.r1,
				c1: 1,
				r2: m.r2,
				c2: MAX_GRID_COLS
			},
			active
		} : {
			type: "cell",
			range: m,
			active
		};
	};
	const update = (next, origin) => {
		if (sameSelection(value, next)) return;
		value = next;
		controller.layout();
		live.textContent = next ? address(next) : "";
		host.changed(value, origin);
	};
	/** 選択の番地（`C5`・`B2:D4`・`5:5`・`C:D`） */
	const address = (sel) => {
		const { r1, c1, r2, c2 } = sel.range;
		if (sel.type === "row") return `${r1}:${r2}`;
		if (sel.type === "column") return `${host.colName(c1)}:${host.colName(c2)}`;
		const a = `${host.colName(c1)}${r1}`;
		return r1 === r2 && c1 === c2 ? a : `${a}:${host.colName(c2)}${r2}`;
	};
	/** 選択（今のセル）が見える所まで、最小の量だけスクロールする（中央に寄せない。Excel と同じ） */
	const reveal = (sel) => {
		const g = host.geometry();
		if (!g) return;
		const body = host.body;
		const o = host.origin();
		const ext = host.extent();
		const fit = (pos, a, b, inner) => b - a > inner || a < pos ? a : b > pos + inner ? b - inner : pos;
		const [r1, r2] = sel.type === "column" ? [sel.active.r, sel.active.r] : [sel.range.r1, Math.min(sel.range.r2, ext.rows)];
		const [c1, c2] = sel.type === "row" ? [sel.active.c, sel.active.c] : [sel.range.c1, Math.min(sel.range.c2, ext.cols)];
		const top = fit(body.scrollTop, g.rowTop(r1), g.rowTop(r2 + 1), body.clientHeight - o.y);
		const left = fit(body.scrollLeft, g.colLeft(c1), g.colLeft(c2 + 1), body.clientWidth - o.x);
		if (top === body.scrollTop && left === body.scrollLeft) return;
		body.scrollTop = top;
		body.scrollLeft = left;
		host.draw();
	};
	/** 今のセル (r, c) を、選択の種類のまま選ぶ */
	const moveTo = (type, r, c) => {
		if (type === "column") return colSel(c, c, r);
		if (type === "cell") return cellSel(r, c, false);
		const row = cellSel(r, c, true);
		return {
			...row,
			active: {
				r: row.active.r,
				c
			}
		};
	};
	const onKey = (e) => {
		if (e.altKey || e.shiftKey || e.isComposing || !SELECTION_KEYS.has(e.key)) return;
		const t = e.target;
		if (t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName))) return;
		const g = host.geometry();
		const sh = host.sheet();
		if (!g || !sh) return;
		const body = host.body;
		const ctrl = e.ctrlKey || e.metaKey;
		let next;
		if (!value) next = moveTo(mode === "cell" ? "cell" : "row", g.rowAt(body.scrollTop), g.colAt(body.scrollLeft));
		else {
			const to = keyTarget(e.key, ctrl, value, {
				geo: g,
				extent: host.extent(),
				used: {
					rows: sh.rowCount,
					cols: sh.colCount
				},
				pageHeight: body.clientHeight - host.origin().y
			});
			if (to === void 0) return;
			e.preventDefault();
			if (to === null) return;
			next = moveTo(value.type, to.r, to.c);
		}
		e.preventDefault();
		update(next, "user");
		reveal(next);
	};
	/** 見出しの印を付け直す（描いた行番号・列記号のうち、選択に入るもの） */
	const markHeaders = (sel) => {
		marked.forEach((el) => el.classList.remove("is-selected"));
		marked = [];
		const table = host.table();
		if (!sel || !table) return;
		const { r1, r2, c1, c2 } = sel.range;
		if (sel.type !== "column") for (const tr of table.querySelectorAll("tbody tr[aria-rowindex]")) {
			const r = Number(tr.getAttribute("aria-rowindex")) - 1;
			const head = tr.firstElementChild;
			if (r >= r1 && r <= r2 && head?.classList.contains("xs-rownum")) marked.push(head);
		}
		if (sel.type !== "row") {
			const names = /* @__PURE__ */ new Set();
			for (let c = c1; c <= Math.min(c2, host.extent().cols); c++) names.add(host.colName(c));
			for (const th of table.querySelectorAll("thead th:not(.xs-rownum)")) if (names.has(th.textContent ?? "")) marked.push(th);
		}
		marked.forEach((el) => el.classList.add("is-selected"));
	};
	/** 画面の位置（スクロール枠の表示の左上からの px）にある物を選ぶ */
	const pick = (x, y) => {
		const g = host.geometry();
		const sh = host.sheet();
		if (!g || !sh) return null;
		const body = host.body;
		if (x < 0 || y < 0 || x >= body.clientWidth || y >= body.clientHeight) return null;
		const o = host.origin();
		const cx = body.scrollLeft + x - o.x;
		const cy = body.scrollTop + y - o.y;
		const inRows = cy >= 0 && cy < g.totalHeight;
		const inCols = cx >= 0 && cx < g.totalWidth;
		const prev = value?.active ?? {
			r: 1,
			c: 1
		};
		if (y < o.y && x >= o.x && inCols) return colSel(g.colAt(cx), g.colAt(cx), prev.r);
		if (x < o.x && y >= o.y && inRows) return rowSel(g.rowAt(cy), g.rowAt(cy), prev.c);
		if (x >= o.x && y >= o.y && inRows && inCols && mode) return cellSel(g.rowAt(cy), g.colAt(cx), mode === "row");
		return null;
	};
	let down = null;
	const onDown = (e) => {
		down = e.button === 0 ? {
			x: e.clientX,
			y: e.clientY
		} : null;
	};
	const onClick = (e) => {
		const start = down;
		down = null;
		if (!start || e.button !== 0) return;
		if (Math.hypot(e.clientX - start.x, e.clientY - start.y) > CLICK_SLOP) return;
		const text = host.body.ownerDocument.getSelection?.();
		if (text && !text.isCollapsed && host.body.contains(text.anchorNode)) return;
		const rect = host.body.getBoundingClientRect();
		const sel = pick(e.clientX - rect.left - host.body.clientLeft, e.clientY - rect.top - host.body.clientTop);
		if (sel) {
			update(sel, "user");
			host.body.focus({ preventScroll: true });
		}
	};
	const onCopy = (e) => {
		if (!value || !e.clipboardData) return;
		const text = host.body.ownerDocument.getSelection?.();
		if (text && !text.isCollapsed) return;
		const out = controller.text();
		e.preventDefault();
		if (out === null) host.copyTooLarge();
		else e.clipboardData.setData("text/plain", out);
	};
	if (mode) {
		host.body.addEventListener("copy", onCopy);
		host.body.addEventListener("mousedown", onDown);
		host.body.addEventListener("click", onClick);
		host.body.addEventListener("keydown", onKey);
		host.body.tabIndex = 0;
	}
	const controller = {
		get: () => value,
		set(input, origin = "api") {
			if (input === null) {
				update(null, origin);
				return;
			}
			const prev = value?.active ?? {
				r: 1,
				c: 1
			};
			let next;
			if (input.type === "row") {
				const r2 = input.r2 ?? input.r;
				if (!inRange(input.r, MAX_GRID_ROWS) || !inRange(r2, MAX_GRID_ROWS)) throw new RangeError("選択の行の番号が範囲外");
				next = rowSel(input.r, r2, prev.c);
			} else if (input.type === "column") {
				const c2 = input.c2 ?? input.c;
				if (!inRange(input.c, MAX_GRID_COLS) || !inRange(c2, MAX_GRID_COLS)) throw new RangeError("選択の列の番号が範囲外");
				next = colSel(input.c, c2, prev.r);
			} else if (input.type === "cell") {
				if (!inRange(input.r, MAX_GRID_ROWS) || !inRange(input.c, MAX_GRID_COLS)) throw new RangeError("選択のセルの番号が範囲外");
				next = cellSel(input.r, input.c, false);
			} else throw new RangeError("選択の種類が正しくない");
			update(next, origin);
		},
		layout() {
			const sel = value;
			const g = host.geometry();
			markHeaders(sel);
			if (!sel || !g) {
				band.hidden = true;
				return;
			}
			const ext = host.extent();
			const { r1, c1 } = sel.range;
			const r2 = Math.min(sel.range.r2, ext.rows);
			const c2 = Math.min(sel.range.c2, ext.cols);
			const o = host.origin();
			const top = g.rowTop(r1);
			const left = g.colLeft(c1);
			const h = r1 > ext.rows ? 0 : g.rowTop(r2 + 1) - top;
			const w = c1 > ext.cols ? 0 : g.colLeft(c2 + 1) - left;
			if (h <= 0 || w <= 0) {
				band.hidden = true;
				return;
			}
			if (band.parentNode !== host.wrapper) host.wrapper.insertBefore(band, host.before);
			band.className = `xs-sel is-${sel.type}`;
			band.style.left = `${o.x + left}px`;
			band.style.top = `${o.y + top}px`;
			band.style.width = `${w}px`;
			band.style.height = `${h}px`;
			band.hidden = false;
		},
		text() {
			const sh = host.sheet();
			if (!value || !sh) return null;
			const range = copyRange(value, {
				rows: sh.rowCount,
				cols: sh.colCount
			});
			return rangeText(sh, range, host.createGeometry(sh.layout, Math.max(1, range.r2), Math.max(1, range.c2)));
		},
		clear() {
			update(null, "api");
		},
		dispose() {
			host.body.removeEventListener("mousedown", onDown);
			host.body.removeEventListener("click", onClick);
			host.body.removeEventListener("keydown", onKey);
			host.body.removeEventListener("copy", onCopy);
			host.body.removeAttribute("tabindex");
			live.remove();
			marked = [];
			value = null;
			band.remove();
		}
	};
	return controller;
}
export { createSelection };
