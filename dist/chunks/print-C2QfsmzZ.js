// src/render/print.ts
/** 既定の 1 ページの中身の大きさ（A4 縦・Excel の「標準」の余白。CSS px） */
var PRINT_PAGE_DEFAULT = {
	width: 659,
	height: 978
};
/**
* 印刷の指定を解く。
* @throws RangeError 用紙の大きさが正の有限の数でない・`fit` が `'width'` でない・`maxCells` が 1 以上の数でない
*/
function resolvePrintOptions(options = {}) {
	const page = options.page ?? PRINT_PAGE_DEFAULT;
	const positive = (v) => typeof v === "number" && Number.isFinite(v) && v > 0;
	if (!positive(page.width) || !positive(page.height)) throw new RangeError("印刷の用紙の大きさが正しくない");
	if (options.fit !== void 0 && options.fit !== "width") throw new RangeError("印刷の fit は width だけ");
	const maxCells = options.maxCells ?? 2e5;
	if (typeof maxCells !== "number" || Number.isNaN(maxCells) || maxCells < 1) throw new RangeError("印刷の maxCells は 1 以上の数");
	return {
		page: {
			width: page.width,
			height: page.height
		},
		fit: options.fit === "width",
		headers: options.headers === true,
		gridlines: options.gridlines === true,
		maxCells: Math.floor(maxCells)
	};
}
/** 範囲を読む行数の単位（はみ出しの届く列を数える読み取り） */
var READ_CHUNK = 512;
/**
* 番号 start〜end を、長さ limit に収まる区切りに分ける（ページの切れ目）。切れ目は行・列の境。
* 1 つだけで limit を超える番号は、その番号だけの区切りにする。長さ 0 の番号（非表示）だけの区切りは返さない。
*
* @param edge 番号 i の始まりの位置（`rowTop`・`colLeft`）
* @param at 位置を含む番号（`rowAt`・`colAt`）
*/
function splitBands(start, end, edge, at, limit) {
	const out = [];
	let a = start;
	while (a <= end) {
		const top = edge(a);
		let b = Math.max(a, Math.min(end, at(top + limit)));
		while (b > a && edge(b + 1) - top > limit) b--;
		if (edge(b + 1) > top) out.push([a, b]);
		a = b + 1;
	}
	return out;
}
/**
* 印刷の範囲・倍率・ページの切れ目を決める（ADR-0010 §2-2・§2-5）。
* - 行は データ・図の最下行まで、列は データ・図の最右列・はみ出しの届く列（見える全部の行から）まで。
* - 見える行を上から数え、見える行 × 見える列が `maxCells` を超える手前で打ち切る（最初の 1 行は必ず描く）。
*/
function planPrint(kit, sheet, boxes, opts) {
	const { createGeometry, drawingExtent, spillReachCols, rowNumWidth, zoomGeometry } = kit;
	const { GRID_HEADER_HEIGHT, MAX_GRID_ROWS, MAX_GRID_COLS } = kit;
	const big = createGeometry(sheet.layout, MAX_GRID_ROWS, MAX_GRID_COLS);
	const fig = drawingExtent(big, boxes);
	const lastRow = Math.min(MAX_GRID_ROWS, Math.max(1, sheet.rowCount, fig.rows));
	let cols = Math.min(MAX_GRID_COLS, Math.max(1, sheet.colCount, fig.cols));
	let counted = 0;
	let shown = 0;
	const shownCols = (n) => {
		for (; counted < n; counted++) if (big.colWidth(counted + 1) > 0) shown++;
		return shown;
	};
	let seen = 0;
	let end = 0;
	let truncated = false;
	for (let a = 1; a <= lastRow && !truncated; a += READ_CHUNK) {
		const b = Math.min(lastRow, a + READ_CHUNK - 1);
		const data = a <= sheet.rowCount ? sheet.getRows(a, Math.min(b, sheet.rowCount)) : [];
		for (let r = a; r <= b; r++) {
			if (big.rowHeight(r) === 0) {
				end = r;
				continue;
			}
			const row = data[r - a];
			const reach = row ? Math.min(MAX_GRID_COLS, spillReachCols(big, [row], sheet.formats)) : 0;
			const need = Math.max(cols, reach);
			if (seen > 0 && (seen + 1) * shownCols(need) > opts.maxCells) {
				truncated = true;
				break;
			}
			seen++;
			cols = need;
			end = r;
		}
	}
	const headX = opts.headers ? rowNumWidth(end) : 0;
	const headY = opts.headers ? GRID_HEADER_HEIGHT : 0;
	let zoom = 1;
	if (opts.fit) {
		const content = big.colLeft(cols + 1);
		const avail = Math.floor(opts.page.width - headX);
		if (avail > 0 && content > avail) zoom = avail / content;
	}
	const g = zoomGeometry(createGeometry(sheet.layout, end, cols), zoom);
	return {
		rows: end,
		cols,
		zoom,
		truncated,
		rowBands: splitBands(1, end, (r) => g.rowTop(r), (y) => g.rowAt(y), opts.page.height - headY),
		colBands: splitBands(1, cols, (c) => g.colLeft(c), (x) => g.colAt(x), opts.page.width - headX)
	};
}
/**
* 寸法を 1 ページの窓にする：範囲の外の行・列は高さ・幅 0、位置は範囲の左上を 0 とする。
* 表・罫線の部品にそのまま渡すと、範囲の外の余白を作らずに範囲だけを描く。
*/
function windowGeometry(g, p) {
	const y0 = g.rowTop(p.r1);
	const x0 = g.colLeft(p.c1);
	const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));
	return {
		rowTop: (r) => r <= p.r1 ? 0 : g.rowTop(Math.min(r, p.r2 + 1)) - y0,
		rowHeight: (r) => r < p.r1 || r > p.r2 ? 0 : g.rowHeight(r),
		colLeft: (c) => c <= p.c1 ? 0 : g.colLeft(Math.min(c, p.c2 + 1)) - x0,
		colWidth: (c) => c < p.c1 || c > p.c2 ? 0 : g.colWidth(c),
		rowAt: (y) => clamp(g.rowAt(y + y0), p.r1, p.r2),
		colAt: (x) => clamp(g.colAt(x + x0), p.c1, p.c2),
		totalWidth: g.colLeft(p.c2 + 1) - x0,
		totalHeight: g.rowTop(p.r2 + 1) - y0
	};
}
/** 範囲の列に、値・書式・罫線のあるセルがあるか（何も描く物の無いページを出さないため） */
function hasCellContent(rows, c1, c2) {
	for (const row of rows) for (let c = c1; c <= c2; c++) {
		const cell = row.cells[c - 1];
		if (cell && (cell.text !== "" || cell.colSpan > 1 || cell.rowSpan > 1) || row.formats?.[c - 1] !== void 0 || row.borders?.[c - 1] !== void 0) return true;
	}
	return false;
}
/** 書式の CSS の連番（印刷の分）。表示の連番（正の数）と重ならないよう負の数にする */
var printSeq = 0;
/** 中断の例外（メッセージにセル値・シート名を含めない） */
function abortError() {
	return new DOMException("印刷の描画を中断しました", "AbortError");
}
/** 次のタスクまで譲る（長い描画の間にも中断を受けられるように） */
var yieldTask = () => new Promise((r) => setTimeout(r, 0));
/** 今の時刻（ms） */
var now = () => typeof performance !== "undefined" ? performance.now() : Date.now();
/** 描く間に何 ms ごとに譲るか */
var YIELD_MS = 16;
/**
* 表示中のシートを、印刷のためにページに分けて `target` の中に描く（ADR-0010 §2）。
* @throws RangeError 正しくない指定
* @throws DOMException AbortError（中断。足した物は片付けてから拒否する）
*/
async function renderPrint(kit, target, src, options = {}) {
	const { colName, createGeometry, createBorderLayer, buildBorderPaths, rectsIntersect } = kit;
	const { buildFormatCss, applyMarginFormats, collectVisibleRows, rowNumWidth } = kit;
	const { buildGridTable, createImageLayer, zoomGeometry } = kit;
	const { GRID_HEADER_HEIGHT, MAX_GRID_ROWS, MAX_GRID_COLS } = kit;
	const opts = resolvePrintOptions(options);
	const signal = options.signal;
	if (signal?.aborted) throw abortError();
	const empty = {
		pages: 0,
		truncated: false,
		destroy: () => void 0
	};
	if (!src) return empty;
	const sh = src.sheet;
	const plan = planPrint(kit, sh, src.boxes, opts);
	const z = plan.zoom;
	const g = zoomGeometry(createGeometry(sh.layout, plan.rows, plan.cols), z);
	const big = createGeometry(sh.layout, MAX_GRID_ROWS, MAX_GRID_COLS);
	const headX = opts.headers ? rowNumWidth(plan.rows) : 0;
	const headY = opts.headers ? GRID_HEADER_HEIGHT : 0;
	const root = document.createElement("div");
	root.className = "xs-sheet xs-print";
	if (opts.gridlines) root.classList.add("is-gridlines");
	root.style.setProperty("--xs-zoom", String(z));
	const seq = --printSeq;
	root.setAttribute("data-xs-fmt", String(seq));
	const fmtStyle = document.createElement("style");
	if (src.nonce !== void 0) fmtStyle.setAttribute("nonce", src.nonce);
	fmtStyle.textContent = buildFormatCss(sh.formats, seq);
	root.append(fmtStyle);
	const layers = [];
	let destroyed = false;
	const destroy = () => {
		if (destroyed) return;
		destroyed = true;
		layers.forEach((l) => l.dispose());
		layers.length = 0;
		root.replaceChildren();
		root.remove();
	};
	const check = () => {
		if (signal?.aborted) throw abortError();
	};
	/** 行 a〜b の読み出し（データより下は空の行） */
	const readRows = (a, b) => {
		const rows = a <= sh.rowCount ? sh.getRows(a, Math.min(b, sh.rowCount)) : [];
		for (let r = Math.max(a, sh.rowCount + 1); r <= b; r++) rows.push({
			r,
			cells: []
		});
		return rows;
	};
	/**
	* 1 ページの要素（表・罫線の層・図の層）
	* @param wg ページの窓の寸法（{@link windowGeometry}）
	* @param win 図の範囲（100% の px。A1 の左上基準）
	* @param figures ページに図があるか
	*/
	const buildPage = (rows, wg, p, win, figures) => {
		const page = document.createElement("div");
		page.className = "xs-print-page";
		page.style.width = `${opts.page.width}px`;
		page.style.height = `${opts.page.height}px`;
		const wrap = document.createElement("div");
		wrap.className = "xs-wrap";
		wrap.append(buildGridTable(colName, p.c2, rows, wg, plan.rows, sh.layout.mdw * z, {
			first: p.c1,
			last: p.c2
		}, sh.formats, null, opts.headers));
		if (sh.borders.length > 0) {
			const borders = createBorderLayer();
			borders.render(buildBorderPaths(rows, sh.borders, wg, {
				first: p.c1,
				last: p.c2
			}), headX, headY);
			wrap.append(borders.element);
		}
		if (figures) {
			const clip = document.createElement("div");
			clip.className = "xs-print-clip";
			clip.style.left = `${headX}px`;
			clip.style.top = `${headY}px`;
			clip.style.width = `${wg.totalWidth}px`;
			clip.style.height = `${wg.totalHeight}px`;
			const layer = createImageLayer({ messages: src.messages });
			layers.push(layer);
			layer.setSheet({
				drawings: src.drawings,
				boxes: src.boxes,
				readImage: (path) => src.book.readImage(src.index, path),
				readChart: (path) => src.book.readChart(src.index, path)
			});
			layer.place(-win.x * z, -win.y * z, win.x + win.w, win.y + win.h);
			layer.element.style.transform = z === 1 ? "" : `scale(${z})`;
			layer.update(win);
			clip.append(layer.element);
			wrap.append(clip);
		}
		page.append(wrap);
		return page;
	};
	try {
		let pages = 0;
		let last = now();
		for (const [c1, c2] of plan.colBands) for (const [r1, r2] of plan.rowBands) {
			const p = {
				r1,
				r2,
				c1,
				c2
			};
			const wg = windowGeometry(g, p);
			const read = collectVisibleRows(wg, r1, r2, r2, readRows);
			const rows = sh.formats.length > 0 ? applyMarginFormats(read, c1, c2, sh.rowCount, sh.colCount, (r) => sh.rowFormat(r), (c) => sh.colFormat(c)) : read;
			const win = {
				x: big.colLeft(c1),
				y: big.rowTop(r1),
				w: big.colLeft(c2 + 1) - big.colLeft(c1),
				h: big.rowTop(r2 + 1) - big.rowTop(r1)
			};
			const figures = src.boxes.some((b) => rectsIntersect(b.rect, win));
			if (!figures && !hasCellContent(rows, c1, c2)) continue;
			root.append(buildPage(rows, wg, p, win, figures));
			pages++;
			if (now() - last > YIELD_MS) {
				await yieldTask();
				check();
				last = now();
			}
		}
		check();
		if (pages === 0) {
			destroy();
			return {
				...empty,
				truncated: plan.truncated
			};
		}
		target.append(root);
		await waitAbortable(settleAll(layers, root), signal);
		check();
		return {
			pages,
			truncated: plan.truncated,
			destroy
		};
	} catch (e) {
		destroy();
		throw e;
	}
}
/** 図の層の画像・グラフの読み込みを待ち、画像のデコードを待つ（壊れた画像は破線枠に替わる） */
async function settleAll(layers, root) {
	await Promise.all(layers.map((l) => l.settled()));
	const images = [...root.querySelectorAll("img")];
	await Promise.all(images.map((img) => typeof img.decode === "function" ? img.decode().catch(() => void 0) : void 0));
}
/** p を待つ。先に中断されたら AbortError で拒否する */
function waitAbortable(p, signal) {
	if (!signal) return p;
	return new Promise((resolve, reject) => {
		const onAbort = () => reject(abortError());
		if (signal.aborted) {
			onAbort();
			return;
		}
		signal.addEventListener("abort", onAbort, { once: true });
		p.then((v) => {
			signal.removeEventListener("abort", onAbort);
			resolve(v);
		}, (e) => {
			signal.removeEventListener("abort", onAbort);
			reject(e instanceof Error ? e : new Error(String(e)));
		});
	});
}
export { renderPrint };
