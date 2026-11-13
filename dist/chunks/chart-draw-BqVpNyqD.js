import { a as applyTransforms, l as DEFAULT_THEME, n as formatNumberByCode } from "./format-CYjzlX0B.js";
import { r as dashArray } from "./render-iJM1Tf4e.js";
// src/render/chart-axis.ts
/** 軸の暫定の係数（C1・C2・C7。ISSUE-104 の照合で直す） */
var AXIS_CONSTANTS = {
	/** C1：dmin がこの割合 × dmax 未満なら自動の最小を 0 にする（5/6） */
	zeroBaselineRatio: 5 / 6,
	/** C1：自動の範囲の端に足す余り（範囲に対する割合・5%） */
	marginRatio: .05,
	/**
	* C1：目盛の数（主の刻みの数）の上限 N ＝ プロットの長さ ÷ 目盛 1 つぶんの最小の間隔。
	* 間隔は縦軸＝文字の高さ × この値（Excel 16 の照合：10.5〜11.0pt の範囲＝9pt の文字で 1.0（手動の内側 110pt で 10 目盛が出る））。
	*/
	labelPitchFactor: 1,
	/**
	* C1：横向きの軸（横棒の数値軸）の最小の間隔＝文字の高さ × この値。最長の目盛の文字の幅＋horizontalLabelGapPx がこれより大きければそちら。
	* 3 回目の照合（ISSUE-110・幅の掃引 96 個）：目盛の数の上限が 10 に届くのは、面の幅 360pt（自動配置の外側 459px）と 380pt（485px）の間
	* ＝ 9pt の文字（文字の高さ 14.4px）で最小の間隔 46.8px（3.25）。目盛の数を決める長さは、プロットの内側でなく自動配置の外側の幅
	* （layout の solveAxes の scaleW。Excel は項目軸の文字・目盛の文字の余白を引く前の幅で刻みを決める）。
	* 旧：内側の幅で 2.92（31.6pt）。外側の幅で 3.25 にすると、旧 KNOWN の C1（積み上げの負・7 桁・100% 積み上げ・長い項目名）が合う
	*/
	horizontalPitchFactor: 3.25,
	/**
	* C1：横向きの軸で、最長の目盛の文字の幅に足す隙間（px）。隣の文字との隙間を 0 にしないための値。
	* 文字の幅の推定（数字 0.55em）は游ゴシック 9pt の実際より約 10% 広いので、項目軸の {@link labelPaddingPx}（4）は足さず 2 にした
	* （照合：barClustered_big_small は 6 桁の目盛で Excel の間隔が 42.6px 以下。4 を足すと 43.6px で目盛が 1 つ減って外れる）。
	*/
	horizontalLabelGapPx: 2,
	/** C1：目盛の数の上限 N の下限と上限（2〜10。Excel 16 の照合：端を丸める前の刻みの数が N 以下の最小の刻み） */
	minTickLimit: 2,
	maxTickLimit: 10,
	/** C1：補助の刻み ＝ 主の刻み ÷ この値 */
	minorDivisions: 5,
	/** C2：100% 積み上げの書式（刻みは他の軸と同じく自動で決める。範囲は 0〜1、負があれば -1〜1） */
	percentFormat: "0%",
	/** 目盛の数の絶対の上限（主・補助。明示の刻みが細かすぎる時の安全弁） */
	maxMajorTicks: 100,
	maxMinorTicks: 500,
	/** C7：項目の文字の間に空ける余白（px） */
	labelPaddingPx: 4
};
/** 浮動小数の誤差を丸める（12 桁） */
function round12(x) {
	const r = Number(x.toPrecision(12));
	return r === 0 ? 0 : r;
}
/** 有限の数だけ通す */
function finiteOr(x) {
	return typeof x === "number" && Number.isFinite(x) ? x : void 0;
}
/**
* 系列からデータの範囲を求める。点が 1 つも無ければ null。
* - `clustered`・`standard`：全点の最小・最大。
* - `stacked`：項目ごとの正の和の最大と負の和の最小（0 を含む）。
* - `percentStacked`：最小は負の値が 1 つでもあれば -1・無ければ 0。最大は 1（全部が負なら 0。C2。Excel 16 の照合）。点があれば。
*/
function dataExtent(grouping, series) {
	let any = false;
	const count = series.reduce((m, s) => Math.max(m, s.values.length), 0);
	if (grouping === "stacked" || grouping === "percentStacked") {
		let hi = 0;
		let lo = 0;
		let hasPos = false;
		const percent = grouping === "percentStacked";
		for (let i = 0; i < count; i++) {
			let pos = 0;
			let neg = 0;
			for (const s of series) {
				const v = s.values[i];
				if (typeof v !== "number" || !Number.isFinite(v)) continue;
				any = true;
				if (v >= 0) pos += v;
				else neg += v;
			}
			if (percent) {
				if (neg < 0) lo = -1;
				if (pos > 0) hasPos = true;
			} else {
				if (pos > hi) hi = pos;
				if (neg < lo) lo = neg;
			}
		}
		if (!any) return null;
		return percent ? {
			min: lo,
			max: lo < 0 && !hasPos ? 0 : 1
		} : {
			min: lo,
			max: hi
		};
	}
	let min = Infinity;
	let max = -Infinity;
	for (const s of series) for (const v of s.values) {
		if (typeof v !== "number" || !Number.isFinite(v)) continue;
		any = true;
		if (v < min) min = v;
		if (v > max) max = v;
	}
	return any ? {
		min,
		max
	} : null;
}
/**
* {1, 2, 5}×10ⁿ のうち、`span ÷ u ≤ maxTicks` を満たす最小の刻みを返す（P6）。
* `span` が 0 以下・非有限なら 1。`maxTicks` は 1 以上に丸める。ループは高々数回。
*/
function niceUnit(span, maxTicks) {
	const s = Number.isFinite(span) && span > 0 ? span : 1;
	const n = Math.max(1, Math.floor(maxTicks));
	let exp = Math.floor(Math.log10(s / n)) - 1;
	for (let guard = 0; guard < 12; guard++) {
		for (const m of [
			1,
			2,
			5
		]) {
			const u = m * 10 ** exp;
			if (s / u <= n + 1e-9) return round12(u);
		}
		exp++;
	}
	return round12(10 ** exp);
}
/**
* 目盛の数（主の刻みの数）の上限 N（C1）：プロットの長さ ÷ 目盛 1 つぶんの最小の間隔 を 2〜10 に丸める。
* 縦軸の間隔は 文字の高さ × {@link AXIS_CONSTANTS.labelPitchFactor}。
* 横向きの軸（横棒の数値軸）は `labelWidthPx`（最長の目盛の文字の幅）を渡すと、（幅 + {@link AXIS_CONSTANTS.horizontalLabelGapPx}）と
* 文字の高さ × {@link AXIS_CONSTANTS.horizontalPitchFactor} の大きい方（隣の文字との隙間が 0 にならないようにする）。
*/
function tickCountLimit(lengthPx, labelHeightPx, labelWidthPx) {
	const c = AXIS_CONSTANTS;
	const h = Math.max(1e-6, labelHeightPx);
	const pitch = labelWidthPx !== void 0 && Number.isFinite(labelWidthPx) ? Math.max(h * c.horizontalPitchFactor, labelWidthPx + c.horizontalLabelGapPx) : h * c.labelPitchFactor;
	const raw = Number.isFinite(lengthPx) ? Math.floor(lengthPx / pitch) : c.minTickLimit;
	return Math.min(c.maxTickLimit, Math.max(c.minTickLimit, raw));
}
/** `min` から `max` まで `unit` おきの値（個数に上限。誤差は丸める） */
function ticksBetween(min, max, unit, cap) {
	const out = [];
	const n = Math.min(cap, Math.floor((max - min) / unit + 1e-9));
	for (let i = 0; i <= n; i++) out.push(round12(min + i * unit));
	return out;
}
/**
* 数値軸の範囲・刻み・目盛を決める（ADR-0031 §決定 5）。
* 明示（`c:min`・`c:max`・`c:majorUnit`・`c:minorUnit`・`c:orientation`）があれば従い、無い値だけ自動で決める。
* - 自動の最小：dmin ≥ 0 で dmin < dmax × 5/6 なら 0。dmax ≤ 0 側も対称に 0。それ以外は範囲の 5% の余りを取って刻みに切り下げる。
* - 自動の最大：dmax に 5% の余りを足して刻みの倍数に切り上げる。
* - 刻み：{@link niceUnit}（目盛の数の上限 {@link tickCountLimit}）。補助は 1/5。
* - 点が無ければ 0〜1。`percentStacked` は 0〜1（負を含めば -1〜1）で、刻みは他の軸と同じく自動（{@link tickCountLimit}。明示があれば優先）。
* - 目盛の数は {@link AXIS_CONSTANTS} の `maxMajorTicks`・`maxMinorTicks` で抑える。
*/
function computeValueScale(input) {
	const c = AXIS_CONSTANTS;
	const axis = input.axis;
	const percent = input.percentStacked === true;
	const reversed = axis?.orientation === "maxMin";
	/** 点が無い・全部 0：0〜1 に固定（自動の余りを足さない） */
	const fixed = percent || input.extent === null || input.extent.min === 0 && input.extent.max === 0;
	const limitN = tickCountLimit(input.lengthPx, input.labelHeightPx, input.labelWidthPx);
	const exMin = finiteOr(axis?.min);
	const exMax = finiteOr(axis?.max);
	let dmin = input.extent?.min ?? 0;
	let dmax = input.extent?.max ?? (input.extent === null ? 1 : 0);
	if (dmin === dmax) {
		if (dmax > 0) dmin = 0;
		else if (dmax < 0) dmax = 0;
		else dmax = 1;
	}
	const margin = (dmax - dmin) * c.marginRatio;
	let lo;
	let hi;
	let loAuto = false;
	let hiAuto = false;
	if (fixed) {
		lo = 0;
		hi = 1;
	} else {
		lo = dmin;
		hi = dmax;
	}
	if (exMin !== void 0) lo = exMin;
	else if (percent && input.extent !== null && input.extent.min < 0) {
		lo = input.extent.min;
		loAuto = true;
	} else if (!fixed) {
		if (dmin >= 0 && dmin < dmax * c.zeroBaselineRatio) lo = 0;
		else {
			lo = dmin - margin;
			loAuto = true;
		}
	}
	if (exMax !== void 0) hi = exMax;
	else if (percent && input.extent !== null) hi = input.extent.max;
	else if (!fixed) {
		if (dmax <= 0 && dmax > dmin * c.zeroBaselineRatio) hi = 0;
		else {
			hi = dmax + margin;
			hiAuto = true;
		}
	}
	if (!(hi > lo)) {
		if (exMin !== void 0 && exMax === void 0) hi = lo + 1;
		else if (exMax !== void 0 && exMin === void 0) lo = hi - 1;
		else {
			lo = exMin ?? 0;
			hi = lo + 1;
		}
		loAuto = false;
		hiAuto = false;
	}
	const span = hi - lo;
	let major;
	const exMajor = finiteOr(axis?.majorUnit);
	if (exMajor !== void 0 && exMajor > 0) major = exMajor;
	else major = niceUnit(span, limitN);
	if (span / major > c.maxMajorTicks) major = niceUnit(span, c.maxMajorTicks);
	if (loAuto) lo = round12(Math.floor(lo / major + 1e-9) * major);
	if (hiAuto) hi = round12(Math.ceil(hi / major - 1e-9) * major);
	if (!(hi > lo)) hi = round12(lo + major);
	let minor;
	const exMinor = finiteOr(axis?.minorUnit);
	if (exMinor !== void 0 && exMinor > 0) minor = exMinor;
	else minor = round12(major / c.minorDivisions);
	if (minor !== null && (hi - lo) / minor > c.maxMinorTicks) minor = null;
	const majorTicks = ticksBetween(lo, hi, major, c.maxMajorTicks);
	let minorTicks = [];
	if (minor !== null) {
		const majorSet = new Set(majorTicks);
		minorTicks = ticksBetween(lo, hi, minor, c.maxMinorTicks).filter((t) => !majorSet.has(t));
	}
	return {
		min: round12(lo),
		max: round12(hi),
		majorUnit: major,
		minorUnit: minor,
		reversed,
		majorTicks,
		minorTicks,
		percent
	};
}
/**
* 値を軸上の位置（0〜1・既定は範囲外を端に収める）にする。反転（`maxMin`）なら 1 − 位置。
* 0 は軸の始点側（縦軸なら下・横軸なら左）、反転すると逆の端。
* `clamp` を false にすると範囲外を収めず、0 未満・1 超の位置をそのまま返す（折れ線を幾何的に切るため。非有限の値は 0）。
*/
function valueFraction(scale, v, clamp = true) {
	const span = scale.max - scale.min;
	const p = span > 0 && Number.isFinite(v) ? (v - scale.min) / span : 0;
	const q = clamp ? p < 0 ? 0 : p > 1 ? 1 : p : p;
	return scale.reversed ? 1 - q : q;
}
/**
* 項目軸が数値軸と交わる値（`c:crosses`・`c:crossesAt` は項目軸の要素に書かれる＝数値軸上の位置。暫定）。
* `crossesAt` があればそれ（範囲に収める）、`min`・`max` は端、`autoZero`（既定）は 0 が範囲内なら 0・外なら近い端。
*/
function crossValue(scale, catAxis) {
	const at = finiteOr(catAxis?.crossesAt);
	const clamp = (v) => Math.min(scale.max, Math.max(scale.min, v));
	if (at !== void 0) return clamp(at);
	if (catAxis?.crosses === "min") return scale.min;
	if (catAxis?.crosses === "max") return scale.max;
	return clamp(0);
}
/**
* 数値軸の目盛の文字の書式コード（C2）：`percentStacked` で明示が無ければ `0%`。
* `c:numFmt` が明示で `sourceLinked` でなければその書式。`sourceLinked` か無指定なら最初の系列の書式（無ければ General）。
*/
function tickFormatCode(axis, firstSeriesFormat, percent) {
	if (axis?.numFmt !== void 0 && !axis.numFmt.sourceLinked) return axis.numFmt.code;
	if (percent) return AXIS_CONSTANTS.percentFormat;
	return firstSeriesFormat ?? "General";
}
/** 目盛の値を書式コードで文字にする（ssf。書式が壊れていても General で出す）。-0 は 0 */
function tickLabel(code, value) {
	return formatNumberByCode(code, value === 0 ? 0 : value, false);
}
/**
* 項目軸の位置（軸の始点からの距離 px。項目ごとの中心）。
* - 既定は帯の等分：中心は (i + 0.5) × 帯の幅。
* - `midCat`（折れ線の点が目盛の上）：項目が 2 つ以上なら 0 〜 長さを (n−1) 等分、1 つなら中央。
* - `reversed` なら始点と終点を入れ替える（`c:orientation maxMin`・横棒の既定の向きの補正）。
*/
function categoryPositions(count, lengthPx, opts = {}) {
	const n = Math.max(0, Math.floor(count));
	const out = [];
	const mid = opts.midCat === true && n >= 2;
	for (let i = 0; i < n; i++) {
		const p = mid ? i * lengthPx / (n - 1) : (i + .5) * lengthPx / n;
		out.push(opts.reversed === true ? lengthPx - p : p);
	}
	return out;
}
/** 項目の間隔（px）：`midCat` なら点どうしの間隔、そうでなければ帯の幅 */
function categorySpacing(count, lengthPx, midCat = false) {
	const n = Math.max(1, Math.floor(count));
	return midCat && n >= 2 ? lengthPx / (n - 1) : lengthPx / n;
}
/**
* 項目の文字を何個おきに出すか（C7・P7）。
* - `c:tickLblSkip` の明示（1 以上の整数）があればそれ。
* - 無ければ、最長の文字の大きさ（軸の向きに沿った幅 px。横軸なら文字の幅・縦軸なら文字の高さ）＋余白が間隔に入る最小の個数。
*   余白は既定で {@link AXIS_CONSTANTS.labelPaddingPx}。縦に回した文字の間引き（ISSUE-107）は `paddingPx` で小さい余白を渡す。
*/
function categoryLabelSkip(count, spacingPx, maxLabelExtentPx, explicit, paddingPx = AXIS_CONSTANTS.labelPaddingPx) {
	const n = Math.max(1, Math.floor(count));
	const ex = finiteOr(explicit);
	if (ex !== void 0 && ex >= 1) return Math.min(n, Math.floor(ex));
	if (!(spacingPx > 0)) return n;
	const need = maxLabelExtentPx + paddingPx;
	const k = Math.ceil(need / spacingPx - 1e-9);
	return Math.min(n, Math.max(1, k));
}
/** 文字を出す項目の位置（0, skip, 2·skip, …） */
function categoryLabelIndices(count, skip) {
	const n = Math.max(0, Math.floor(count));
	const k = Math.max(1, Math.floor(skip));
	const out = [];
	for (let i = 0; i < n; i += k) out.push(i);
	return out;
}
// src/render/chart-thin.ts
/** 間引きの定数（単位は CSS px・`smoothSteps` だけ分割数） */
var CHART_THIN = {
	/** 項目の間隔がこれ未満のグラフを「密」とみなす */
	denseSpacingPx: 1,
	/** M4 の列の幅 */
	gridPx: .5,
	/** Douglas–Peucker の許容 */
	epsPx: .25,
	/** その他のマーカーで、直前に残した点からこの距離以内なら省く */
	markerSpacingPx: 1,
	/** 棒をまとめる列の幅 */
	barColumnPx: 1,
	/** 集合の棒で、系列 1 つにつき列の幅へ足す量（列の幅＝系列の数×これを整数 px へ丸めた値。1px 以上・下の上限以下） */
	barColumnPerSeriesPx: .2,
	/** 集合の棒の列の幅の上限（これより広げると列が縞に見えて元の絵から離れる。ISSUE-111 の計測） */
	barColumnMaxPx: 2,
	/** 滑らかな線を折れ線に直す時の区間ごとの分割数 */
	smoothSteps: 4
};
/**
* M4 の間引き：項目軸の方向に `gridPx` の幅で並ぶ列ごとに、同じ列に続く点の最初・最小・最大・最後だけを元の順で残す。
* 続く点を列とするので、項目の並びが逆順（右から左・上から下）でも動く。
* @param pts 点列
* @param catAxis 項目軸の方向（0＝x・1＝y）。最小・最大は他方の座標で決める
* @param gridPx 列の幅
* @returns 点が 3 未満なら入力の写し。それ以外は間引いた点列（最初と最後の点は必ず残る）
*/
function thinM4(pts, catAxis, gridPx = CHART_THIN.gridPx) {
	const n = pts.length;
	if (n < 3) return pts.slice();
	const v = catAxis === 0 ? 1 : 0;
	const out = [];
	let i = 0;
	while (i < n) {
		const col = Math.floor(pts[i][catAxis] / gridPx);
		let j = i + 1;
		while (j < n && Math.floor(pts[j][catAxis] / gridPx) === col) j++;
		if (j - i <= 4) for (let k = i; k < j; k++) out.push(pts[k]);
		else {
			let lo = i;
			let hi = i;
			for (let k = i + 1; k < j; k++) {
				if (pts[k][v] < pts[lo][v]) lo = k;
				if (pts[k][v] > pts[hi][v]) hi = k;
			}
			const keep = [.../* @__PURE__ */ new Set([
				i,
				lo,
				hi,
				j - 1
			])].sort((a, b) => a - b);
			for (const k of keep) out.push(pts[k]);
		}
		i = j;
	}
	return out;
}
/** 点 p から線分 ab までの距離 */
function distToSegment(p, a, b) {
	const dx = b[0] - a[0];
	const dy = b[1] - a[1];
	const len2 = dx * dx + dy * dy;
	let t = len2 === 0 ? 0 : ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / len2;
	t = t < 0 ? 0 : t > 1 ? 1 : t;
	return Math.hypot(p[0] - (a[0] + t * dx), p[1] - (a[1] + t * dy));
}
/** Douglas–Peucker の距離の計算回数の上限（これを超えたら間引きをやめる。通常の入力では届かない） */
var DP_WORK_BUDGET = 2e7;
/**
* Douglas–Peucker で残す点の添字（昇順）。反復で解く（再帰は深くなるため使わない）。端点は必ず残す。
* 残さなかった点は、残した点を結ぶ線分から `eps` 以内にある。
* 最悪の入力（鋸歯など）では 2 乗の時間になるため、距離の計算が {@link DP_WORK_BUDGET} 回を超えたら、
* 残りの区間は全点を残す（点が増えるだけで、ずれの上限は守られる）。
*/
function simplifyIndices(pts, eps = CHART_THIN.epsPx) {
	const n = pts.length;
	if (n < 3) return pts.map((_, i) => i);
	const keep = new Uint8Array(n);
	keep[0] = 1;
	keep[n - 1] = 1;
	const stack = [[0, n - 1]];
	let work = 0;
	while (stack.length > 0) {
		const [lo, hi] = stack.pop();
		if (hi - lo < 2) continue;
		work += hi - lo;
		if (work > DP_WORK_BUDGET) {
			for (let k = lo; k <= hi; k++) keep[k] = 1;
			continue;
		}
		const a = pts[lo];
		const b = pts[hi];
		let worst = -1;
		let at = -1;
		for (let k = lo + 1; k < hi; k++) {
			const d = distToSegment(pts[k], a, b);
			if (d > worst) {
				worst = d;
				at = k;
			}
		}
		if (worst > eps) {
			keep[at] = 1;
			stack.push([lo, at], [at, hi]);
		}
	}
	const out = [];
	for (let i = 0; i < n; i++) if (keep[i] === 1) out.push(i);
	return out;
}
/** Douglas–Peucker（端点を残す・許容 `eps`）で単純化した点列。点が 3 未満なら入力の写し */
function simplifyDP(pts, eps = CHART_THIN.epsPx) {
	if (pts.length < 3) return pts.slice();
	return simplifyIndices(pts, eps).map((i) => pts[i]);
}
/**
* 折れ線の間引き：M4 → Douglas–Peucker。中心線のずれは 0.25px 以下（M4 の列の中は正確）。
* @param pts 線の 1 つの部分（空の点・プロットの外で切れた後の 1 本）の点列
* @param catAxis 項目軸の方向（0＝x・1＝y）
*/
function thinPolyline(pts, catAxis) {
	return simplifyDP(thinM4(pts, catAxis));
}
/**
* 円のマーカーの帯の区間分け。続く中心の間隔が `maxGap` 以下の区間（2 点以上）を「帯」にする。
* 帯は呼び出し側が間引いた中心線を、太さ 2R・端と結びが round の線 1 本で描く。区間に入らない点は単独のマーカー。
* @param centers マーカーの中心（元の順）
* @param maxGap 続く中心の間隔の上限。{@link bandGap} で求める
* @returns `runs`＝帯の区間（各 2 点以上・元の順）／`singles`＝帯に入らない点（元の順）
*/
function splitMarkerBands(centers, maxGap) {
	const runs = [];
	const singles = [];
	let cur = [];
	const flush = () => {
		if (cur.length >= 2) runs.push(cur);
		else if (cur.length === 1) singles.push(cur[0]);
		cur = [];
	};
	for (const p of centers) {
		const last = cur[cur.length - 1];
		if (last !== void 0 && Math.hypot(p[0] - last[0], p[1] - last[1]) > maxGap) flush();
		cur.push(p);
	}
	flush();
	return {
		runs,
		singles
	};
}
/**
* 円のマーカー（外側の半径 R）の帯にできる、続く中心の間隔の上限 √(8R·eps)。
* 円の連なりの縁の波 d²/8R がこの間隔で eps 以下になる。
*/
function bandGap(R, eps = CHART_THIN.epsPx) {
	return Math.sqrt(8 * Math.max(0, R) * eps);
}
/**
* その他のマーカーの間引きで残す点の添字（昇順）。Douglas–Peucker（許容 `eps`）で選んだ点は残し、
* その間は直前に残した点から `spacing` 以内の点を省く（省いた点は残した点から `spacing` 以内）。
*/
function thinMarkerIndices(centers, eps = CHART_THIN.epsPx, spacing = CHART_THIN.markerSpacingPx) {
	const n = centers.length;
	if (n < 3) return centers.map((_, i) => i);
	const dp = new Uint8Array(n);
	for (const i of simplifyIndices(centers, eps)) dp[i] = 1;
	const out = [0];
	let last = centers[0];
	for (let i = 1; i < n; i++) {
		const p = centers[i];
		if (dp[i] === 1 || Math.hypot(p[0] - last[0], p[1] - last[1]) > spacing) {
			out.push(i);
			last = p;
		}
	}
	return out;
}
/**
* 棒の列ごとのまとめ。同じ系列・同じ塗りの棒を、項目軸の方向の `colPx` の列（棒の中心で決める）ごとに 1 つへまとめる。
* - 幅は合計（塗りの面積＝被覆を保つ）。位置は幅の重みの中心で、列に収まる時は列の中に収める。
* - 値の方向は和集合（最小〜最大）。正の側と負の側は別にまとめる。
* @returns 最初に現れた順のまとめた箱
*/
function mergeBarColumns(bars, colPx = CHART_THIN.barColumnPx) {
	const acc = /* @__PURE__ */ new Map();
	for (const bar of bars) {
		const w = bar.hi - bar.lo;
		const mid = (bar.lo + bar.hi) / 2;
		const col = Math.floor(mid / colPx);
		const key = col * 2 + (bar.neg ? 1 : 0);
		const cur = acc.get(key);
		if (cur === void 0) acc.set(key, {
			w,
			moment: w * mid,
			a: bar.a,
			b: bar.b,
			col
		});
		else {
			cur.w += w;
			cur.moment += w * mid;
			cur.a = Math.min(cur.a, bar.a);
			cur.b = Math.max(cur.b, bar.b);
		}
	}
	const out = [];
	for (const g of acc.values()) {
		let lo = (g.w > 0 ? g.moment / g.w : (g.col + .5) * colPx) - g.w / 2;
		if (g.w < colPx) {
			const colLo = g.col * colPx;
			lo = Math.min(Math.max(lo, colLo), colLo + colPx - g.w);
		}
		out.push({
			lo,
			hi: lo + g.w,
			a: g.a,
			b: g.b
		});
	}
	return out;
}
/**
* 集合の棒（重なり < 100）の列の幅。系列が多いほど同じ画素に重なる矩形が増えるので、
* 系列の数 × `barColumnPerSeriesPx` を整数 px に丸めた幅へ広げて矩形の数を減らす（1px 以上・`barColumnMaxPx` 以下）。
* 被覆（幅の合計）は保たれ、値の方向の包絡の細かさだけが 1〜2px になる（ISSUE-111・案 D）。
* 整数 px にするのは、画素の格子とずれる列（2.5px など）が縞になって元の絵から離れるため（実測：1px→2px は差が小さく、2.5〜4px は大きい）。
* @param seriesCount 系列の数
*/
function clusteredBarColumnPx(seriesCount) {
	const w = Math.round(Math.max(0, seriesCount) * CHART_THIN.barColumnPerSeriesPx);
	return Math.min(CHART_THIN.barColumnMaxPx, Math.max(CHART_THIN.barColumnPx, w));
}
/**
* 円・ドーナツの扇の連なりの束ね方（ISSUE-111・案 A）。外周の弧の長さ `ro × (a1 − a0)` が線の太さ `linePx` 以下で、
* `fixed` でない扇が 2 つ以上続く区間を返す。扇は順に「塗り→線」で描くため、そうした扇の中は両側の境の線で
* 隅々まで覆われ（線は不透明・実線の時）、区間の中身は線の色になる。呼び出し側が区間を線の色で塗った扇 1 つ
* （同じ線付き）に置き換えれば、区間の両端の境も残り、見た目は変わらない。
* @param spans 描く扇（元の順・隣り合う扇の角度は続いている）
* @param ro 扇の外の半径
* @param linePx 境の線の太さ（px）
* @returns 束ねる区間（`spans` の添字の両端を含む範囲・昇順・重ならない・各 2 つ以上）
*/
function thinSectorRuns(spans, ro, linePx) {
	const runs = [];
	if (!(linePx > 0) || !(ro > 0)) return runs;
	const thinOne = (s) => !s.fixed && ro * (s.a1 - s.a0) <= linePx + 1e-9;
	let i = 0;
	while (i < spans.length) {
		if (!thinOne(spans[i])) {
			i++;
			continue;
		}
		let j = i + 1;
		while (j < spans.length && thinOne(spans[j])) j++;
		if (j - i >= 2) runs.push([i, j - 1]);
		i = j;
	}
	return runs;
}
/** pt → px（96/72） */
var PX_PER_PT = 96 / 72;
/**
* 部品の既定（`c:spPr`・`c:txPr`・`c:style` が無い時。C4・C8）。
* ISSUE-104：`c:spPr` 等の無い合成のグラフを Excel 16 で開いた値に合わせた（面の枠線 0.5pt・#898989、文字 10pt・黒、タイトルは太字、折れ線 1.5pt、マーカー 6pt）。
* Excel が保存したグラフは `c:spPr`・`c:txPr` を持つのでこの表は使われない。
*/
var CHART_DEFAULTS = {
	/** グラフの面：塗り `bg1`・枠線 #898989（黒＋輝度 53.725%）・0.5pt */
	chartAreaFill: { slot: "lt1" },
	chartAreaLine: {
		slot: "dk1",
		lumOff: 53725
	},
	chartAreaLinePt: .5,
	/**
	* 主の目盛線・補助の目盛線・項目軸の線・数値軸の線（`c:spPr` の無い軸。ISSUE-105 の PNG の画素の照合）。
	* Excel 16 は主の目盛線と両方の軸の線を黒（`dk1`）・補助の目盛線を #BCBCBC（黒＋輝度 73.725%）の 1 画素の線で描く（COM の `Format.Line` の黒 0.5pt と合う）。
	* 旧値（目盛線 #d9d9d9・項目軸 #bfbfbf・数値軸なし）は、Excel が COM で作って保存した `c14:style` 付きのグラフ（#898989）とも合わなかった。
	*/
	gridMajor: { slot: "dk1" },
	gridMinor: {
		slot: "dk1",
		lumOff: 73725
	},
	catAxisLine: { slot: "dk1" },
	valAxisLine: { slot: "dk1" },
	/** 上の線の太さ（pt。COM の Weight。Excel は 1 画素未満の線を 1 画素に描く） */
	axisLinePt: .5,
	/**
	* `c14:style` を持つグラフ（Excel が COM の旧 API などで作った `c:spPr` 無しのグラフ。ISSUE-109・ISSUE-105 の COM）：
	* 主の目盛線・項目軸・数値軸の線は #898989（黒＋輝度 53.725%）0.5pt（太さは {@link axisLinePt}）、タイトルは 18pt 太字。
	* 補助の目盛線・軸と凡例の文字・面・折れ線は `c14:style` 無しと同じ。暗い背景（`c:style` 41〜48）とは独立に効く
	*/
	c14Line: {
		slot: "dk1",
		lumOff: 53725
	},
	c14TitlePt: 18,
	/**
	* `c:style` 41〜48（暗い背景。ISSUE-109・ISSUE-110 の COM と PNG の値）：面の塗り 黒・面の枠線 黒 0.5pt（Excel の COM は太さ「自動」で、黒い面の上では見えない）・
	* 文字（軸・凡例・タイトル）白・プロット（棒・折れ線）#3F3F3F・円とドーナツのプロットは塗りなし（黒い面が見える）。
	* 軸線・補助の目盛線・折れ線の太さ・系列の色の組は未確定で変えない（41・43〜48 の系列の色の組・円の凡例の `#595959` 9pt はこの ISSUE の範囲外）。
	* ダーク表示でも補正しない（P5）
	*/
	darkStyleMin: 41,
	darkStyleMax: 48,
	darkChartAreaFill: {
		rgb: 0,
		alpha: 1
	},
	darkChartAreaLine: {
		rgb: 0,
		alpha: 1
	},
	darkText: {
		rgb: 16777215,
		alpha: 1
	},
	darkPlotFill: {
		rgb: 4144959,
		alpha: 1
	},
	/**
	* 暗い背景（41〜48）の棒の既定の塗り（単色近似。ISSUE-109・ADR-0031 §決定 4）。Excel の棒は上から下へのグラデーション（明→系列の色に近い色→暗）で、
	* 50% の位置の色は系列の色の彩度を 110% にした色（`a:satMod` 110%。ADR-0028 §決定 6.4 の `a:gradFill`＝50% の位置の色と同じ規則）。`linearGradient` は使わない。1/1000 %
	*/
	darkBarSatMod: 11e4,
	/** 文字（黒） */
	text: { slot: "dk1" },
	/**
	* グラフの面の角丸（`c:roundedCorners` が真・要素が無い時も真。ISSUE-109）。半径は 8.25pt（11px）で、グラフの大きさによらず一定
	* （Excel 16 の PNG：216×144・360×216・720×432pt のどれも 11px。小さい面では 幅÷2・高さ÷2 を超えない）
	*/
	roundedCorners: true,
	cornerRadiusPt: 8.25,
	/** 円・ドーナツの枠線 */
	pieLine: { slot: "lt1" },
	/** 円・ドーナツの枠線の太さ（pt）。Excel は点（`c:dPt`）ごとに 1.5pt を書く（点の線は読まないので、系列に明示が無ければこの値。C9） */
	pieLinePt: 1.5,
	/** 線の太さ（pt）：マーカーの枠・線だけの記号 0.75、折れ線 1.5（c:spPr の無いグラフの値。目盛線・軸線は {@link axisLinePt}） */
	thinLinePt: .75,
	seriesLinePt: 1.5,
	/** マーカーの自動の形と大きさ（pt・C8） */
	autoMarker: "circle",
	markerSizePt: 6,
	/** 文字の大きさ（pt・C4）：タイトル 10（太字）・軸と凡例 10・軸タイトル 10 */
	titlePt: 10,
	titleBold: true,
	axisPt: 10,
	legendPt: 10,
	axisTitlePt: 10,
	/** ドーナツの穴（%・C9） */
	holeSizePercent: 50,
	/** 棒の隙間（%）・重なり（集合 0・積み上げ 100） */
	gapWidth: 150,
	overlapClustered: 0,
	overlapStacked: 100,
	/** 円の開始角（度・12 時から時計回り） */
	firstSliceAng: 0,
	/** 項目軸の文字の回転の既定（`a:bodyPr rot` が無く、自動の形の判定を使わない時＝横棒の項目軸・小さい面。P7・ISSUE-107：回さない） */
	autoRotation: 0,
	/** `c:majorTickMark`・`c:minorTickMark` の要素が無い軸の目盛の印（スキーマの既定。Excel 16 の COM は 4 つとも Cross を返す・C12） */
	tickMark: "cross"
};
/**
* 配置の寸法（C5・C6・C10。pt は × 96/72、`Ratio` は文字の大きさ（px）に対する割合）。
* ISSUE-104 で Excel 16 の COM の値（座標は面の左上から 4pt 内側が原点）に合わせた：
* タイトルの上端は面の上から 6pt・プロット（軸の文字を除く内側）の上と右はタイトル／凡例の下・面の右（凡例の左）から 11pt、
* 軸の文字の左と下は面の左・下（凡例の右・上）から 5pt、凡例は面の縁から 6pt。円・ドーナツは四方 11pt。
*/
var CHART_LAYOUT = {
	/** 軸の文字の左・下の余白（pt。面の左端・下端、凡例の右・上から） */
	marginPt: 5,
	/** プロットの上・右の余白（pt。タイトル・凡例の下／面の右・凡例の左から）。円・ドーナツは四方 */
	plotPadPt: 11,
	/** タイトルの上端（pt） */
	titleTopPt: 6,
	/** 文字の高さ ＝ 大きさ × この値 */
	lineHeightRatio: 1.2,
	/** タイトルの高さ ＝ 大きさ × この値（Excel の Calibri 14pt で 1.435・10pt 太字で 1.52） */
	titleLineHeightRatio: 1.45,
	/** 目盛の長さ（主・補助。pt）と、目盛の文字との間隔（px） */
	tickLengthPt: 5,
	minorTickLengthPt: 3,
	labelGapPx: 3,
	/** 軸タイトルの帯の余白（px） */
	axisTitleGapPx: 4,
	/** 凡例 */
	legendSymbolRatio: .55,
	legendLineSymbolRatio: 1.33,
	legendGapRatio: .4,
	/**
	* 横並び（1 行）で項目と項目の間（C6）。Excel の項目の間隔は、記号＋文字の幅（実測）+ 約 0.47〜0.49em。
	* 文字の幅の推定が広めに出る名前（円 5 項目の凡例）があるので、1 回目・2 回目の照合の両方が 3% 以内に収まる 0.4 にした
	*/
	legendEntryGapRatio: .4,
	/** 横並びの格子の桁の余白（最も長い項目の幅に足す。Excel の桁の幅は文字の幅の推定より少し広い。C6） */
	legendColPadRatio: .1,
	legendRowGapRatio: .675,
	/** 凡例の箱の余白（左右・上下） */
	legendPadRatio: .34,
	legendPadYRatio: 0,
	legendMaxFraction: .4,
	/** 右上の凡例でプロットを凡例の下まで縮める条件：縮めた後のプロットの高さが縮める前のこの割合以上（Excel 16 の COM：小さい面は 0.50・標準は 0.71） */
	legendCornerShrinkMin: .6,
	/** 凡例と面の縁（タイトルの下）の間（pt） */
	legendEdgeGapPt: 6,
	/** 回転した文字の本文の最大の長さの打ち切り（1 行の文字数） */
	maxLineChars: 255,
	/**
	* 項目軸の文字の形の自動判定（縦棒・折れ線の横の項目軸。ISSUE-107・ADR-0031 §決定 5。Excel 16 の c7sw_*・c7w_* の基準値から）。
	* 水平 → 折り返し → 45° → 90° の順に最初に当たった形にする。値の「暫定」は基準値の無い境目（外れた組は照合テストの KNOWN）。
	* - `catRotateMinSlotRatio`：45° と 90° の境目（項目の間隔 ≧ 文字の大きさ × この値なら 45°）。観測は 90° の最大 10.7pt・45° の最小 14.4pt（9pt 文字で 1.6 倍）。暫定
	* - `catRotatedBandExtraPt`：回した文字の帯の厚さ ＝ (幅 ＋ 行の高さ)·sin45° ＋ この値（pt。目盛・文字との間隔を含む。実測 40・50・75pt と ±1pt）。90° も暫定で同じ式
	* - `catWrapRowRatio`：折り返した行の送り ＝ 文字の大きさ × この値（9pt 文字で 11pt。帯は 1 行 18pt・2 行 29pt・3 行 40pt）
	* - `catRotatedSkipPadPt`：90° に回した文字の間引きの余白（pt）。4px のままだと 1 回目の cat30_small が 3 個おきになる（Excel は 2 個おき）
	* - `catReachGapPt`：回した文字が左へ出る分の下限 ＝ (幅 ＋ 行の高さ)·cos45° − 最初の項目の中心の位置 ＋ この値（pt。面の左端から見たプロットの左）
	* - `catBandMaxRatio`：推定の帯が面の高さのこの割合を超える小さい面は、回さず今までどおり間引く（安全弁。暫定。1 回目の catlong_small・catlong_std は推定の幅が広めに出て戻る。0.45 にすると std は回るが、プロットの高さが Excel より 17px 低く数値軸の刻みが変わって 1 回目の KNOWN が増える）
	*/
	catRotateMinSlotRatio: 1.6,
	catRotatedBandExtraPt: 6.5,
	catWrapRowRatio: 11 / 9,
	catRotatedSkipPadPt: 1,
	catReachGapPt: 2,
	catBandMaxRatio: .4
};
/** 文字の幅の係数（P4・暫定）：半角の英数字と記号・半角カナ・全角・その他（em に対する割合） */
var TEXT_WIDTH_EM = {
	ascii: .55,
	halfKana: .5,
	full: 1,
	other: .6
};
/** 既定の色の変化（C3）：周回 k（9 周で繰り返す）の lumMod・lumOff（1/1000 %）。k=0 は変化なし */
var CHART_COLOR_VARIATIONS = [
	{},
	{ lumMod: 6e4 },
	{
		lumMod: 8e4,
		lumOff: 2e4
	},
	{ lumMod: 8e4 },
	{
		lumMod: 6e4,
		lumOff: 4e4
	},
	{ lumMod: 5e4 },
	{
		lumMod: 7e4,
		lumOff: 3e4
	},
	{ lumMod: 7e4 },
	{
		lumMod: 5e4,
		lumOff: 5e4
	}
];
var ACCENT_SLOTS = [
	"accent1",
	"accent2",
	"accent3",
	"accent4",
	"accent5",
	"accent6"
];
/** 変換の子の列（`color.ts` の `applyTransforms` に渡す本文） */
function transformBody(lumMod, lumOff) {
	let s = "";
	if (lumMod !== void 0) s += `<a:lumMod val="${lumMod}"/>`;
	if (lumOff !== void 0) s += `<a:lumOff val="${lumOff}"/>`;
	return s === "" ? null : s;
}
/** テーマの色（無ければ既定のテーマ） */
function themeRgb(theme, slot) {
	return theme.get(slot) ?? DEFAULT_THEME.get(slot) ?? 0;
}
/** 既定の色の指定を色にする */
function specColor(theme, spec) {
	return applyTransforms({
		rgb: themeRgb(theme, spec.slot),
		alpha: 1
	}, transformBody(spec.lumMod, spec.lumOff));
}
/**
* 既定の系列（・点）の色（C3）：位置 `index`（系列は `c:idx`・点は点の位置）に対し、テーマの `accent((index mod 6)+1)` に
* 周回 k＝floor(index / 6) の変化（{@link CHART_COLOR_VARIATIONS}。9 周で繰り返す）を掛ける。
*/
function autoColor(theme, index) {
	const i = Number.isFinite(index) && index > 0 ? Math.floor(index) : 0;
	const base = {
		rgb: themeRgb(theme, ACCENT_SLOTS[i % 6]),
		alpha: 1
	};
	const v = CHART_COLOR_VARIATIONS[Math.floor(i / 6) % CHART_COLOR_VARIATIONS.length];
	return applyTransforms(base, transformBody(v.lumMod, v.lumOff));
}
var hex = (rgb) => `#${(rgb & 16777215).toString(16).padStart(6, "0")}`;
var clamp01 = (x) => x > 0 ? x < 1 ? x : 1 : 0;
var toColor = (c) => ({
	color: hex(c.rgb),
	opacity: clamp01(c.alpha)
});
/** 1 文字の幅（em）。半角の英数字と記号 0.55・半角カナ 0.5・全角 1・その他 0.6（P4・暫定） */
function charEm(cp) {
	if (cp >= 32 && cp <= 126) return TEXT_WIDTH_EM.ascii;
	if (cp >= 65377 && cp <= 65439) return TEXT_WIDTH_EM.halfKana;
	if (cp >= 4352 && cp <= 4447 || cp >= 11904 && cp <= 42191 || cp >= 44032 && cp <= 55203 || cp >= 63744 && cp <= 64255 || cp >= 65072 && cp <= 65135 || cp >= 65281 && cp <= 65376 || cp >= 65504 && cp <= 65510 || cp >= 131072 && cp <= 262141) return TEXT_WIDTH_EM.full;
	return TEXT_WIDTH_EM.other;
}
/** 文字の幅の推定（px）。`sizePx` は文字の大きさ（px） */
function estimateTextWidth(text, sizePx) {
	let em = 0;
	for (const ch of text) em += charEm(ch.codePointAt(0));
	return em * sizePx;
}
/** 1 行を最大 `max` 文字（既定 255）に切る（超えたら末尾を `…` にする） */
function truncateLine(text, max = CHART_LAYOUT.maxLineChars) {
	if (text.length <= max) return text;
	const cps = Array.from(text);
	return cps.length <= max ? text : `${cps.slice(0, max - 1).join("")}…`;
}
/** 推定の幅が `maxWidth` に入るように末尾を `…` で切る */
function fitText(text, sizePx, maxWidth) {
	if (estimateTextWidth(text, sizePx) <= maxWidth) return text;
	const cps = Array.from(text);
	let lo = 0;
	let hi = cps.length;
	while (lo < hi) {
		const mid = lo + hi + 1 >> 1;
		if (estimateTextWidth(`${cps.slice(0, mid).join("")}…`, sizePx) <= maxWidth) lo = mid;
		else hi = mid - 1;
	}
	return lo === 0 ? "" : `${cps.slice(0, lo).join("")}…`;
}
/** 経路の数値：有限に直し、小数 2 桁に丸める（指数表記にしない） */
function f(x) {
	const r = Math.round((Number.isFinite(x) ? Math.max(-1e7, Math.min(1e7, x)) : 0) * 100) / 100;
	return String(r === 0 ? 0 : r);
}
/** 有限でなければ 0 にする（場面の数値は常に有限） */
var fin = (x) => Number.isFinite(x) ? x : 0;
/** manualLayout を矩形に当てる（`edge`＝グラフの面に対する位置の割合・`factor`＝自動の位置からのずれ。ADR-0031 §決定 6） */
function applyManual(l, W, H, auto) {
	const x = l.x === void 0 ? auto.x : l.xMode === "edge" ? l.x * W : auto.x + l.x * W;
	const y = l.y === void 0 ? auto.y : l.yMode === "edge" ? l.y * H : auto.y + l.y * H;
	const w = l.w === void 0 ? auto.w : l.wMode === "edge" ? l.w * W - x : l.w * W;
	const h = l.h === void 0 ? auto.h : l.hMode === "edge" ? l.h * H - y : l.h * H;
	return {
		x: fin(x),
		y: fin(y),
		w: Math.max(1, fin(w)),
		h: Math.max(1, fin(h))
	};
}
var hasManual = (l) => l !== void 0 && (l.x !== void 0 || l.y !== void 0 || l.w !== void 0 || l.h !== void 0);
/** 要素数の上限を超えた（内部で投げ、layoutChart が null にする） */
var SceneOverflow = class extends Error {};
/** `c:style` が 41〜48（暗い背景）か（ISSUE-109） */
function isDarkStyle(m) {
	return m.style !== void 0 && m.style >= CHART_DEFAULTS.darkStyleMin && m.style <= CHART_DEFAULTS.darkStyleMax;
}
/** 目盛線・軸線の既定の色：`c14:style` を持てば #898989、無ければ `normal`（ISSUE-109・CHART_DEFAULTS.c14Line） */
function lineSpec(cx, normal) {
	return cx.model.c14Style !== void 0 ? CHART_DEFAULTS.c14Line : normal;
}
function push(cx, item) {
	if (cx.items.length >= 5e3) throw new SceneOverflow();
	cx.items.push(item);
}
/** 文字の見た目（部品 → グラフ全体 → 既定の順・§決定 8） */
function textLook(cx, part, defaultPt, defaultBold = false) {
	const pt = part?.sizePt ?? cx.model.text.sizePt ?? defaultPt;
	const color = part?.color ?? cx.model.text.color ?? (isDarkStyle(cx.model) ? CHART_DEFAULTS.darkText : specColor(cx.theme, CHART_DEFAULTS.text));
	return {
		px: pt * PX_PER_PT,
		bold: part?.bold ?? cx.model.text.bold ?? defaultBold,
		color: toColor(color)
	};
}
/** 線を解く（明示の `a:ln` ＞ 既定。`none` と既定が無ければ null） */
function resolveStroke(cx, line, def, caps = {}) {
	if (line?.paint?.kind === "none") return null;
	let rgba = null;
	if (line?.paint?.kind === "solid") rgba = line.paint.color;
	else if (def !== null) rgba = "slot" in def.spec ? specColor(cx.theme, def.spec) : def.spec;
	if (rgba === null) return null;
	const width = line?.widthEmu !== void 0 ? line.widthEmu / 9525 : (def?.widthPt ?? .75) * PX_PER_PT;
	return {
		...toColor(rgba),
		width: Math.max(0, fin(width)),
		dash: line?.dash ?? "solid",
		cap: caps.cap ?? "butt",
		join: caps.join ?? "miter"
	};
}
/** 塗りを解く（`none` は null・明示が無ければ `def`） */
function resolveFill(paint, def) {
	if (paint === void 0) return def;
	return paint.kind === "none" ? null : toColor(paint.color);
}
function pushText(cx, x, y, text, look, anchor, baseline, rotate = 0) {
	push(cx, {
		kind: "text",
		x: fin(x),
		y: fin(y),
		text,
		size: look.px,
		bold: look.bold,
		color: look.color,
		anchor,
		baseline,
		rotate: fin(rotate)
	});
}
/** マーカー 1 つの経路（中心 (x, y)・半径 r）。`filled` が偽の形（線だけ）は塗らない */
function markerPath(sym, x, y, r) {
	const a = f(x - r);
	const b = f(x + r);
	const t = f(y - r);
	const u = f(y + r);
	switch (sym) {
		case "square": return {
			d: `M${a} ${t}L${b} ${t}L${b} ${u}L${a} ${u}Z`,
			filled: true
		};
		case "diamond": return {
			d: `M${f(x)} ${t}L${b} ${f(y)}L${f(x)} ${u}L${a} ${f(y)}Z`,
			filled: true
		};
		case "triangle": return {
			d: `M${f(x)} ${t}L${b} ${u}L${a} ${u}Z`,
			filled: true
		};
		case "x": return {
			d: `M${a} ${t}L${b} ${u}M${b} ${t}L${a} ${u}`,
			filled: false
		};
		case "plus": return {
			d: `M${a} ${f(y)}L${b} ${f(y)}M${f(x)} ${t}L${f(x)} ${u}`,
			filled: false
		};
		case "star": return {
			d: `M${a} ${t}L${b} ${u}M${b} ${t}L${a} ${u}M${a} ${f(y)}L${b} ${f(y)}M${f(x)} ${t}L${f(x)} ${u}`,
			filled: false
		};
		case "dash": return {
			d: `M${a} ${f(y)}L${b} ${f(y)}`,
			filled: false
		};
		case "dot": {
			const h = r * .5;
			return {
				d: `M${f(x - h)} ${f(y - h)}L${f(x + h)} ${f(y - h)}L${f(x + h)} ${f(y + h)}L${f(x - h)} ${f(y + h)}Z`,
				filled: true
			};
		}
		default: return {
			d: `M${a} ${f(y)}A${f(r)} ${f(r)} 0 1 1 ${b} ${f(y)}A${f(r)} ${f(r)} 0 1 1 ${a} ${f(y)}Z`,
			filled: true
		};
	}
}
/**
* 系列のマーカーの形（`none` の時だけ null）。自動・無指定は {@link CHART_DEFAULTS.autoMarker}（C8・C13）。
* グループの `c:marker`（`lineMarkers`）は見ない：Excel 16 は、系列に `c:marker` が無ければ、グループの `c:marker` が 0 でも無くても、自動のマーカーを描く（COM の MarkerStyle が自動）。
*/
function markerSymbol(s) {
	const sym = s.marker?.symbol;
	if (sym === "none") return null;
	if (sym !== void 0 && sym !== "auto") return sym;
	return CHART_DEFAULTS.autoMarker;
}
/** 棒・折れ線で `varyColors` が点の位置の色になるか（1 系列の時だけ・C3） */
var varyByPoint = (model) => model.varyColors && model.series.length === 1;
/** 系列の既定の色（明示が無い時） */
var seriesAuto = (cx, s) => toColor(autoColor(cx.theme, s.idx));
/**
* 点 i の塗り：`c:dPt` ＞ 系列の明示の塗り ＞ 既定（`varyColors` の効く時は点の位置・そうでなければ系列の `idx`）。
* `none` は null（描かない）。
*/
function pointFill(cx, s, i, vary) {
	const dpt = s.points?.get(i);
	if (dpt !== void 0) return resolveFill(dpt, null);
	if (s.fill !== void 0) return resolveFill(s.fill, null);
	const auto = vary ? autoColor(cx.theme, i) : autoColor(cx.theme, s.idx);
	if (cx.model.type === "bar" && isDarkStyle(cx.model)) return toColor(applyTransforms(auto, `<a:satMod val="${CHART_DEFAULTS.darkBarSatMod}"/>`));
	return toColor(auto);
}
/**
* グラフのタイトルの行（無ければ空）。文字の無い `c:title` は系列が 1 つなら系列名（C10）。
* `c:title` が無くても `c:autoTitleDeleted` が偽で系列が 1 つなら、Excel は系列名を自動のタイトルに出す（Excel 16 の照合）。
*/
function titleLines(model) {
	const t = model.title;
	if (t === void 0) {
		const only = model.series.length === 1 ? model.series[0].name : "";
		return model.autoTitleDeleted || only === "" ? [] : [truncateLine(only)];
	}
	if (t.lines.length > 0) return t.lines.map((l) => truncateLine(l));
	const only = model.series.length === 1 ? model.series[0].name : "";
	return only === "" ? [] : [truncateLine(only)];
}
/**
* 文字の無い `c:title`（自動）で系列が 2 つ以上か。Excel は言語ごとの代替の文字（日本語 UI では「グラフ タイトル」）を出して帯を取る（C10）。
* 代替の文字は言語に依るので描かず、帯（1 行ぶんの高さ）だけ取る。
*/
function isBlankAutoTitle(model) {
	const t = model.title;
	return t !== void 0 && t.lines.length === 0 && model.series.length >= 2;
}
/** 凡例の項目を作る（系列名／`varyColors` の項目・`c:legendEntry` の削除を除く。並びは系列の順・C6） */
function legendEntries(cx, legend) {
	const m = cx.model;
	const out = [];
	const del = new Set(legend.deleted);
	const raw = [];
	if (m.type === "pie" ? true : m.type === "doughnut" ? m.varyColors : varyByPoint(m)) {
		const s = m.series[0];
		if (s === void 0) return out;
		m.categories.forEach((name, i) => {
			raw.push(entryFor(cx, m.type, s, truncateLine(name), pointFill(cx, s, i, m.varyColors)));
		});
	} else m.series.forEach((s, pos) => {
		const name = s.name === "" ? `Series ${pos + 1}` : truncateLine(s.name);
		raw.push(entryFor(cx, m.type, s, name, pointFill(cx, s, 0, false)));
	});
	const stacked = m.grouping === "stacked" || m.grouping === "percentStacked";
	const verticalLegend = legend.pos === "r" || legend.pos === "l" || legend.pos === "tr";
	if (m.type === "bar" && verticalLegend && (m.barDir === "bar" ? !stacked : stacked)) raw.reverse();
	raw.forEach((e, i) => {
		if (!del.has(i)) out.push(e);
	});
	return out;
}
/** 凡例の項目 1 つ（棒・円は四角・折れ線は線＋マーカー） */
function entryFor(cx, type, s, text, fill) {
	if (type !== "line") return {
		text,
		symbol: "square",
		fill,
		stroke: type === "bar" ? resolveStroke(cx, s.line, null) : resolveStroke(cx, s.line, {
			spec: CHART_DEFAULTS.pieLine,
			widthPt: CHART_DEFAULTS.pieLinePt
		}),
		marker: null
	};
	const lineColor = s.line?.paint?.kind === "solid" ? toColor(s.line.paint.color) : fill;
	const stroke = resolveStroke(cx, s.line, {
		spec: autoColor(cx.theme, s.idx),
		widthPt: CHART_DEFAULTS.seriesLinePt
	}, {
		cap: "round",
		join: "round"
	});
	const sym = markerSymbol(s);
	let marker = null;
	if (sym !== null) marker = {
		sym,
		fill: resolveFill(s.marker?.fill, lineColor ?? fill),
		stroke: resolveStroke(cx, s.marker?.line, {
			spec: autoColor(cx.theme, s.idx),
			widthPt: CHART_DEFAULTS.thinLinePt
		}),
		sizePx: (s.marker?.size ?? CHART_DEFAULTS.markerSizePt) * PX_PER_PT
	};
	return {
		text,
		symbol: "line",
		fill,
		stroke,
		marker
	};
}
/**
* 凡例の項目を箱の中に並べる（右・左・右上は縦、上・下は横に並べて折り返す。入らない項目は出さない・§決定 7）。
* 項目が 1 つも入らなければ null。
*/
function buildLegend(entries, pos, look, maxW, maxH) {
	const L = CHART_LAYOUT;
	const s = look.px;
	const pad = s * L.legendPadRatio;
	const padY = s * L.legendPadYRatio;
	const gap = s * L.legendGapRatio;
	const rowH = s * L.lineHeightRatio + s * L.legendRowGapRatio;
	const maxRows = Math.floor((maxH - 2 * padY) / rowH);
	if (maxRows < 1 || entries.length === 0) return null;
	const vertical = pos === "r" || pos === "l" || pos === "tr";
	const symOf = (e) => s * (e.symbol === "line" ? L.legendLineSymbolRatio : L.legendSymbolRatio);
	const placed = [];
	if (vertical) {
		const innerW = Math.max(0, maxW - 2 * pad);
		let widest = 0;
		const shown = entries.slice(0, maxRows);
		shown.forEach((e, k) => {
			const symW = symOf(e);
			const text = fitText(e.text, s, Math.max(0, innerW - symW - gap));
			const w = symW + gap + estimateTextWidth(text, s);
			widest = Math.max(widest, w);
			placed.push({
				entry: e,
				x: pad,
				y: padY + k * rowH + rowH / 2,
				symW,
				text
			});
		});
		return {
			w: widest + 2 * pad,
			h: shown.length * rowH + 2 * padY,
			placed
		};
	}
	const innerW = Math.max(0, maxW - 2 * pad);
	const gapX = s * L.legendEntryGapRatio;
	const cells = entries.map((e) => {
		const symW = symOf(e);
		const text = fitText(e.text, s, Math.max(0, innerW - symW - gap));
		return {
			entry: e,
			symW,
			text,
			w: symW + gap + estimateTextWidth(text, s)
		};
	});
	const flowW = cells.reduce((a, c) => a + c.w, 0) + (cells.length - 1) * gapX;
	if (flowW <= innerW) {
		let x = pad;
		for (const c of cells) {
			placed.push({
				entry: c.entry,
				x,
				y: padY + rowH / 2,
				symW: c.symW,
				text: c.text
			});
			x += c.w + gapX;
		}
		return {
			w: flowW + 2 * pad,
			h: rowH + 2 * padY,
			placed
		};
	}
	const colW = cells.reduce((m, c) => Math.max(m, c.w), 0) + s * L.legendColPadRatio;
	const maxCols = Math.max(1, Math.floor(innerW / colW));
	let rowCount = Math.ceil(cells.length / maxCols);
	let shown = cells.length;
	if (rowCount > maxRows) {
		rowCount = maxRows;
		shown = rowCount * maxCols;
	}
	const cols = Math.ceil(Math.min(shown, cells.length) / rowCount);
	cells.slice(0, shown).forEach((c, k) => {
		placed.push({
			entry: c.entry,
			x: pad + k % cols * colW,
			y: padY + Math.floor(k / cols) * rowH + rowH / 2,
			symW: c.symW,
			text: c.text
		});
	});
	return {
		w: cols * colW + 2 * pad,
		h: rowCount * rowH + 2 * padY,
		placed
	};
}
/** 凡例を場面へ出す（箱の左上 (ox, oy)） */
function emitLegend(cx, box, look, ox, oy) {
	const s = look.px;
	const gap = s * CHART_LAYOUT.legendGapRatio;
	for (const p of box.placed) {
		const e = p.entry;
		const x = ox + p.x;
		const y = oy + p.y;
		if (e.symbol === "square") {
			const q = s * CHART_LAYOUT.legendSymbolRatio;
			if (e.fill !== null || e.stroke !== null) push(cx, {
				kind: "rect",
				x: fin(x),
				y: fin(y - q / 2),
				width: q,
				height: q,
				fill: e.fill,
				stroke: e.stroke
			});
		} else {
			if (e.stroke !== null) push(cx, {
				kind: "line",
				x1: fin(x),
				y1: fin(y),
				x2: fin(x + p.symW),
				y2: fin(y),
				stroke: e.stroke
			});
			if (e.marker !== null) {
				const m = markerPath(e.marker.sym, x + p.symW / 2, y, Math.min(e.marker.sizePx, s * 1.2) / 2);
				push(cx, {
					kind: "path",
					d: m.d,
					fill: m.filled ? e.marker.fill : null,
					stroke: e.marker.stroke ?? (m.filled ? null : e.fill === null ? null : strokeFrom(e.fill))
				});
			}
		}
		pushText(cx, x + p.symW + gap, y, p.text, look, "start", "middle");
	}
}
/** 色から細い線（線だけの記号の既定） */
function strokeFrom(c) {
	return {
		...c,
		width: CHART_DEFAULTS.thinLinePt * PX_PER_PT,
		dash: "solid",
		cap: "butt",
		join: "miter"
	};
}
var TAU = Math.PI * 2;
/** 12 時から時計回りの角度 θ（rad）の円周上の点 */
function polar(cx0, cy0, r, theta) {
	return [cx0 + r * Math.sin(theta), cy0 - r * Math.cos(theta)];
}
function circlePath(cx0, cy0, r, sweep) {
	return `M${f(cx0 - r)} ${f(cy0)}A${f(r)} ${f(r)} 0 1 ${sweep} ${f(cx0 + r)} ${f(cy0)}A${f(r)} ${f(r)} 0 1 ${sweep} ${f(cx0 - r)} ${f(cy0)}Z`;
}
/** 扇（`a0` から `a1` へ時計回り）。全周なら円 */
function sectorPath(cx0, cy0, r, a0, a1) {
	if (a1 - a0 >= TAU - 1e-6) return circlePath(cx0, cy0, r, 1);
	const [x0, y0] = polar(cx0, cy0, r, a0);
	const [x1, y1] = polar(cx0, cy0, r, a1);
	return `M${f(cx0)} ${f(cy0)}L${f(x0)} ${f(y0)}A${f(r)} ${f(r)} 0 ${a1 - a0 > Math.PI ? 1 : 0} 1 ${f(x1)} ${f(y1)}Z`;
}
/** 輪の扇（外半径 ro・内半径 ri）。全周なら外円と逆回りの内円（穴） */
function ringPath(cx0, cy0, ro, ri, a0, a1) {
	if (a1 - a0 >= TAU - 1e-6) return circlePath(cx0, cy0, ro, 1) + circlePath(cx0, cy0, ri, 0);
	const [ox0, oy0] = polar(cx0, cy0, ro, a0);
	const [ox1, oy1] = polar(cx0, cy0, ro, a1);
	const [ix0, iy0] = polar(cx0, cy0, ri, a0);
	const [ix1, iy1] = polar(cx0, cy0, ri, a1);
	const large = a1 - a0 > Math.PI ? 1 : 0;
	return `M${f(ox0)} ${f(oy0)}A${f(ro)} ${f(ro)} 0 ${large} 1 ${f(ox1)} ${f(oy1)}L${f(ix1)} ${f(iy1)}A${f(ri)} ${f(ri)} 0 ${large} 0 ${f(ix0)} ${f(iy0)}Z`;
}
/** 円・ドーナツの扇を場面へ出す（円は最初の系列だけ・ドーナツは系列ごとの同心の輪で最初の系列が内側・C9） */
function emitPie(cx, plot) {
	const m = cx.model;
	const cxp = plot.x + plot.w / 2;
	const cyp = plot.y + plot.h / 2;
	const R = Math.max(0, Math.min(plot.w, plot.h) / 2);
	const start = (m.firstSliceAng ?? CHART_DEFAULTS.firstSliceAng) * Math.PI / 180;
	const rings = m.type === "doughnut" ? m.series : m.series.slice(0, 1);
	const hole = m.type === "doughnut" ? R * (m.holeSize ?? CHART_DEFAULTS.holeSizePercent) / 100 : 0;
	const ringW = (R - hole) / Math.max(1, rings.length);
	rings.forEach((s, j) => {
		const abs = s.values.map((v) => typeof v === "number" && Number.isFinite(v) ? Math.abs(v) : 0);
		const total = abs.reduce((a, b) => a + b, 0);
		if (!(total > 0)) return;
		const ro = m.type === "doughnut" ? hole + (j + 1) * ringW : R;
		const ri = m.type === "doughnut" ? ro - ringW : 0;
		const stroke = resolveStroke(cx, s.line, {
			spec: CHART_DEFAULTS.pieLine,
			widthPt: CHART_DEFAULTS.pieLinePt
		});
		const wedges = [];
		let a0 = start;
		abs.forEach((v, i) => {
			if (!(v > 0)) return;
			const a1 = a0 + v / total * TAU;
			wedges.push({
				i,
				a0,
				a1
			});
			a0 = a1;
		});
		const path = (x0, x1) => m.type === "doughnut" ? ringPath(cxp, cyp, ro, ri, x0, x1) : sectorPath(cxp, cyp, ro, x0, x1);
		const runs = cx.thin && stroke !== null && stroke.opacity >= 1 && stroke.dash === "solid" ? thinSectorRuns(wedges.map((w) => ({
			a0: w.a0,
			a1: w.a1,
			fixed: s.points?.has(w.i) === true
		})), ro, stroke.width) : [];
		let r = 0;
		for (let k = 0; k < wedges.length; k++) {
			const w = wedges[k];
			const run = r < runs.length && runs[r][0] === k ? runs[r] : null;
			if (run !== null && stroke !== null) {
				const last = wedges[run[1]];
				push(cx, {
					kind: "path",
					d: path(w.a0, last.a1),
					fill: {
						color: stroke.color,
						opacity: stroke.opacity
					},
					stroke
				});
				k = run[1];
				r++;
				continue;
			}
			push(cx, {
				kind: "path",
				d: path(w.a0, w.a1),
				fill: pointFill(cx, s, w.i, m.varyColors),
				stroke
			});
		}
	});
}
/**
* 1 つの項目名を、幅 `slotPx` に収まるよう空白で折り返した行にする（語を左から詰める）。
* 1 行に収まれば `[text]`。どれかの語の幅が `slotPx` を超える（空白が無い・語が長い）時は折り返せないので `null`。
*/
function wrapCategoryLabel(text, sizePx, slotPx) {
	if (estimateTextWidth(text, sizePx) <= slotPx) return [text];
	const words = text.split(/\s+/).filter((w) => w !== "");
	if (words.length < 2) return null;
	const lines = [];
	let cur = "";
	for (const w of words) {
		if (estimateTextWidth(w, sizePx) > slotPx) return null;
		const next = cur === "" ? w : `${cur} ${w}`;
		if (cur === "" || estimateTextWidth(next, sizePx) <= slotPx) cur = next;
		else {
			lines.push(cur);
			cur = w;
		}
	}
	lines.push(cur);
	return lines;
}
/**
* 縦棒・折れ線の項目軸の文字の形を選ぶ（ISSUE-107・ADR-0031 §決定 5。`a:bodyPr rot` の明示が無い時だけ呼ぶ）。
* 上から順に判定し、最初に当たった形にする。
* 1. 水平：どの項目名の幅（推定）も `slotPx`（項目の間隔）以下。
* 2. 折り返し：幅が超える項目名が全部、空白を含み、どの語も `slotPx` に入る（`rows` は行数の最大）。
* 3. 45°：`slotPx` が 45° の境目（{@link CHART_LAYOUT.catRotateMinSlotRatio} × 文字の大きさ）以上。
* 4. 90°：それより狭い（呼び出し側が入らない分を間引く）。
*/
function planCategoryLabelForm(labels, sizePx, slotPx) {
	if (Number.isNaN(slotPx)) return {
		form: "horizontal",
		rows: 1
	};
	let over = false;
	let wrappable = true;
	let rows = 1;
	for (const t of labels) {
		if (estimateTextWidth(t, sizePx) <= slotPx) continue;
		over = true;
		const ls = wrapCategoryLabel(t, sizePx, slotPx);
		if (ls === null) {
			wrappable = false;
			break;
		}
		rows = Math.max(rows, ls.length);
	}
	if (!over) return {
		form: "horizontal",
		rows: 1
	};
	if (wrappable) return {
		form: "wrap",
		rows
	};
	return {
		form: slotPx >= CHART_LAYOUT.catRotateMinSlotRatio * sizePx ? "rotate45" : "rotate90",
		rows: 1
	};
}
/** 回転した文字の枠（幅 w・高さ h を θ 度回す）の軸に沿った大きさと直交する大きさ */
function rotatedExtent(w, h, deg) {
	const t = deg * Math.PI / 180;
	const c = Math.abs(Math.cos(t));
	const s = Math.abs(Math.sin(t));
	return {
		along: w * c + h * s,
		perp: w * s + h * c
	};
}
/** 軸のタイトルの行（無ければ空）。軸が消されていても出す（Excel の挙動に合わせない・暫定） */
function axisTitleLines(axis) {
	const t = axis?.title;
	return t === void 0 ? [] : t.lines.map((l) => truncateLine(l));
}
function cartCtx(cx) {
	const m = cx.model;
	const horizontalBars = m.type === "bar" && m.barDir === "bar";
	return {
		horizontalBars,
		valueVertical: !horizontalBars,
		catLook: textLook(cx, m.catAxis?.text, CHART_DEFAULTS.axisPt),
		valLook: textLook(cx, m.valAxis?.text, CHART_DEFAULTS.axisPt),
		catTitleLook: textLook(cx, m.catAxis?.title?.text, CHART_DEFAULTS.axisTitlePt),
		valTitleLook: textLook(cx, m.valAxis?.title?.text, CHART_DEFAULTS.axisTitlePt),
		catAxisTitle: axisTitleLines(m.catAxis),
		valAxisTitle: axisTitleLines(m.valAxis),
		catRot: m.catAxis?.text.rotDeg ?? CHART_DEFAULTS.autoRotation,
		catRotExplicit: m.catAxis?.text.rotDeg !== void 0,
		catLabelsAll: m.categories.map((c) => truncateLine(c))
	};
}
/**
* 縦棒・折れ線の項目軸の文字の形を決め、回転・間引き・行・帯の厚さ・左の張り出しにする（ISSUE-107。規則は {@link planCategoryLabelForm}）。
* `a:bodyPr rot` の明示が無い時だけ呼ぶ。`c:tickLblSkip` の明示は最優先で、形は表示する項目の間隔（間隔 × skip）で選ぶ。
* 回した・折り返した帯が面の高さの {@link CHART_LAYOUT.catBandMaxRatio} を超える小さい面は `null`（今までどおり回さず間引く）。
*/
function autoCatLabels(cx, cc, n, catLen, midCat, reversed, explicitSkip) {
	if (n <= 0) return null;
	const L = CHART_LAYOUT;
	const px = cc.catLook.px;
	const lineH = px * L.lineHeightRatio;
	const spacing = categorySpacing(n, catLen, midCat);
	const ex = explicitSkip !== void 0 && Number.isFinite(explicitSkip) && explicitSkip >= 1 ? Math.min(n, Math.floor(explicitSkip)) : void 0;
	const step = ex ?? 1;
	const slot = spacing * step;
	const labels = cc.catLabelsAll;
	const plan = planCategoryLabelForm(categoryLabelIndices(n, step).map((i) => labels[i]), px, slot);
	const idx = categoryLabelIndices(n, plan.form === "rotate90" && ex === void 0 ? categoryLabelSkip(n, spacing, lineH, void 0, L.catRotatedSkipPadPt * PX_PER_PT) : step);
	const lines = labels.map((t) => plan.form === "wrap" ? wrapCategoryLabel(t, px, slot) ?? [t] : [t]);
	const widths = idx.map((i) => estimateTextWidth(labels[i], px));
	const tickGap = L.tickLengthPt * PX_PER_PT + L.labelGapPx;
	const rotated = plan.form === "rotate45" || plan.form === "rotate90";
	let perp = lineH;
	if (plan.form === "wrap") perp = lineH + (plan.rows - 1) * px * L.catWrapRowRatio;
	else if (rotated) {
		const wMax = widths.reduce((a, b) => Math.max(a, b), 0);
		perp = Math.max(lineH, (wMax + lineH) * Math.SQRT1_2 + L.catRotatedBandExtraPt * PX_PER_PT - tickGap);
	}
	if (plan.form !== "horizontal" && perp + tickGap > cx.H * L.catBandMaxRatio) return null;
	let reach = 0;
	if (rotated) {
		const centers = categoryPositions(n, catLen, {
			midCat,
			reversed
		});
		idx.forEach((i, k) => {
			reach = Math.max(reach, (widths[k] + lineH) * Math.SQRT1_2 - centers[i]);
		});
		reach = Math.max(0, reach + L.catReachGapPt * PX_PER_PT);
	}
	return {
		rot: rotated ? plan.form === "rotate45" ? -45 : -90 : 0,
		idx,
		lines,
		perp,
		reach
	};
}
/**
* プロットの大きさ（px）から、軸の範囲・目盛の文字・項目の文字の間引き・文字を置く側を決める。
* `scaleW` は横向きの数値軸（横棒）の目盛の数を決める長さ（px。省略はプロットの幅）。
* Excel は横棒の刻みを、項目軸の文字・目盛の文字の余白を引く前の幅（自動配置の外側）で決める（C1・ISSUE-110）。
*/
function solveAxes(cx, cc, plotW, plotH, scaleW) {
	const m = cx.model;
	const L = CHART_LAYOUT;
	const valAxis = m.valAxis;
	const catAxis = m.catAxis;
	const valLen = cc.valueVertical ? plotH : scaleW ?? plotW;
	const catLen = cc.valueVertical ? plotW : plotH;
	const lineH = (look) => look.px * L.lineHeightRatio;
	const scaleInput = {
		axis: valAxis,
		extent: dataExtent(m.grouping, m.series),
		percentStacked: m.grouping === "percentStacked",
		lengthPx: valLen,
		labelHeightPx: lineH(cc.valLook)
	};
	const labelsOf = (sc) => {
		const code = tickFormatCode(valAxis, m.series[0]?.formatCode, sc.percent);
		return sc.majorTicks.map((t) => truncateLine(tickLabel(code, t)));
	};
	const widest = (ls) => ls.reduce((a, t) => Math.max(a, estimateTextWidth(t, cc.valLook.px)), 0);
	let scale = computeValueScale(scaleInput);
	let valLabels = labelsOf(scale);
	if (!cc.valueVertical) {
		scale = computeValueScale({
			...scaleInput,
			labelWidthPx: widest(valLabels)
		});
		valLabels = labelsOf(scale);
	}
	const showValLabels = valAxis?.deleted !== true && valAxis?.tickLblPos !== "none";
	const valPerp = showValLabels ? cc.valueVertical ? widest(valLabels) : lineH(cc.valLook) : 0;
	const midCat = m.type === "line" && (valAxis?.crossBetween ?? catAxis?.crossBetween) === "midCat";
	const catMaxMin = catAxis?.orientation === "maxMin";
	const n = m.categories.length;
	const showCatLabels = catAxis?.deleted !== true && catAxis?.tickLblPos !== "none";
	const auto = cc.valueVertical && !cc.catRotExplicit && showCatLabels ? autoCatLabels(cx, cc, n, catLen, midCat, catMaxMin, catAxis?.tickLblSkip) : null;
	let catIdx;
	let catPerp;
	if (auto !== null) {
		catIdx = auto.idx;
		catPerp = auto.perp;
	} else {
		let maxAlong = 0;
		const dims = cc.catLabelsAll.map((t) => {
			const e = rotatedExtent(estimateTextWidth(t, cc.catLook.px), lineH(cc.catLook), cc.catRot);
			const d = cc.valueVertical ? e : {
				along: e.perp,
				perp: e.along
			};
			maxAlong = Math.max(maxAlong, d.along);
			return d;
		});
		const skip = categoryLabelSkip(n, categorySpacing(n, catLen, midCat), maxAlong, catAxis?.tickLblSkip);
		catIdx = showCatLabels ? categoryLabelIndices(n, skip) : [];
		catPerp = catIdx.reduce((a, i) => Math.max(a, dims[i].perp), 0);
	}
	const crossFrac = valueFraction(scale, crossValue(scale, catAxis));
	const pos = catAxis?.tickLblPos;
	const atMax = crossFrac >= 1 - 1e-9 && crossFrac > 1e-9;
	const atMin = crossFrac <= 1e-9;
	const catAtLine = pos !== "low" && pos !== "high";
	const crossAuto = (catAxis?.crosses ?? "autoZero") === "autoZero" && catAxis?.crossesAt === void 0;
	const maxByData = atMax && crossAuto && !scale.reversed;
	const catSide = pos === "high" ? "high" : pos === "low" ? "low" : atMax && !maxByData ? "high" : "low";
	const catReserve = !catAtLine || (catSide === "high" ? atMax : atMin);
	const valLineHigh = catMaxMin !== (valAxis?.crosses === "max");
	const vpos = valAxis?.tickLblPos;
	const valSide = vpos === "high" ? "high" : vpos === "low" ? "low" : valLineHigh ? "high" : "low";
	return {
		scale,
		valLabels,
		showValLabels,
		valPerp,
		catLabels: cc.catLabelsAll,
		catLines: auto?.lines ?? cc.catLabelsAll.map((t) => [t]),
		showCatLabels,
		catIdx,
		catPerp,
		catRot: auto?.rot ?? cc.catRot,
		catReach: auto?.reach ?? 0,
		valLineHigh,
		valSide,
		catSide,
		catAtLine,
		catReserve,
		crossFrac,
		midCat,
		catMaxMin
	};
}
/** 軸の文字・軸タイトルが外側に取る余白 */
function axisInsets(cc, ax) {
	const L = CHART_LAYOUT;
	const tick = L.tickLengthPt * PX_PER_PT;
	const ins = {
		l: 0,
		r: 0,
		t: 0,
		b: 0
	};
	const band = (horizontalAxis, side, size) => {
		if (size <= 0) return;
		const key = horizontalAxis ? side === "low" ? "b" : "t" : side === "low" ? "l" : "r";
		ins[key] += size;
	};
	const valBand = ax.showValLabels ? ax.valPerp + tick + L.labelGapPx : 0;
	const catBand = ax.showCatLabels && ax.catReserve ? ax.catPerp + tick + L.labelGapPx : 0;
	band(true, cc.valueVertical ? ax.catSide : ax.valSide, cc.valueVertical ? catBand : valBand);
	band(false, cc.valueVertical ? ax.valSide : ax.catSide, cc.valueVertical ? valBand : catBand);
	if (!cc.valueVertical && ax.showValLabels && ax.valLabels.length > 0) {
		const half = (t) => t === void 0 ? 0 : estimateTextWidth(t, cc.valLook.px) / 2;
		const first = ax.valLabels[0];
		const last = ax.valLabels[ax.valLabels.length - 1];
		ins.l = Math.max(ins.l, half(ax.scale.reversed ? last : first));
		ins.r = Math.max(ins.r, half(ax.scale.reversed ? first : last));
	}
	if (cc.valueVertical && ax.showCatLabels && ax.catReserve && ax.catReach > 0) ins.l = Math.max(ins.l, ax.catReach - L.marginPt * PX_PER_PT);
	const tl = (lines, look) => lines.length === 0 ? 0 : lines.length * look.px * L.lineHeightRatio + L.axisTitleGapPx;
	ins.b += cc.valueVertical ? tl(cc.catAxisTitle, cc.catTitleLook) : tl(cc.valAxisTitle, cc.valTitleLook);
	ins.l += cc.valueVertical ? tl(cc.valAxisTitle, cc.valTitleLook) : tl(cc.catAxisTitle, cc.catTitleLook);
	return ins;
}
/** 自動配置の外側の矩形からプロットの内側（軸の文字・軸タイトルを引いた矩形）を解く */
function solvePlot(cx, cc, outer) {
	let plot = {
		x: outer.x,
		y: outer.y,
		w: outer.w,
		h: outer.h
	};
	let axes = solveAxes(cx, cc, plot.w, plot.h, outer.w);
	for (let k = 0; k < 2; k++) {
		const ins = axisInsets(cc, axes);
		plot = {
			x: outer.x + ins.l,
			y: outer.y + ins.t,
			w: Math.max(1, outer.w - ins.l - ins.r),
			h: Math.max(1, outer.h - ins.t - ins.b)
		};
		axes = solveAxes(cx, cc, plot.w, plot.h, outer.w);
	}
	return {
		plot,
		axes
	};
}
/** 矩形の経路 */
var rectD = (x0, y0, x1, y1) => `M${f(x0)} ${f(y0)}L${f(x1)} ${f(y0)}L${f(x1)} ${f(y1)}L${f(x0)} ${f(y1)}Z`;
/** 点列を直線（または Catmull-Rom を 3 次ベジェにした近似）でつなぐ経路 */
function polylineD(pts, smooth) {
	const p0 = pts[0];
	let d = `M${f(p0[0])} ${f(p0[1])}`;
	if (!smooth || pts.length < 3) {
		for (let i = 1; i < pts.length; i++) d += `L${f(pts[i][0])} ${f(pts[i][1])}`;
		return d;
	}
	for (let i = 0; i < pts.length - 1; i++) {
		const a = pts[Math.max(0, i - 1)];
		const b = pts[i];
		const c = pts[i + 1];
		const e = pts[Math.min(pts.length - 1, i + 2)];
		d += `C${f(b[0] + (c[0] - a[0]) / 6)} ${f(b[1] + (c[1] - a[1]) / 6)} ${f(c[0] - (e[0] - b[0]) / 6)} ${f(c[1] - (e[1] - b[1]) / 6)} ${f(c[0])} ${f(c[1])}`;
	}
	return d;
}
/** 直交座標のグラフを場面へ出す（プロットの面 → 目盛線 → 系列 → 軸線・目盛・文字 → 軸タイトル） */
function emitCartesian(cx, cc, outer, plot, ax) {
	const m = cx.model;
	const L = CHART_LAYOUT;
	const { scale } = ax;
	const vv = cc.valueVertical;
	const n = m.categories.length;
	const catLen = vv ? plot.w : plot.h;
	/** 項目の間隔が 1px 未満（点の多いグラフ）。系列の幾何を間引く（ADR-0031 §決定 11.1） */
	const dense = cx.thin && n > 0 && catLen / n < CHART_THIN.denseSpacingPx;
	const valAxis = m.valAxis;
	const catAxis = m.catAxis;
	const pFill = resolveFill(m.plotArea.fill, isDarkStyle(m) ? toColor(CHART_DEFAULTS.darkPlotFill) : null);
	const pStroke = resolveStroke(cx, m.plotArea.line, null);
	if (pFill !== null || pStroke !== null) push(cx, {
		kind: "rect",
		x: plot.x,
		y: plot.y,
		width: plot.w,
		height: plot.h,
		fill: pFill,
		stroke: pStroke
	});
	const valCoord = (v) => {
		const fr = valueFraction(scale, v);
		return vv ? plot.y + plot.h * (1 - fr) : plot.x + plot.w * fr;
	};
	const catCoord = (off) => vv ? plot.x + off : plot.y + plot.h - off;
	/** 範囲外を端に収めない数値軸の座標（折れ線をプロットの矩形で切るため） */
	const valCoordRaw = (v) => {
		const fr = valueFraction(scale, v, false);
		return vv ? plot.y + plot.h * (1 - fr) : plot.x + plot.w * fr;
	};
	const xy = (off, v) => vv ? [catCoord(off), valCoordRaw(v)] : [valCoordRaw(v), catCoord(off)];
	/** 数値軸に垂直な線（値 v の目盛線）・項目軸に垂直な線（位置 off） */
	const valLineAt = (v) => {
		const c = valCoord(v);
		return vv ? `M${f(plot.x)} ${f(c)}L${f(plot.x + plot.w)} ${f(c)}` : `M${f(c)} ${f(plot.y)}L${f(c)} ${f(plot.y + plot.h)}`;
	};
	const catLineAt = (off) => {
		const c = catCoord(off);
		return vv ? `M${f(c)} ${f(plot.y)}L${f(c)} ${f(plot.y + plot.h)}` : `M${f(plot.x)} ${f(c)}L${f(plot.x + plot.w)} ${f(c)}`;
	};
	const centers = categoryPositions(n, catLen, {
		midCat: ax.midCat,
		reversed: ax.catMaxMin
	});
	const bandCenters = categoryPositions(n, catLen, { reversed: ax.catMaxMin });
	const boundaries = [];
	if (ax.midCat) boundaries.push(...centers);
	else if (n + 1 <= 5e3) for (let i = 0; i <= n; i++) boundaries.push(i * catLen / Math.max(1, n));
	const grid = (line, spec, ds) => {
		if (line === void 0 || ds.length === 0) return;
		const stroke = resolveStroke(cx, line, {
			spec,
			widthPt: CHART_DEFAULTS.axisLinePt
		});
		if (stroke !== null) push(cx, {
			kind: "path",
			d: ds.join(""),
			fill: null,
			stroke
		});
	};
	grid(valAxis?.minorGridlines, CHART_DEFAULTS.gridMinor, scale.minorTicks.map(valLineAt));
	const gridMajor = lineSpec(cx, CHART_DEFAULTS.gridMajor);
	grid(valAxis?.majorGridlines, gridMajor, scale.majorTicks.map(valLineAt));
	grid(catAxis?.majorGridlines, gridMajor, boundaries.map(catLineAt));
	if (m.type === "bar") emitBars(cx, cc, ax, bandCenters, catLen, valCoord, catCoord, dense);
	else emitLines(cx, centers, xy, plot, dense ? { catAxis: vv ? 0 : 1 } : null);
	const crossC = valCoord(crossValue(scale, catAxis));
	const catStroke = catAxis?.deleted === true ? null : resolveStroke(cx, catAxis?.line, {
		spec: lineSpec(cx, CHART_DEFAULTS.catAxisLine),
		widthPt: CHART_DEFAULTS.axisLinePt
	});
	if (catStroke !== null) {
		if (vv) push(cx, {
			kind: "line",
			x1: plot.x,
			y1: crossC,
			x2: plot.x + plot.w,
			y2: crossC,
			stroke: catStroke
		});
		else push(cx, {
			kind: "line",
			x1: crossC,
			y1: plot.y,
			x2: crossC,
			y2: plot.y + plot.h,
			stroke: catStroke
		});
	}
	const valStroke = valAxis?.deleted === true ? null : resolveStroke(cx, valAxis?.line, {
		spec: lineSpec(cx, CHART_DEFAULTS.valAxisLine),
		widthPt: CHART_DEFAULTS.axisLinePt
	});
	const valLinePos = vv ? ax.valLineHigh ? plot.x + plot.w : plot.x : ax.valLineHigh ? plot.y : plot.y + plot.h;
	if (valStroke !== null) {
		if (vv) push(cx, {
			kind: "line",
			x1: valLinePos,
			y1: plot.y,
			x2: valLinePos,
			y2: plot.y + plot.h,
			stroke: valStroke
		});
		else push(cx, {
			kind: "line",
			x1: plot.x,
			y1: valLinePos,
			x2: plot.x + plot.w,
			y2: valLinePos,
			stroke: valStroke
		});
	}
	const tickMarks = (markOrNone, positions, horizontalAxis, linePos, labelsHigh, len, stroke) => {
		const mark = markOrNone ?? CHART_DEFAULTS.tickMark;
		if (stroke === null || mark === "none" || positions.length === 0) return;
		const out = horizontalAxis ? labelsHigh ? -1 : 1 : labelsHigh ? 1 : -1;
		const signs = mark === "out" ? [out] : mark === "in" ? [-out] : [out, -out];
		let d = "";
		for (const p of positions) for (const sg of signs) d += horizontalAxis ? `M${f(p)} ${f(linePos)}L${f(p)} ${f(linePos + sg * len)}` : `M${f(linePos)} ${f(p)}L${f(linePos + sg * len)} ${f(p)}`;
		push(cx, {
			kind: "path",
			d,
			fill: null,
			stroke
		});
	};
	const tickLen = L.tickLengthPt * PX_PER_PT;
	const minorLen = L.minorTickLengthPt * PX_PER_PT;
	const tickStroke = (axis, def) => axis?.deleted === true ? null : resolveStroke(cx, axis?.line, {
		spec: def,
		widthPt: CHART_DEFAULTS.axisLinePt
	});
	const valTickStroke = tickStroke(valAxis, lineSpec(cx, CHART_DEFAULTS.catAxisLine));
	const catTickStroke = tickStroke(catAxis, lineSpec(cx, CHART_DEFAULTS.catAxisLine));
	const catPosAbs = dense ? [] : boundaries.map(catCoord);
	tickMarks(catAxis?.majorTickMark, catPosAbs, vv, crossC, ax.catSide === "high", tickLen, catTickStroke);
	tickMarks(valAxis?.majorTickMark, scale.majorTicks.map(valCoord), !vv, valLinePos, ax.valSide === "high", tickLen, valTickStroke);
	tickMarks(valAxis?.minorTickMark, scale.minorTicks.map(valCoord), !vv, valLinePos, ax.valSide === "high", minorLen, valTickStroke);
	const gap = L.labelGapPx;
	if (ax.showValLabels) {
		const hi = ax.valSide === "high";
		ax.valLabels.forEach((t, i) => {
			const v = scale.majorTicks[i];
			if (vv) pushText(cx, hi ? plot.x + plot.w + tickLen + gap : plot.x - tickLen - gap, valCoord(v), t, cc.valLook, hi ? "start" : "end", "middle");
			else {
				const y = hi ? plot.y - tickLen - gap : plot.y + plot.h + tickLen + gap;
				pushText(cx, valCoord(v), y, t, cc.valLook, "middle", hi ? "alphabetic" : "hanging");
			}
		});
	}
	if (ax.showCatLabels) {
		const hi = ax.catSide === "high";
		const rot = ax.catRot;
		const edgeBase = vv ? hi ? plot.y : plot.y + plot.h : hi ? plot.x + plot.w : plot.x;
		const base = ax.catAtLine ? crossC : edgeBase;
		for (const i of ax.catIdx) {
			const off = centers[i];
			const t = ax.catLabels[i];
			if (vv) {
				const y = hi ? base - tickLen - gap : base + tickLen + gap;
				if (rot === 0) {
					const ls = ax.catLines[i] ?? [t];
					const pitch = cc.catLook.px * L.catWrapRowRatio;
					ls.forEach((line, k) => pushText(cx, catCoord(off), hi ? y - (ls.length - 1 - k) * pitch : y + k * pitch, line, cc.catLook, "middle", hi ? "alphabetic" : "hanging"));
				} else pushText(cx, catCoord(off), y, t, cc.catLook, rot < 0 !== hi ? "end" : "start", "middle", rot);
			} else pushText(cx, hi ? base + tickLen + gap : base - tickLen - gap, catCoord(off), t, cc.catLook, hi ? "start" : "end", "middle", rot);
		}
	}
	const emitAxisTitle = (axis, lines, horizontalAxis) => {
		if (lines.length === 0) return;
		const look = textLook(cx, axis?.title?.text, CHART_DEFAULTS.axisTitlePt);
		const lineH = look.px * L.lineHeightRatio;
		lines.forEach((t, k) => {
			if (horizontalAxis) {
				const y = outer.y + outer.h - (lines.length - k - .5) * lineH;
				pushText(cx, plot.x + plot.w / 2, y, t, look, "middle", "middle", axis?.title?.text.rotDeg ?? 0);
			} else pushText(cx, outer.x + (k + .5) * lineH, plot.y + plot.h / 2, t, look, "middle", "middle", axis?.title?.text.rotDeg ?? -90);
		});
	};
	emitAxisTitle(catAxis, cc.catAxisTitle, vv);
	emitAxisTitle(valAxis, cc.valAxisTitle, !vv);
}
/** 棒（集合・積み上げ・100% 積み上げ）。系列ごとに塗りの違いごとの経路にまとめる */
function emitBars(cx, cc, ax, centers, catLen, valCoord, catCoord, dense) {
	const m = cx.model;
	const vv = cc.valueVertical;
	const n = m.categories.length;
	const stacked = m.grouping === "stacked" || m.grouping === "percentStacked";
	const nBars = stacked ? 1 : m.series.length;
	const g = (m.gapWidth ?? CHART_DEFAULTS.gapWidth) / 100;
	const o = (m.overlap ?? (stacked ? CHART_DEFAULTS.overlapStacked : CHART_DEFAULTS.overlapClustered)) / 100;
	const band = catLen / Math.max(1, n);
	const denom = nBars - (nBars - 1) * o + g;
	const w = denom > 0 ? band / denom : band;
	const step = w * (1 - o);
	const groupW = w + (nBars - 1) * step;
	const base = crossValue(ax.scale, cx.model.catAxis);
	const pos = new Array(n).fill(0);
	const neg = new Array(n).fill(0);
	const totals = [];
	if (m.grouping === "percentStacked") for (let i = 0; i < n; i++) totals.push(m.series.reduce((a, s) => a + (typeof s.values[i] === "number" && Number.isFinite(s.values[i]) ? Math.abs(s.values[i]) : 0), 0));
	const vary = varyByPoint(m);
	const barColPx = !stacked && o < 1 ? clusteredBarColumnPx(m.series.length) : CHART_THIN.barColumnPx;
	m.series.forEach((s, j) => {
		const groups = /* @__PURE__ */ new Map();
		for (let i = 0; i < n; i++) {
			const raw = s.values[i];
			if (typeof raw !== "number" || !Number.isFinite(raw)) continue;
			let v0;
			let v1;
			if (stacked) {
				const v = m.grouping === "percentStacked" ? totals[i] > 0 ? raw / totals[i] : 0 : raw;
				if (v >= 0) {
					v0 = pos[i];
					v1 = v0 + v;
					pos[i] = v1;
				} else {
					v0 = neg[i];
					v1 = v0 + v;
					neg[i] = v1;
				}
			} else {
				v0 = base;
				v1 = raw;
			}
			const slot = stacked ? 0 : ax.catMaxMin ? nBars - 1 - j : j;
			const lo = centers[i] - groupW / 2 + slot * step;
			const hi = lo + w;
			const a = valCoord(v0);
			const b = valCoord(v1);
			if (Math.max(v0, v1) < ax.scale.min || Math.min(v0, v1) > ax.scale.max) continue;
			if (Math.abs(a - b) < 1e-9) continue;
			const c1 = catCoord(lo);
			const c2 = catCoord(hi);
			const fill = pointFill(cx, s, i, vary);
			const key = fill === null ? "none" : `${fill.color}/${fill.opacity}`;
			let grp = groups.get(key);
			if (grp === void 0) {
				grp = {
					fill,
					ds: [],
					boxes: []
				};
				groups.set(key, grp);
			}
			if (dense) grp.boxes.push({
				lo: Math.min(c1, c2),
				hi: Math.max(c1, c2),
				a: Math.min(a, b),
				b: Math.max(a, b),
				neg: v1 < v0
			});
			else grp.ds.push(vv ? rectD(Math.min(c1, c2), Math.min(a, b), Math.max(c1, c2), Math.max(a, b)) : rectD(Math.min(a, b), Math.min(c1, c2), Math.max(a, b), Math.max(c1, c2)));
		}
		const stroke = resolveStroke(cx, s.line, null);
		for (const grp of groups.values()) {
			if (grp.fill === null && stroke === null) continue;
			push(cx, {
				kind: "path",
				d: dense ? mergeBarColumns(grp.boxes, barColPx).map((bx) => vv ? rectD(bx.lo, bx.a, bx.hi, bx.b) : rectD(bx.a, bx.lo, bx.b, bx.hi)).join("") : grp.ds.join(""),
				fill: grp.fill,
				stroke
			});
		}
	});
}
/** 線分 p→q を矩形（境界を含む）で切る（Liang–Barsky）。矩形の外なら null。t0・t1 は元の線分上の位置（0〜1） */
function clipSegment(p, q, r) {
	const eps = 1e-6;
	const dx = q[0] - p[0];
	const dy = q[1] - p[1];
	let t0 = 0;
	let t1 = 1;
	const edges = [
		[-dx, p[0] - (r.x - eps)],
		[dx, r.x + r.w + eps - p[0]],
		[-dy, p[1] - (r.y - eps)],
		[dy, r.y + r.h + eps - p[1]]
	];
	for (const [pp, qq] of edges) {
		if (pp === 0) {
			if (qq < 0) return null;
			continue;
		}
		const t = qq / pp;
		if (pp < 0) {
			if (t > t1) return null;
			if (t > t0) t0 = t;
		} else {
			if (t < t0) return null;
			if (t < t1) t1 = t;
		}
	}
	return {
		a: [p[0] + t0 * dx, p[1] + t0 * dy],
		b: [p[0] + t1 * dx, p[1] + t1 * dy],
		t0,
		t1
	};
}
/** 点列の折れ線を矩形で切る。矩形の外を通る区間で分かれ、2 点以上の部分だけを返す（端に寄せず、傾きを保つ） */
function clipPolyline(pts, r) {
	const out = [];
	let cur = [];
	const flush = () => {
		const p0 = cur[0];
		if (p0 !== void 0 && cur.some((p) => p[0] !== p0[0] || p[1] !== p0[1])) out.push(cur);
		cur = [];
	};
	/** 直前の線分が終点まで矩形の中に入っていた（＝次の線分は続きとしてつなぐ） */
	let open = false;
	for (let i = 0; i + 1 < pts.length; i++) {
		const c = clipSegment(pts[i], pts[i + 1], r);
		if (c === null) {
			flush();
			open = false;
			continue;
		}
		if (!(open && c.t0 === 0)) {
			flush();
			cur = [c.a];
		}
		cur.push(c.b);
		open = c.t1 === 1;
	}
	flush();
	return out;
}
/** 点が矩形の中（境界を含む）か */
function inRect(p, r) {
	const eps = 1e-6;
	return p[0] >= r.x - eps && p[0] <= r.x + r.w + eps && p[1] >= r.y - eps && p[1] <= r.y + r.h + eps;
}
/** 滑らかな線（Catmull-Rom）を、切るために折れ線に直す（区間ごとに `steps` 分割） */
function flattenSmooth(pts, steps) {
	const out = [[pts[0][0], pts[0][1]]];
	for (let i = 0; i < pts.length - 1; i++) {
		const a = pts[Math.max(0, i - 1)];
		const b = pts[i];
		const c = pts[i + 1];
		const e = pts[Math.min(pts.length - 1, i + 2)];
		const c1 = [b[0] + (c[0] - a[0]) / 6, b[1] + (c[1] - a[1]) / 6];
		const c2 = [c[0] - (e[0] - b[0]) / 6, c[1] - (e[1] - b[1]) / 6];
		for (let k = 1; k <= steps; k++) {
			const t = k / steps;
			const u = 1 - t;
			const w0 = u * u * u;
			const w1 = 3 * u * u * t;
			const w2 = 3 * u * t * t;
			const w3 = t * t * t;
			out.push([w0 * b[0] + w1 * c1[0] + w2 * c2[0] + w3 * c[0], w0 * b[1] + w1 * c1[1] + w2 * c2[1] + w3 * c[1]]);
		}
	}
	return out;
}
/**
* 折れ線（空の点 3 種・積み上げの累積・滑らか・マーカー）。
* `xy` は範囲外を端に収めない座標。軸の範囲の外へ出る区間は `clip`（プロットの矩形）で幾何的に切り（傾きを保つ）、
* 外を通る区間で線は複数の部分に分かれる。プロットの外の点のマーカーは出さない。
* `thin` が非 null なら点の多いグラフ：線は部分ごとに M4→Douglas–Peucker で間引き（滑らかな線は 4 分割の折れ線にして `C` を出さない）、
* マーカーは既定の見た目の円を帯に・他を間引く（`c:dPt` の点・`varyColors` の点ごとの色は間引かない。ADR-0031 §決定 11.1）。
*/
function emitLines(cx, centers, xy, clip, thin) {
	const m = cx.model;
	const n = m.categories.length;
	const stacked = m.grouping === "stacked" || m.grouping === "percentStacked";
	const cum = new Array(n).fill(0);
	const totals = [];
	if (m.grouping === "percentStacked") for (let i = 0; i < n; i++) totals.push(m.series.reduce((a, s) => a + (typeof s.values[i] === "number" && Number.isFinite(s.values[i]) ? Math.abs(s.values[i]) : 0), 0));
	const vary = varyByPoint(m);
	m.series.forEach((s) => {
		const pts = [];
		for (let i = 0; i < n; i++) {
			const raw = s.values[i];
			const ok = typeof raw === "number" && Number.isFinite(raw);
			if (!ok && m.dispBlanksAs !== "zero") {
				pts.push(null);
				continue;
			}
			let v = ok ? raw : 0;
			if (m.grouping === "percentStacked") v = totals[i] > 0 ? v / totals[i] : 0;
			if (stacked) {
				cum[i] = cum[i] + v;
				v = cum[i];
			}
			pts.push(xy(centers[i], v));
		}
		const segs = [];
		let cur = [];
		for (const p of pts) {
			if (p === null) {
				if (m.dispBlanksAs === "gap" && cur.length > 0) {
					segs.push(cur);
					cur = [];
				}
				continue;
			}
			cur.push(p);
		}
		if (cur.length > 0) segs.push(cur);
		const lineSpecColor = autoColor(cx.theme, s.idx);
		const stroke = resolveStroke(cx, s.line, {
			spec: lineSpecColor,
			widthPt: CHART_DEFAULTS.seriesLinePt
		}, {
			cap: "round",
			join: "round"
		});
		if (stroke !== null) {
			const d = segs.filter((sg) => sg.length >= 2).map((sg) => {
				const smooth = s.smooth === true && sg.length >= 3;
				const inside = sg.every((p) => inRect(p, clip));
				if (thin !== null) {
					const flat = smooth ? flattenSmooth(sg, CHART_THIN.smoothSteps) : sg;
					return (inside ? [flat] : clipPolyline(flat, clip)).map((piece) => polylineD(thinPolyline(piece, thin.catAxis), false)).join("");
				}
				if (inside) return polylineD(sg, smooth);
				return clipPolyline(smooth ? flattenSmooth(sg, 16) : sg, clip).map((piece) => polylineD(piece, false)).join("");
			}).join("");
			if (d !== "") push(cx, {
				kind: "path",
				d,
				fill: null,
				stroke
			});
		}
		const sym = markerSymbol(s);
		if (sym === null) return;
		const r = (s.marker?.size ?? CHART_DEFAULTS.markerSizePt) * PX_PER_PT / 2;
		const base = s.line?.paint?.kind === "solid" ? toColor(s.line.paint.color) : seriesAuto(cx, s);
		const filled = markerPath(sym, 0, 0, r).filled;
		const groups = /* @__PURE__ */ new Map();
		pts.forEach((p, i) => {
			if (p === null || !inRect(p, clip)) return;
			const dpt = s.points?.get(i);
			let fill;
			if (dpt !== void 0) fill = resolveFill(dpt, null);
			else if (s.marker?.fill !== void 0) fill = resolveFill(s.marker.fill, null);
			else fill = vary ? toColor(autoColor(cx.theme, i)) : base;
			const key = fill === null ? "none" : `${fill.color}/${fill.opacity}`;
			let grp = groups.get(key);
			if (grp === void 0) {
				const ms = s.marker?.line !== void 0 ? resolveStroke(cx, s.marker.line, fill === null ? null : {
					spec: {
						rgb: parseInt(fill.color.slice(1), 16),
						alpha: fill.opacity
					},
					widthPt: CHART_DEFAULTS.thinLinePt
				}) : fill === null ? null : strokeFrom(fill);
				grp = {
					fill: filled ? fill : null,
					stroke: ms,
					d: "",
					centers: []
				};
				groups.set(key, grp);
			}
			if (thin === null || dpt !== void 0 || vary) grp.d += markerPath(sym, p[0], p[1], r).d;
			else grp.centers.push(p);
		});
		for (const grp of groups.values()) {
			if (grp.fill === null && grp.stroke === null) continue;
			let band = null;
			if (thin !== null && grp.centers.length > 0) {
				const st = grp.stroke;
				const fl = grp.fill;
				if (markerPath(sym, 0, 0, r).d.includes("A") && fl !== null && st !== null && fl.opacity === 1 && st.opacity === 1 && st.color === fl.color && st.dash === "solid") {
					const R = r + st.width / 2;
					const { runs, singles } = splitMarkerBands(grp.centers, bandGap(R));
					for (const p of singles) grp.d += markerPath(sym, p[0], p[1], r).d;
					const bd = runs.map((run) => polylineD(thinPolyline(run, thin.catAxis), false)).join("");
					if (bd !== "") band = {
						kind: "path",
						d: bd,
						fill: null,
						stroke: {
							color: fl.color,
							opacity: 1,
							width: 2 * R,
							dash: "solid",
							cap: "round",
							join: "round"
						}
					};
				} else for (const i of thinMarkerIndices(grp.centers)) {
					const p = grp.centers[i];
					grp.d += markerPath(sym, p[0], p[1], r).d;
				}
			}
			if (grp.d !== "") push(cx, {
				kind: "path",
				d: grp.d,
				fill: grp.fill,
				stroke: grp.stroke
			});
			if (band !== null) push(cx, band);
		}
	});
}
/**
* グラフの場面を計算する（ISSUE-102・ADR-0031 §決定 3・6）。
* @param model `parseChart` の結果のグラフ
* @param widthPx グラフの面の幅（px）
* @param heightPx グラフの面の高さ（px）
* @param theme テーマの色（既定の色に使う）
* @param opts `thin`＝点の多いグラフ（項目の間隔 1px 未満）の幾何を間引くか（既定 真。ADR-0031 §決定 11.1。偽は照合・試験で間引きの有無を比べる用）
* @returns 場面。大きさが正の有限値でない・場面の要素が {@link MAX_CHART_SCENE_ITEMS} を超えるなら null（表示側は破線枠）
*/
function layoutChart(model, widthPx, heightPx, theme = DEFAULT_THEME, opts) {
	if (!(Number.isFinite(widthPx) && widthPx > 0 && Number.isFinite(heightPx) && heightPx > 0)) return null;
	const pie = model.type === "pie" || model.type === "doughnut";
	if (pie && (model.type === "pie" ? 1 : model.series.length) * model.categories.length > 5e3) return null;
	const radius = model.roundedCorners ?? CHART_DEFAULTS.roundedCorners ? Math.max(0, Math.min(CHART_DEFAULTS.cornerRadiusPt * PX_PER_PT, widthPx / 2, heightPx / 2)) : 0;
	const cx = {
		model,
		theme,
		W: widthPx,
		H: heightPx,
		items: [],
		thin: opts?.thin ?? true,
		radius
	};
	try {
		build(cx, pie);
	} catch (e) {
		if (e instanceof SceneOverflow) return null;
		throw e;
	}
	return {
		width: widthPx,
		height: heightPx,
		backdrop: model.chartArea.fill?.kind === "none",
		cornerRadius: radius,
		items: cx.items
	};
}
/** 場面を組み立てる（グラフの面 → プロット・系列 → 凡例 → タイトルの順に出す） */
function build(cx, pie) {
	const m = cx.model;
	const L = CHART_LAYOUT;
	const { W, H } = cx;
	const dark = isDarkStyle(m);
	const caFill = resolveFill(m.chartArea.fill, toColor(dark ? CHART_DEFAULTS.darkChartAreaFill : specColor(cx.theme, CHART_DEFAULTS.chartAreaFill)));
	const caStroke = resolveStroke(cx, m.chartArea.line, {
		spec: dark ? CHART_DEFAULTS.darkChartAreaLine : CHART_DEFAULTS.chartAreaLine,
		widthPt: CHART_DEFAULTS.chartAreaLinePt
	});
	if (caFill !== null || caStroke !== null) {
		const rect = {
			kind: "rect",
			x: 0,
			y: 0,
			width: W,
			height: H,
			fill: caFill,
			stroke: caStroke
		};
		if (cx.radius > 0) rect.radius = cx.radius;
		push(cx, rect);
	}
	const margin = L.marginPt * PX_PER_PT;
	const pad = L.plotPadPt * PX_PER_PT;
	const area = {
		x: 0,
		y: 0,
		w: W,
		h: H
	};
	const blankTitle = isBlankAutoTitle(m);
	const tLines = blankTitle ? [""] : titleLines(m);
	let emitTitle = null;
	if (tLines.length > 0) {
		const look = textLook(cx, m.title?.text, m.c14Style !== void 0 ? CHART_DEFAULTS.c14TitlePt : CHART_DEFAULTS.titlePt, CHART_DEFAULTS.titleBold);
		const lineH = look.px * L.titleLineHeightRatio;
		const tw = tLines.reduce((a, t) => Math.max(a, estimateTextWidth(t, look.px)), 0);
		const th = tLines.length * lineH;
		const auto = {
			x: (W - tw) / 2,
			y: L.titleTopPt * PX_PER_PT,
			w: tw,
			h: th
		};
		const manual = hasManual(m.title?.layout);
		const box = manual ? applyManual(m.title.layout, W, H, auto) : auto;
		if (!manual && m.title?.overlay !== true) {
			area.y = auto.y + th;
			area.h = Math.max(1, H - area.y);
		}
		if (!blankTitle) emitTitle = () => {
			tLines.forEach((t, k) => pushText(cx, box.x + tw / 2, box.y + (k + .5) * lineH, t, look, "middle", "middle", m.title?.text.rotDeg ?? 0));
		};
	}
	let emitLeg = null;
	const legend = m.legend;
	if (legend !== void 0) {
		const look = textLook(cx, legend.text, CHART_DEFAULTS.legendPt);
		const entries = legendEntries(cx, legend);
		const vertical = legend.pos === "r" || legend.pos === "l" || legend.pos === "tr";
		const gap = L.legendEdgeGapPt * PX_PER_PT;
		const maxW = vertical ? area.w * L.legendMaxFraction : area.w;
		const maxH = vertical ? Math.max(0, area.h - 2 * gap) : area.h * L.legendMaxFraction;
		let box = buildLegend(entries, legend.pos, look, maxW, maxH);
		if (box !== null) {
			let ox;
			let oy;
			switch (legend.pos) {
				case "l":
					ox = gap;
					oy = area.y + (area.h - box.h) / 2;
					break;
				case "t":
					ox = area.x + (area.w - box.w) / 2;
					oy = area.y + gap;
					break;
				case "b":
					ox = area.x + (area.w - box.w) / 2;
					oy = H - gap - box.h;
					break;
				case "tr":
					ox = W - gap - box.w;
					oy = area.y + gap;
					break;
				default:
					ox = W - gap - box.w;
					oy = area.y + (area.h - box.h) / 2;
			}
			if (hasManual(legend.layout)) {
				const rect = applyManual(legend.layout, W, H, {
					x: ox,
					y: oy,
					w: box.w,
					h: box.h
				});
				if (legend.layout.w !== void 0 || legend.layout.h !== void 0) box = buildLegend(entries, legend.pos, look, rect.w, rect.h) ?? box;
				ox = rect.x;
				oy = rect.y;
			} else if (!legend.overlay) {
				if (legend.pos === "l") {
					area.w = Math.max(1, area.x + area.w - (ox + box.w));
					area.x = ox + box.w;
				} else if (legend.pos === "t") {
					area.h = Math.max(1, area.y + area.h - (oy + box.h));
					area.y = oy + box.h;
				} else if (legend.pos === "b") area.h = Math.max(1, oy - area.y);
				else {
					area.w = Math.max(1, ox - area.x);
					if (legend.pos === "tr" && !pie) {
						const bottom = area.y + area.h;
						const below = oy + box.h - gap;
						if ((bottom - below - pad - margin) / Math.max(1, area.h - pad - margin) >= L.legendCornerShrinkMin) {
							area.h = Math.max(1, bottom - below);
							area.y = below;
						} else {
							oy += gap;
							area.y += gap;
							area.h = Math.max(1, area.h - gap);
						}
					}
				}
			}
			const done = box;
			emitLeg = () => emitLegend(cx, done, look, ox, oy);
		}
	}
	const pl = m.plotLayout;
	if (pie) {
		const free = {
			x: area.x + pad,
			y: area.y + pad,
			w: Math.max(1, area.w - 2 * pad),
			h: Math.max(1, area.h - 2 * pad)
		};
		emitPieArea(cx, hasManual(pl) ? applyManual(pl, W, H, free) : free);
	} else {
		const cc = cartCtx(cx);
		const free = {
			x: area.x + margin,
			y: area.y + pad,
			w: Math.max(1, area.w - margin - pad),
			h: Math.max(1, area.h - pad - margin)
		};
		const auto = solvePlot(cx, cc, free);
		let outer = free;
		let plot = auto.plot;
		let axes = auto.axes;
		if (hasManual(pl)) {
			if (pl.target === "inner") {
				plot = applyManual(pl, W, H, auto.plot);
				axes = solveAxes(cx, cc, plot.w, plot.h);
			} else {
				outer = applyManual(pl, W, H, free);
				const s = solvePlot(cx, cc, outer);
				plot = s.plot;
				axes = s.axes;
			}
		}
		emitCartesian(cx, cc, outer, plot, axes);
	}
	emitLeg?.();
	emitTitle?.();
}
/** 円・ドーナツ：プロットの面の矩形（塗りがあれば）と扇 */
function emitPieArea(cx, plot) {
	const m = cx.model;
	const pFill = resolveFill(m.plotArea.fill, null);
	const pStroke = resolveStroke(cx, m.plotArea.line, null);
	if (pFill !== null || pStroke !== null) push(cx, {
		kind: "rect",
		x: plot.x,
		y: plot.y,
		width: plot.w,
		height: plot.h,
		fill: pFill,
		stroke: pStroke
	});
	emitPie(cx, plot);
}
// src/render/chart-svg.ts
var SVG_NS = "http://www.w3.org/2000/svg";
/** 文字の揃え → `text-anchor`（固定の語の表） */
var TEXT_ANCHOR = {
	start: "start",
	middle: "middle",
	end: "end"
};
/**
* 文字の縦位置 → `dominant-baseline`（固定の語の表）。場面の `middle` は行の中央に置く意味なので、em の箱の中央（`central`）に写す。
*/
var TEXT_BASELINE = {
	alphabetic: "alphabetic",
	middle: "central",
	hanging: "hanging"
};
/** 線の端・結びの固定の語（場面の値を通す表。表に無い値は入れない） */
var LINE_CAP = {
	butt: "butt",
	round: "round"
};
var LINE_JOIN = {
	miter: "miter",
	round: "round"
};
/** `d` に許す文字（数値と固定の命令 `M L C A Z`・区切り・符号・小数点・指数） */
var PATH_DATA = /^[MLCAZ0-9\s.,+\-eE]*$/;
/** `#rrggbb` */
var HEX_COLOR = /^#[0-9a-fA-F]{6}$/;
/** 有限でない値を 0 にして小数 3 桁へ丸める（属性に入れる数値用） */
function num(n) {
	return Number.isFinite(n) ? Math.round(n * 1e3) / 1e3 : 0;
}
/** 色が `#rrggbb` で、不透明度が有限か */
function validColor(c) {
	return !!c && HEX_COLOR.test(c.color) && Number.isFinite(c.opacity);
}
/** 不透明度を 0〜1・小数 3 桁に収める */
function alpha(a) {
	return Math.round(Math.min(1, Math.max(0, a)) * 1e3) / 1e3;
}
/** SVG の要素を作り、属性（数値・固定の語・検証済みの値だけ）を入れる */
function el(doc, tag, attrs) {
	const e = doc.createElementNS(SVG_NS, tag);
	for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
	return e;
}
/** 塗りの属性（`fill`・`fill-opacity`）。null・不正な色は `fill="none"` */
function fillAttrs(fill) {
	if (!validColor(fill)) return { fill: "none" };
	const a = alpha(fill.opacity);
	return a < 1 ? {
		fill: fill.color,
		"fill-opacity": a
	} : { fill: fill.color };
}
/** 線の属性。null・不正な色は `stroke="none"` */
function strokeAttrs(stroke) {
	if (!stroke || !validColor(stroke)) return { stroke: "none" };
	const w = num(stroke.width);
	const out = {
		stroke: stroke.color,
		"stroke-width": w,
		"stroke-linecap": LINE_CAP[stroke.cap] ?? "butt",
		"stroke-linejoin": LINE_JOIN[stroke.join] ?? "miter"
	};
	const a = alpha(stroke.opacity);
	if (a < 1) out["stroke-opacity"] = a;
	const dash = dashArray(stroke.dash, w).filter((n) => Number.isFinite(n) && n >= 0);
	if (dash.length > 0) out["stroke-dasharray"] = dash.map(num).join(" ");
	return out;
}
/** 項目 1 つを要素にする（不正なら null） */
function itemEl(doc, item) {
	switch (item.kind) {
		case "rect": return el(doc, "rect", {
			x: num(item.x),
			y: num(item.y),
			width: num(item.width),
			height: num(item.height),
			...item.radius !== void 0 && Number.isFinite(item.radius) && item.radius > 0 ? {
				rx: num(item.radius),
				ry: num(item.radius)
			} : {},
			...fillAttrs(item.fill),
			...strokeAttrs(item.stroke)
		});
		case "path":
			if (item.d === "" || !PATH_DATA.test(item.d)) return null;
			return el(doc, "path", {
				d: item.d,
				...fillAttrs(item.fill),
				...strokeAttrs(item.stroke)
			});
		case "line": return el(doc, "line", {
			x1: num(item.x1),
			y1: num(item.y1),
			x2: num(item.x2),
			y2: num(item.y2),
			fill: "none",
			...strokeAttrs(item.stroke)
		});
		case "text": {
			if (!validColor(item.color)) return null;
			const x = num(item.x);
			const y = num(item.y);
			const attrs = {
				x,
				y,
				"font-size": num(item.size),
				"text-anchor": TEXT_ANCHOR[item.anchor] ?? "start",
				"dominant-baseline": TEXT_BASELINE[item.baseline] ?? "alphabetic",
				...fillAttrs(item.color)
			};
			if (item.bold) attrs["font-weight"] = "bold";
			const rot = num(item.rotate);
			if (rot !== 0) attrs.transform = `rotate(${rot} ${x} ${y})`;
			const t = el(doc, "text", attrs);
			t.textContent = item.text;
			return t;
		}
		default: return null;
	}
}
/**
* グラフの場面を `<svg>` にする。大きさ（`width`・`height`・`viewBox`）は場面の px。クラス `xs-chart`（CSS が枠いっぱいに伸ばす）。
* 下地（`scene.backdrop`）はこの `<svg>` には含めない（呼び出し側の画像レイヤーが図の層の最背面に置く）。
*
* @param scene `layoutChart` の結果（null は呼び出し側が破線枠にする）
* @param doc 要素を作る文書
* @param title 読み上げの名前（呼び出し側の i18n の文言。`role="img"` の `<title>` に `textContent` で入れる。省略すると `aria-hidden`）
*/
function buildChartSvg(scene, doc, title) {
	const w = num(scene.width);
	const h = num(scene.height);
	const svg = el(doc, "svg", {
		width: w,
		height: h,
		viewBox: `0 0 ${w} ${h}`,
		focusable: "false",
		...title === void 0 ? { "aria-hidden": "true" } : { role: "img" }
	});
	svg.setAttribute("class", "xs-chart");
	if (title !== void 0) {
		const t = doc.createElementNS(SVG_NS, "title");
		t.textContent = title;
		svg.append(t);
	}
	for (const item of scene.items) {
		const e = itemEl(doc, item);
		if (e) svg.append(e);
	}
	return svg;
}
export { buildChartSvg, layoutChart };
