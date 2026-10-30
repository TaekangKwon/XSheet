// 正: InterHub docs/plan/adr/0031-xlsx-charts.md（§決定 11.1）
/**
 * `.xlsx` のグラフの幾何の間引き（ISSUE-106・ADR-0031 §決定 11.1）：点の間隔が 1px 未満の「密な」グラフで、
 * 画素より細かい幾何を、見た目がほぼ変わらない範囲で減らす純関数の集まり。
 * - 単位はグラフの面の CSS px（場面の px）。`devicePixelRatio` は読まない（場面は入力の純関数のまま）。
 * - 折れ線：M4（0.5px の列ごとの最初・最小・最大・最後）→ Douglas–Peucker（許容 0.25px）。
 * - マーカー：円は太い線の「帯」に、他は Douglas–Peucker の点を残して 1px 以内を省く。
 * - 棒：1px の列ごとに 1 つの矩形にまとめる（幅は合計・値の範囲は和集合）。集合の棒は列の幅を系列の数に合わせて広げる（最大 2px）。
 * - 円・ドーナツ：外周の弧が線の太さ以下の扇の連なりを、線の色で塗った扇 1 つへ束ねる（ISSUE-111）。
 * - DOM・canvas・`chrome.*` に触らない。
 */

/** 点（x, y） */
export type Pt = readonly [number, number];

/** 間引きの定数（単位は CSS px・`smoothSteps` だけ分割数） */
export const CHART_THIN = {
  /** 項目の間隔がこれ未満のグラフを「密」とみなす */
  denseSpacingPx: 1,
  /** M4 の列の幅 */
  gridPx: 0.5,
  /** Douglas–Peucker の許容 */
  epsPx: 0.25,
  /** その他のマーカーで、直前に残した点からこの距離以内なら省く */
  markerSpacingPx: 1,
  /** 棒をまとめる列の幅 */
  barColumnPx: 1,
  /** 集合の棒で、系列 1 つにつき列の幅へ足す量（列の幅＝系列の数×これを整数 px へ丸めた値。1px 以上・下の上限以下） */
  barColumnPerSeriesPx: 0.2,
  /** 集合の棒の列の幅の上限（これより広げると列が縞に見えて元の絵から離れる。ISSUE-111 の計測） */
  barColumnMaxPx: 2,
  /** 滑らかな線を折れ線に直す時の区間ごとの分割数 */
  smoothSteps: 4,
} as const;

/**
 * M4 の間引き：項目軸の方向に `gridPx` の幅で並ぶ列ごとに、同じ列に続く点の最初・最小・最大・最後だけを元の順で残す。
 * 続く点を列とするので、項目の並びが逆順（右から左・上から下）でも動く。
 * @param pts 点列
 * @param catAxis 項目軸の方向（0＝x・1＝y）。最小・最大は他方の座標で決める
 * @param gridPx 列の幅
 * @returns 点が 3 未満なら入力の写し。それ以外は間引いた点列（最初と最後の点は必ず残る）
 */
export function thinM4(
  pts: readonly Pt[],
  catAxis: 0 | 1,
  gridPx: number = CHART_THIN.gridPx,
): Pt[] {
  const n = pts.length;
  if (n < 3) return pts.slice();
  const v = catAxis === 0 ? 1 : 0;
  const out: Pt[] = [];
  let i = 0;
  while (i < n) {
    const col = Math.floor(pts[i]![catAxis] / gridPx);
    let j = i + 1;
    while (j < n && Math.floor(pts[j]![catAxis] / gridPx) === col) j++;
    if (j - i <= 4) {
      for (let k = i; k < j; k++) out.push(pts[k]!);
    } else {
      let lo = i;
      let hi = i;
      for (let k = i + 1; k < j; k++) {
        if (pts[k]![v] < pts[lo]![v]) lo = k;
        if (pts[k]![v] > pts[hi]![v]) hi = k;
      }
      const keep = [...new Set([i, lo, hi, j - 1])].sort((a, b) => a - b);
      for (const k of keep) out.push(pts[k]!);
    }
    i = j;
  }
  return out;
}

/** 点 p から線分 ab までの距離 */
function distToSegment(p: Pt, a: Pt, b: Pt): number {
  const dx = b[0] - a[0];
  const dy = b[1] - a[1];
  const len2 = dx * dx + dy * dy;
  let t = len2 === 0 ? 0 : ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / len2;
  t = t < 0 ? 0 : t > 1 ? 1 : t;
  return Math.hypot(p[0] - (a[0] + t * dx), p[1] - (a[1] + t * dy));
}

/** Douglas–Peucker の距離の計算回数の上限（これを超えたら間引きをやめる。通常の入力では届かない） */
const DP_WORK_BUDGET = 20_000_000;

/**
 * Douglas–Peucker で残す点の添字（昇順）。反復で解く（再帰は深くなるため使わない）。端点は必ず残す。
 * 残さなかった点は、残した点を結ぶ線分から `eps` 以内にある。
 * 最悪の入力（鋸歯など）では 2 乗の時間になるため、距離の計算が {@link DP_WORK_BUDGET} 回を超えたら、
 * 残りの区間は全点を残す（点が増えるだけで、ずれの上限は守られる）。
 */
export function simplifyIndices(pts: readonly Pt[], eps: number = CHART_THIN.epsPx): number[] {
  const n = pts.length;
  if (n < 3) return pts.map((_, i) => i);
  const keep = new Uint8Array(n);
  keep[0] = 1;
  keep[n - 1] = 1;
  const stack: [number, number][] = [[0, n - 1]];
  let work = 0;
  while (stack.length > 0) {
    const [lo, hi] = stack.pop()!;
    if (hi - lo < 2) continue;
    work += hi - lo;
    if (work > DP_WORK_BUDGET) {
      for (let k = lo; k <= hi; k++) keep[k] = 1;
      continue;
    }
    const a = pts[lo]!;
    const b = pts[hi]!;
    let worst = -1;
    let at = -1;
    for (let k = lo + 1; k < hi; k++) {
      const d = distToSegment(pts[k]!, a, b);
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
  const out: number[] = [];
  for (let i = 0; i < n; i++) if (keep[i] === 1) out.push(i);
  return out;
}

/** Douglas–Peucker（端点を残す・許容 `eps`）で単純化した点列。点が 3 未満なら入力の写し */
export function simplifyDP(pts: readonly Pt[], eps: number = CHART_THIN.epsPx): Pt[] {
  if (pts.length < 3) return pts.slice();
  return simplifyIndices(pts, eps).map((i) => pts[i]!);
}

/**
 * 折れ線の間引き：M4 → Douglas–Peucker。中心線のずれは 0.25px 以下（M4 の列の中は正確）。
 * @param pts 線の 1 つの部分（空の点・プロットの外で切れた後の 1 本）の点列
 * @param catAxis 項目軸の方向（0＝x・1＝y）
 */
export function thinPolyline(pts: readonly Pt[], catAxis: 0 | 1): Pt[] {
  return simplifyDP(thinM4(pts, catAxis));
}

/**
 * 円のマーカーの帯の区間分け。続く中心の間隔が `maxGap` 以下の区間（2 点以上）を「帯」にする。
 * 帯は呼び出し側が間引いた中心線を、太さ 2R・端と結びが round の線 1 本で描く。区間に入らない点は単独のマーカー。
 * @param centers マーカーの中心（元の順）
 * @param maxGap 続く中心の間隔の上限。{@link bandGap} で求める
 * @returns `runs`＝帯の区間（各 2 点以上・元の順）／`singles`＝帯に入らない点（元の順）
 */
export function splitMarkerBands(
  centers: readonly Pt[],
  maxGap: number,
): { runs: Pt[][]; singles: Pt[] } {
  const runs: Pt[][] = [];
  const singles: Pt[] = [];
  let cur: Pt[] = [];
  const flush = (): void => {
    if (cur.length >= 2) runs.push(cur);
    else if (cur.length === 1) singles.push(cur[0]!);
    cur = [];
  };
  for (const p of centers) {
    const last = cur[cur.length - 1];
    if (last !== undefined && Math.hypot(p[0] - last[0], p[1] - last[1]) > maxGap) flush();
    cur.push(p);
  }
  flush();
  return { runs, singles };
}

/**
 * 円のマーカー（外側の半径 R）の帯にできる、続く中心の間隔の上限 √(8R·eps)。
 * 円の連なりの縁の波 d²/8R がこの間隔で eps 以下になる。
 */
export function bandGap(R: number, eps: number = CHART_THIN.epsPx): number {
  return Math.sqrt(8 * Math.max(0, R) * eps);
}

/**
 * その他のマーカーの間引きで残す点の添字（昇順）。Douglas–Peucker（許容 `eps`）で選んだ点は残し、
 * その間は直前に残した点から `spacing` 以内の点を省く（省いた点は残した点から `spacing` 以内）。
 */
export function thinMarkerIndices(
  centers: readonly Pt[],
  eps: number = CHART_THIN.epsPx,
  spacing: number = CHART_THIN.markerSpacingPx,
): number[] {
  const n = centers.length;
  if (n < 3) return centers.map((_, i) => i);
  const dp = new Uint8Array(n);
  for (const i of simplifyIndices(centers, eps)) dp[i] = 1;
  const out: number[] = [0];
  let last = centers[0]!;
  for (let i = 1; i < n; i++) {
    const p = centers[i]!;
    if (dp[i] === 1 || Math.hypot(p[0] - last[0], p[1] - last[1]) > spacing) {
      out.push(i);
      last = p;
    }
  }
  return out;
}

/** 棒 1 本の箱（画面の座標）。`lo < hi` が項目軸の方向・`a < b` が値の方向 */
export interface BarBox {
  lo: number;
  hi: number;
  a: number;
  b: number;
}

/** {@link mergeBarColumns} の入力（`neg`＝値が基準より負の側へ伸びる棒） */
export interface BarBoxIn extends BarBox {
  neg: boolean;
}

/**
 * 棒の列ごとのまとめ。同じ系列・同じ塗りの棒を、項目軸の方向の `colPx` の列（棒の中心で決める）ごとに 1 つへまとめる。
 * - 幅は合計（塗りの面積＝被覆を保つ）。位置は幅の重みの中心で、列に収まる時は列の中に収める。
 * - 値の方向は和集合（最小〜最大）。正の側と負の側は別にまとめる。
 * @returns 最初に現れた順のまとめた箱
 */
export function mergeBarColumns(
  bars: readonly BarBoxIn[],
  colPx: number = CHART_THIN.barColumnPx,
): BarBox[] {
  const acc = new Map<number, { w: number; moment: number; a: number; b: number; col: number }>();
  for (const bar of bars) {
    const w = bar.hi - bar.lo;
    const mid = (bar.lo + bar.hi) / 2;
    const col = Math.floor(mid / colPx);
    const key = col * 2 + (bar.neg ? 1 : 0);
    const cur = acc.get(key);
    if (cur === undefined) acc.set(key, { w, moment: w * mid, a: bar.a, b: bar.b, col });
    else {
      cur.w += w;
      cur.moment += w * mid;
      cur.a = Math.min(cur.a, bar.a);
      cur.b = Math.max(cur.b, bar.b);
    }
  }
  const out: BarBox[] = [];
  for (const g of acc.values()) {
    const center = g.w > 0 ? g.moment / g.w : (g.col + 0.5) * colPx;
    let lo = center - g.w / 2;
    if (g.w < colPx) {
      const colLo = g.col * colPx;
      lo = Math.min(Math.max(lo, colLo), colLo + colPx - g.w);
    }
    out.push({ lo, hi: lo + g.w, a: g.a, b: g.b });
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
export function clusteredBarColumnPx(seriesCount: number): number {
  const w = Math.round(Math.max(0, seriesCount) * CHART_THIN.barColumnPerSeriesPx);
  return Math.min(CHART_THIN.barColumnMaxPx, Math.max(CHART_THIN.barColumnPx, w));
}

/** 円・ドーナツの扇 1 つ（`fixed`＝`c:dPt` で上書きした扇。束ねない） */
export interface SectorSpan {
  /** 開始・終了の角度（rad・時計回り・`a0 < a1`） */
  a0: number;
  a1: number;
  fixed: boolean;
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
export function thinSectorRuns(
  spans: readonly SectorSpan[],
  ro: number,
  linePx: number,
): [number, number][] {
  const runs: [number, number][] = [];
  if (!(linePx > 0) || !(ro > 0)) return runs;
  const thinOne = (s: SectorSpan): boolean => !s.fixed && ro * (s.a1 - s.a0) <= linePx + 1e-9;
  let i = 0;
  while (i < spans.length) {
    if (!thinOne(spans[i]!)) {
      i++;
      continue;
    }
    let j = i + 1;
    while (j < spans.length && thinOne(spans[j]!)) j++;
    if (j - i >= 2) runs.push([i, j - 1]);
    i = j;
  }
  return runs;
}
