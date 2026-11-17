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
export declare const CHART_THIN: {
    /** 項目の間隔がこれ未満のグラフを「密」とみなす */
    readonly denseSpacingPx: 1;
    /** M4 の列の幅 */
    readonly gridPx: 0.5;
    /** Douglas–Peucker の許容 */
    readonly epsPx: 0.25;
    /** その他のマーカーで、直前に残した点からこの距離以内なら省く */
    readonly markerSpacingPx: 1;
    /** 棒をまとめる列の幅 */
    readonly barColumnPx: 1;
    /** 集合の棒で、系列 1 つにつき列の幅へ足す量（列の幅＝系列の数×これを整数 px へ丸めた値。1px 以上・下の上限以下） */
    readonly barColumnPerSeriesPx: 0.2;
    /** 集合の棒の列の幅の上限（これより広げると列が縞に見えて元の絵から離れる。ISSUE-111 の計測） */
    readonly barColumnMaxPx: 2;
    /** 滑らかな線を折れ線に直す時の区間ごとの分割数 */
    readonly smoothSteps: 4;
};
/**
 * M4 の間引き：項目軸の方向に `gridPx` の幅で並ぶ列ごとに、同じ列に続く点の最初・最小・最大・最後だけを元の順で残す。
 * 続く点を列とするので、項目の並びが逆順（右から左・上から下）でも動く。
 * @param pts 点列
 * @param catAxis 項目軸の方向（0＝x・1＝y）。最小・最大は他方の座標で決める
 * @param gridPx 列の幅
 * @returns 点が 3 未満なら入力の写し。それ以外は間引いた点列（最初と最後の点は必ず残る）
 */
export declare function thinM4(pts: readonly Pt[], catAxis: 0 | 1, gridPx?: number): Pt[];
/**
 * Douglas–Peucker で残す点の添字（昇順）。反復で解く（再帰は深くなるため使わない）。端点は必ず残す。
 * 残さなかった点は、残した点を結ぶ線分から `eps` 以内にある。
 * 最悪の入力（鋸歯など）では 2 乗の時間になるため、距離の計算が {@link DP_WORK_BUDGET} 回を超えたら、
 * 残りの区間は全点を残す（点が増えるだけで、ずれの上限は守られる）。
 */
export declare function simplifyIndices(pts: readonly Pt[], eps?: number): number[];
/** Douglas–Peucker（端点を残す・許容 `eps`）で単純化した点列。点が 3 未満なら入力の写し */
export declare function simplifyDP(pts: readonly Pt[], eps?: number): Pt[];
/**
 * 折れ線の間引き：M4 → Douglas–Peucker。中心線のずれは 0.25px 以下（M4 の列の中は正確）。
 * @param pts 線の 1 つの部分（空の点・プロットの外で切れた後の 1 本）の点列
 * @param catAxis 項目軸の方向（0＝x・1＝y）
 */
export declare function thinPolyline(pts: readonly Pt[], catAxis: 0 | 1): Pt[];
/**
 * 円のマーカーの帯の区間分け。続く中心の間隔が `maxGap` 以下の区間（2 点以上）を「帯」にする。
 * 帯は呼び出し側が間引いた中心線を、太さ 2R・端と結びが round の線 1 本で描く。区間に入らない点は単独のマーカー。
 * @param centers マーカーの中心（元の順）
 * @param maxGap 続く中心の間隔の上限。{@link bandGap} で求める
 * @returns `runs`＝帯の区間（各 2 点以上・元の順）／`singles`＝帯に入らない点（元の順）
 */
export declare function splitMarkerBands(centers: readonly Pt[], maxGap: number): {
    runs: Pt[][];
    singles: Pt[];
};
/**
 * 円のマーカー（外側の半径 R）の帯にできる、続く中心の間隔の上限 √(8R·eps)。
 * 円の連なりの縁の波 d²/8R がこの間隔で eps 以下になる。
 */
export declare function bandGap(R: number, eps?: number): number;
/**
 * その他のマーカーの間引きで残す点の添字（昇順）。Douglas–Peucker（許容 `eps`）で選んだ点は残し、
 * その間は直前に残した点から `spacing` 以内の点を省く（省いた点は残した点から `spacing` 以内）。
 */
export declare function thinMarkerIndices(centers: readonly Pt[], eps?: number, spacing?: number): number[];
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
export declare function mergeBarColumns(bars: readonly BarBoxIn[], colPx?: number): BarBox[];
/**
 * 集合の棒（重なり < 100）の列の幅。系列が多いほど同じ画素に重なる矩形が増えるので、
 * 系列の数 × `barColumnPerSeriesPx` を整数 px に丸めた幅へ広げて矩形の数を減らす（1px 以上・`barColumnMaxPx` 以下）。
 * 被覆（幅の合計）は保たれ、値の方向の包絡の細かさだけが 1〜2px になる（ISSUE-111・案 D）。
 * 整数 px にするのは、画素の格子とずれる列（2.5px など）が縞になって元の絵から離れるため（実測：1px→2px は差が小さく、2.5〜4px は大きい）。
 * @param seriesCount 系列の数
 */
export declare function clusteredBarColumnPx(seriesCount: number): number;
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
export declare function thinSectorRuns(spans: readonly SectorSpan[], ro: number, linePx: number): [number, number][];
