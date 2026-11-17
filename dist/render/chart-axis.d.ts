/**
 * `.xlsx` のグラフの軸の計算（ISSUE-102）：数値軸の範囲・目盛間隔・目盛の値と文字、項目軸の帯・間引き。
 * - 純関数だけ（DOM・`chrome.*` に触らない）。Excel の自動範囲の近似で、係数は {@link AXIS_CONSTANTS} の 1 か所にまとめる。
 *   ISSUE-104 の Excel との照合では、この定数だけを直す（C1・C2・C7）。
 * - 目盛の数には上限を置く（明示の刻みが小さすぎても、自動でも、無限・巨大なループにならない）。
 * - ファイルの文字列は書式コード（ssf に渡す）にだけ使い、数値の文字として返す。
 */
import type { ChartAxis, ChartModel, ChartSeries } from '../core/index';
/** 軸の暫定の係数（C1・C2・C7。ISSUE-104 の照合で直す） */
export declare const AXIS_CONSTANTS: {
    /** C1：dmin がこの割合 × dmax 未満なら自動の最小を 0 にする（5/6） */
    readonly zeroBaselineRatio: number;
    /** C1：自動の範囲の端に足す余り（範囲に対する割合・5%） */
    readonly marginRatio: 0.05;
    /**
     * C1：目盛の数（主の刻みの数）の上限 N ＝ プロットの長さ ÷ 目盛 1 つぶんの最小の間隔。
     * 間隔は縦軸＝文字の高さ × この値（Excel 16 の照合：10.5〜11.0pt の範囲＝9pt の文字で 1.0（手動の内側 110pt で 10 目盛が出る））。
     */
    readonly labelPitchFactor: 1;
    /**
     * C1：横向きの軸（横棒の数値軸）の最小の間隔＝文字の高さ × この値。最長の目盛の文字の幅＋horizontalLabelGapPx がこれより大きければそちら。
     * 3 回目の照合（ISSUE-110・幅の掃引 96 個）：目盛の数の上限が 10 に届くのは、面の幅 360pt（自動配置の外側 459px）と 380pt（485px）の間
     * ＝ 9pt の文字（文字の高さ 14.4px）で最小の間隔 46.8px（3.25）。目盛の数を決める長さは、プロットの内側でなく自動配置の外側の幅
     * （layout の solveAxes の scaleW。Excel は項目軸の文字・目盛の文字の余白を引く前の幅で刻みを決める）。
     * 旧：内側の幅で 2.92（31.6pt）。外側の幅で 3.25 にすると、旧 KNOWN の C1（積み上げの負・7 桁・100% 積み上げ・長い項目名）が合う
     */
    readonly horizontalPitchFactor: 3.25;
    /**
     * C1：横向きの軸で、最長の目盛の文字の幅に足す隙間（px）。隣の文字との隙間を 0 にしないための値。
     * 文字の幅の推定（数字 0.55em）は游ゴシック 9pt の実際より約 10% 広いので、項目軸の {@link labelPaddingPx}（4）は足さず 2 にした
     * （照合：barClustered_big_small は 6 桁の目盛で Excel の間隔が 42.6px 以下。4 を足すと 43.6px で目盛が 1 つ減って外れる）。
     */
    readonly horizontalLabelGapPx: 2;
    /** C1：目盛の数の上限 N の下限と上限（2〜10。Excel 16 の照合：端を丸める前の刻みの数が N 以下の最小の刻み） */
    readonly minTickLimit: 2;
    readonly maxTickLimit: 10;
    /** C1：補助の刻み ＝ 主の刻み ÷ この値 */
    readonly minorDivisions: 5;
    /** C2：100% 積み上げの書式（刻みは他の軸と同じく自動で決める。範囲は 0〜1、負があれば -1〜1） */
    readonly percentFormat: "0%";
    /** 目盛の数の絶対の上限（主・補助。明示の刻みが細かすぎる時の安全弁） */
    readonly maxMajorTicks: 100;
    readonly maxMinorTicks: 500;
    /** C7：項目の文字の間に空ける余白（px） */
    readonly labelPaddingPx: 4;
};
/** データの範囲（最小・最大） */
export interface ValueExtent {
    min: number;
    max: number;
}
/** 数値軸の結果 */
export interface ValueScale {
    min: number;
    max: number;
    /** 主の刻み（正） */
    majorUnit: number;
    /** 補助の刻み。補助の目盛が上限を超える・明示の刻みが主の刻みより粗い時は null */
    minorUnit: number | null;
    /** 反転（`c:orientation` が `maxMin`） */
    reversed: boolean;
    /** 主の目盛の値（min から max まで。上限あり） */
    majorTicks: readonly number[];
    /** 補助の目盛の値（主の目盛と重なるものを除く。上限あり） */
    minorTicks: readonly number[];
    /** `percentStacked`（0〜1・目盛の文字は `0%`） */
    percent: boolean;
}
/**
 * 系列からデータの範囲を求める。点が 1 つも無ければ null。
 * - `clustered`・`standard`：全点の最小・最大。
 * - `stacked`：項目ごとの正の和の最大と負の和の最小（0 を含む）。
 * - `percentStacked`：最小は負の値が 1 つでもあれば -1・無ければ 0。最大は 1（全部が負なら 0。C2。Excel 16 の照合）。点があれば。
 */
export declare function dataExtent(grouping: ChartModel['grouping'], series: readonly Pick<ChartSeries, 'values'>[]): ValueExtent | null;
/**
 * {1, 2, 5}×10ⁿ のうち、`span ÷ u ≤ maxTicks` を満たす最小の刻みを返す（P6）。
 * `span` が 0 以下・非有限なら 1。`maxTicks` は 1 以上に丸める。ループは高々数回。
 */
export declare function niceUnit(span: number, maxTicks: number): number;
/**
 * 目盛の数（主の刻みの数）の上限 N（C1）：プロットの長さ ÷ 目盛 1 つぶんの最小の間隔 を 2〜10 に丸める。
 * 縦軸の間隔は 文字の高さ × {@link AXIS_CONSTANTS.labelPitchFactor}。
 * 横向きの軸（横棒の数値軸）は `labelWidthPx`（最長の目盛の文字の幅）を渡すと、（幅 + {@link AXIS_CONSTANTS.horizontalLabelGapPx}）と
 * 文字の高さ × {@link AXIS_CONSTANTS.horizontalPitchFactor} の大きい方（隣の文字との隙間が 0 にならないようにする）。
 */
export declare function tickCountLimit(lengthPx: number, labelHeightPx: number, labelWidthPx?: number): number;
/**
 * 数値軸の範囲・刻み・目盛を決める（ADR-0031 §決定 5）。
 * 明示（`c:min`・`c:max`・`c:majorUnit`・`c:minorUnit`・`c:orientation`）があれば従い、無い値だけ自動で決める。
 * - 自動の最小：dmin ≥ 0 で dmin < dmax × 5/6 なら 0。dmax ≤ 0 側も対称に 0。それ以外は範囲の 5% の余りを取って刻みに切り下げる。
 * - 自動の最大：dmax に 5% の余りを足して刻みの倍数に切り上げる。
 * - 刻み：{@link niceUnit}（目盛の数の上限 {@link tickCountLimit}）。補助は 1/5。
 * - 点が無ければ 0〜1。`percentStacked` は 0〜1（負を含めば -1〜1）で、刻みは他の軸と同じく自動（{@link tickCountLimit}。明示があれば優先）。
 * - 目盛の数は {@link AXIS_CONSTANTS} の `maxMajorTicks`・`maxMinorTicks` で抑える。
 */
export declare function computeValueScale(input: {
    axis?: ChartAxis;
    extent: ValueExtent | null;
    percentStacked?: boolean;
    lengthPx: number;
    labelHeightPx: number;
    /** 横向きの軸：最長の目盛の文字の幅（px）。あれば目盛の数の上限をこの幅で決める */
    labelWidthPx?: number;
}): ValueScale;
/**
 * 値を軸上の位置（0〜1・既定は範囲外を端に収める）にする。反転（`maxMin`）なら 1 − 位置。
 * 0 は軸の始点側（縦軸なら下・横軸なら左）、反転すると逆の端。
 * `clamp` を false にすると範囲外を収めず、0 未満・1 超の位置をそのまま返す（折れ線を幾何的に切るため。非有限の値は 0）。
 */
export declare function valueFraction(scale: Pick<ValueScale, 'min' | 'max' | 'reversed'>, v: number, clamp?: boolean): number;
/**
 * 項目軸が数値軸と交わる値（`c:crosses`・`c:crossesAt` は項目軸の要素に書かれる＝数値軸上の位置。暫定）。
 * `crossesAt` があればそれ（範囲に収める）、`min`・`max` は端、`autoZero`（既定）は 0 が範囲内なら 0・外なら近い端。
 */
export declare function crossValue(scale: Pick<ValueScale, 'min' | 'max'>, catAxis?: ChartAxis): number;
/**
 * 数値軸の目盛の文字の書式コード（C2）：`percentStacked` で明示が無ければ `0%`。
 * `c:numFmt` が明示で `sourceLinked` でなければその書式。`sourceLinked` か無指定なら最初の系列の書式（無ければ General）。
 */
export declare function tickFormatCode(axis: ChartAxis | undefined, firstSeriesFormat: string | undefined, percent: boolean): string;
/** 目盛の値を書式コードで文字にする（ssf。書式が壊れていても General で出す）。-0 は 0 */
export declare function tickLabel(code: string, value: number): string;
/**
 * 項目軸の位置（軸の始点からの距離 px。項目ごとの中心）。
 * - 既定は帯の等分：中心は (i + 0.5) × 帯の幅。
 * - `midCat`（折れ線の点が目盛の上）：項目が 2 つ以上なら 0 〜 長さを (n−1) 等分、1 つなら中央。
 * - `reversed` なら始点と終点を入れ替える（`c:orientation maxMin`・横棒の既定の向きの補正）。
 */
export declare function categoryPositions(count: number, lengthPx: number, opts?: {
    midCat?: boolean;
    reversed?: boolean;
}): number[];
/** 項目の間隔（px）：`midCat` なら点どうしの間隔、そうでなければ帯の幅 */
export declare function categorySpacing(count: number, lengthPx: number, midCat?: boolean): number;
/**
 * 項目の文字を何個おきに出すか（C7・P7）。
 * - `c:tickLblSkip` の明示（1 以上の整数）があればそれ。
 * - 無ければ、最長の文字の大きさ（軸の向きに沿った幅 px。横軸なら文字の幅・縦軸なら文字の高さ）＋余白が間隔に入る最小の個数。
 *   余白は既定で {@link AXIS_CONSTANTS.labelPaddingPx}。縦に回した文字の間引き（ISSUE-107）は `paddingPx` で小さい余白を渡す。
 */
export declare function categoryLabelSkip(count: number, spacingPx: number, maxLabelExtentPx: number, explicit?: number, paddingPx?: number): number;
/** 文字を出す項目の位置（0, skip, 2·skip, …） */
export declare function categoryLabelIndices(count: number, skip: number): number[];
