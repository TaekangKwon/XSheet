/**
 * `.xlsx` のグラフの配置の計算（ISSUE-102）：{@link ChartModel} とグラフの面の px の大きさから、
 * 軸の範囲・目盛・文字、既定の色、タイトル・凡例・プロットの矩形、棒・線・マーカー・扇の図形を計算し、
 * 「場面」（{@link ChartScene}）を返す純関数。
 * - DOM・canvas・`chrome.*` に触らない。文字の幅は文字の種類ごとの係数で推定する（P4）。
 * - 場面の値は数値・`#rrggbb`・不透明度・固定の語と、`text` の文字だけ。ファイル由来の文字列（系列名・項目・タイトル）は
 *   `text` 項目の `text` にしか入らない（`path` の `d` は数値と固定の命令だけ）。ISSUE-103 が属性へ文字列を入れずに済む。
 * - Excel の既定値（系列色の変化・文字の大きさ・余白・凡例の寸法・目盛の長さなど）は {@link CHART_DEFAULTS}・{@link CHART_LAYOUT}・
 *   {@link CHART_COLOR_VARIATIONS} の表にまとめる。ISSUE-104 の Excel との照合ではこの表だけを直す（C3〜C10）。
 */
import { type Rgba } from '../core/internal';
import { type Theme, type ThemeSlot } from '../core/internal';
import type { ChartMarkerSymbol, ChartModel, ChartTickMark, LineDash } from '../core/index';
/** 場面の要素数の上限（円の扇・点ごとの上書き・文字の数。超えたら null＝破線枠。ADR-0031 §決定 9） */
export declare const MAX_CHART_SCENE_ITEMS = 5000;
/** 色と不透明度（`#rrggbb`＋0〜1） */
export interface ChartColor {
    color: string;
    opacity: number;
}
/** 線の見た目（太さは px。端・結びは固定の語） */
export interface ChartStroke extends ChartColor {
    width: number;
    dash: LineDash;
    cap: 'butt' | 'round';
    join: 'miter' | 'round';
}
export type ChartAnchor = 'start' | 'middle' | 'end';
export type ChartBaseline = 'alphabetic' | 'middle' | 'hanging';
/** 矩形（`fill`・`stroke` が null なら塗らない・描かない） */
export interface ChartRectItem {
    kind: 'rect';
    x: number;
    y: number;
    width: number;
    height: number;
    fill: ChartColor | null;
    stroke: ChartStroke | null;
    /** 角丸の半径（px。`rx`・`ry` の数値）。グラフの面の矩形だけが持つ（`c:roundedCorners` が真の時。ISSUE-109）。無ければ直角 */
    radius?: number;
}
/** 経路（`d` は数値と固定の命令 `M L C A Z` だけ） */
export interface ChartPathItem {
    kind: 'path';
    d: string;
    fill: ChartColor | null;
    stroke: ChartStroke | null;
}
/** 線分 */
export interface ChartLineItem {
    kind: 'line';
    x1: number;
    y1: number;
    x2: number;
    y2: number;
    stroke: ChartStroke;
}
/** 文字（`text` だけがファイル由来の文字列。回転は (x, y) を中心に度） */
export interface ChartTextItem {
    kind: 'text';
    x: number;
    y: number;
    text: string;
    size: number;
    bold: boolean;
    color: ChartColor;
    anchor: ChartAnchor;
    baseline: ChartBaseline;
    rotate: number;
}
export type ChartItem = ChartRectItem | ChartPathItem | ChartLineItem | ChartTextItem;
/** 場面：グラフの面の大きさ（px）と、描く順の要素の並び */
export interface ChartScene {
    width: number;
    height: number;
    /** グラフの面が塗りなしの時に真（ダーク表示で白い下地を敷く。ADR-0031 §決定 10） */
    backdrop: boolean;
    /** グラフの面の角丸の半径（px）。0 または無ければ直角。ダーク表示の下地（`backdrop`）の矩形にも同じ値を付ける（ISSUE-109） */
    cornerRadius?: number;
    items: ChartItem[];
}
/** pt → px（96/72） */
export declare const PX_PER_PT: number;
/** 色の既定の指定（テーマの色の位置＋lumMod・lumOff。1/1000 %） */
interface ColorSpec {
    slot: ThemeSlot;
    lumMod?: number;
    lumOff?: number;
}
/**
 * 部品の既定（`c:spPr`・`c:txPr`・`c:style` が無い時。C4・C8）。
 * ISSUE-104：`c:spPr` 等の無い合成のグラフを Excel 16 で開いた値に合わせた（面の枠線 0.5pt・#898989、文字 10pt・黒、タイトルは太字、折れ線 1.5pt、マーカー 6pt）。
 * Excel が保存したグラフは `c:spPr`・`c:txPr` を持つのでこの表は使われない。
 */
export declare const CHART_DEFAULTS: {
    /** グラフの面：塗り `bg1`・枠線 #898989（黒＋輝度 53.725%）・0.5pt */
    readonly chartAreaFill: ColorSpec;
    readonly chartAreaLine: ColorSpec;
    readonly chartAreaLinePt: 0.5;
    /**
     * 主の目盛線・補助の目盛線・項目軸の線・数値軸の線（`c:spPr` の無い軸。ISSUE-105 の PNG の画素の照合）。
     * Excel 16 は主の目盛線と両方の軸の線を黒（`dk1`）・補助の目盛線を #BCBCBC（黒＋輝度 73.725%）の 1 画素の線で描く（COM の `Format.Line` の黒 0.5pt と合う）。
     * 旧値（目盛線 #d9d9d9・項目軸 #bfbfbf・数値軸なし）は、Excel が COM で作って保存した `c14:style` 付きのグラフ（#898989）とも合わなかった。
     */
    readonly gridMajor: ColorSpec;
    readonly gridMinor: ColorSpec;
    readonly catAxisLine: ColorSpec;
    readonly valAxisLine: ColorSpec;
    /** 上の線の太さ（pt。COM の Weight。Excel は 1 画素未満の線を 1 画素に描く） */
    readonly axisLinePt: 0.5;
    /**
     * `c14:style` を持つグラフ（Excel が COM の旧 API などで作った `c:spPr` 無しのグラフ。ISSUE-109・ISSUE-105 の COM）：
     * 主の目盛線・項目軸・数値軸の線は #898989（黒＋輝度 53.725%）0.5pt（太さは {@link axisLinePt}）、タイトルは 18pt 太字。
     * 補助の目盛線・軸と凡例の文字・面・折れ線は `c14:style` 無しと同じ。暗い背景（`c:style` 41〜48）とは独立に効く
     */
    readonly c14Line: ColorSpec;
    readonly c14TitlePt: 18;
    /**
     * `c:style` 41〜48（暗い背景。ISSUE-109・ISSUE-110 の COM と PNG の値）：面の塗り 黒・面の枠線 黒 0.5pt（Excel の COM は太さ「自動」で、黒い面の上では見えない）・
     * 文字（軸・凡例・タイトル）白・プロット（棒・折れ線）#3F3F3F・円とドーナツのプロットは塗りなし（黒い面が見える）。
     * 軸線・補助の目盛線・折れ線の太さ・系列の色の組は未確定で変えない（41・43〜48 の系列の色の組・円の凡例の `#595959` 9pt はこの ISSUE の範囲外）。
     * ダーク表示でも補正しない（P5）
     */
    readonly darkStyleMin: 41;
    readonly darkStyleMax: 48;
    readonly darkChartAreaFill: Rgba;
    readonly darkChartAreaLine: Rgba;
    readonly darkText: Rgba;
    readonly darkPlotFill: Rgba;
    /**
     * 暗い背景（41〜48）の棒の既定の塗り（単色近似。ISSUE-109・ADR-0031 §決定 4）。Excel の棒は上から下へのグラデーション（明→系列の色に近い色→暗）で、
     * 50% の位置の色は系列の色の彩度を 110% にした色（`a:satMod` 110%。ADR-0028 §決定 6.4 の `a:gradFill`＝50% の位置の色と同じ規則）。`linearGradient` は使わない。1/1000 %
     */
    readonly darkBarSatMod: 110000;
    /** 文字（黒） */
    readonly text: ColorSpec;
    /**
     * グラフの面の角丸（`c:roundedCorners` が真・要素が無い時も真。ISSUE-109）。半径は 8.25pt（11px）で、グラフの大きさによらず一定
     * （Excel 16 の PNG：216×144・360×216・720×432pt のどれも 11px。小さい面では 幅÷2・高さ÷2 を超えない）
     */
    readonly roundedCorners: true;
    readonly cornerRadiusPt: 8.25;
    /** 円・ドーナツの枠線 */
    readonly pieLine: ColorSpec;
    /** 円・ドーナツの枠線の太さ（pt）。Excel は点（`c:dPt`）ごとに 1.5pt を書く（点の線は読まないので、系列に明示が無ければこの値。C9） */
    readonly pieLinePt: 1.5;
    /** 線の太さ（pt）：マーカーの枠・線だけの記号 0.75、折れ線 1.5（c:spPr の無いグラフの値。目盛線・軸線は {@link axisLinePt}） */
    readonly thinLinePt: 0.75;
    readonly seriesLinePt: 1.5;
    /** マーカーの自動の形と大きさ（pt・C8） */
    readonly autoMarker: ChartMarkerSymbol;
    readonly markerSizePt: 6;
    /** 文字の大きさ（pt・C4）：タイトル 10（太字）・軸と凡例 10・軸タイトル 10 */
    readonly titlePt: 10;
    readonly titleBold: true;
    readonly axisPt: 10;
    readonly legendPt: 10;
    readonly axisTitlePt: 10;
    /** ドーナツの穴（%・C9） */
    readonly holeSizePercent: 50;
    /** 棒の隙間（%）・重なり（集合 0・積み上げ 100） */
    readonly gapWidth: 150;
    readonly overlapClustered: 0;
    readonly overlapStacked: 100;
    /** 円の開始角（度・12 時から時計回り） */
    readonly firstSliceAng: 0;
    /** 項目軸の文字の回転の既定（`a:bodyPr rot` が無く、自動の形の判定を使わない時＝横棒の項目軸・小さい面。P7・ISSUE-107：回さない） */
    readonly autoRotation: 0;
    /** `c:majorTickMark`・`c:minorTickMark` の要素が無い軸の目盛の印（スキーマの既定。Excel 16 の COM は 4 つとも Cross を返す・C12） */
    readonly tickMark: ChartTickMark;
};
/**
 * 配置の寸法（C5・C6・C10。pt は × 96/72、`Ratio` は文字の大きさ（px）に対する割合）。
 * ISSUE-104 で Excel 16 の COM の値（座標は面の左上から 4pt 内側が原点）に合わせた：
 * タイトルの上端は面の上から 6pt・プロット（軸の文字を除く内側）の上と右はタイトル／凡例の下・面の右（凡例の左）から 11pt、
 * 軸の文字の左と下は面の左・下（凡例の右・上）から 5pt、凡例は面の縁から 6pt。円・ドーナツは四方 11pt。
 */
export declare const CHART_LAYOUT: {
    /** 軸の文字の左・下の余白（pt。面の左端・下端、凡例の右・上から） */
    readonly marginPt: 5;
    /** プロットの上・右の余白（pt。タイトル・凡例の下／面の右・凡例の左から）。円・ドーナツは四方 */
    readonly plotPadPt: 11;
    /** タイトルの上端（pt） */
    readonly titleTopPt: 6;
    /** 文字の高さ ＝ 大きさ × この値 */
    readonly lineHeightRatio: 1.2;
    /** タイトルの高さ ＝ 大きさ × この値（Excel の Calibri 14pt で 1.435・10pt 太字で 1.52） */
    readonly titleLineHeightRatio: 1.45;
    /** 目盛の長さ（主・補助。pt）と、目盛の文字との間隔（px） */
    readonly tickLengthPt: 5;
    readonly minorTickLengthPt: 3;
    readonly labelGapPx: 3;
    /** 軸タイトルの帯の余白（px） */
    readonly axisTitleGapPx: 4;
    /** 凡例 */
    readonly legendSymbolRatio: 0.55;
    readonly legendLineSymbolRatio: 1.33;
    readonly legendGapRatio: 0.4;
    /**
     * 横並び（1 行）で項目と項目の間（C6）。Excel の項目の間隔は、記号＋文字の幅（実測）+ 約 0.47〜0.49em。
     * 文字の幅の推定が広めに出る名前（円 5 項目の凡例）があるので、1 回目・2 回目の照合の両方が 3% 以内に収まる 0.4 にした
     */
    readonly legendEntryGapRatio: 0.4;
    /** 横並びの格子の桁の余白（最も長い項目の幅に足す。Excel の桁の幅は文字の幅の推定より少し広い。C6） */
    readonly legendColPadRatio: 0.1;
    readonly legendRowGapRatio: 0.675;
    /** 凡例の箱の余白（左右・上下） */
    readonly legendPadRatio: 0.34;
    readonly legendPadYRatio: 0;
    readonly legendMaxFraction: 0.4;
    /** 右上の凡例でプロットを凡例の下まで縮める条件：縮めた後のプロットの高さが縮める前のこの割合以上（Excel 16 の COM：小さい面は 0.50・標準は 0.71） */
    readonly legendCornerShrinkMin: 0.6;
    /** 凡例と面の縁（タイトルの下）の間（pt） */
    readonly legendEdgeGapPt: 6;
    /** 回転した文字の本文の最大の長さの打ち切り（1 行の文字数） */
    readonly maxLineChars: 255;
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
    readonly catRotateMinSlotRatio: 1.6;
    readonly catRotatedBandExtraPt: 6.5;
    readonly catWrapRowRatio: number;
    readonly catRotatedSkipPadPt: 1;
    readonly catReachGapPt: 2;
    readonly catBandMaxRatio: 0.4;
};
/** 文字の幅の係数（P4・暫定）：半角の英数字と記号・半角カナ・全角・その他（em に対する割合） */
export declare const TEXT_WIDTH_EM: {
    readonly ascii: 0.55;
    readonly halfKana: 0.5;
    readonly full: 1;
    readonly other: 0.6;
};
/** 既定の色の変化（C3）：周回 k（9 周で繰り返す）の lumMod・lumOff（1/1000 %）。k=0 は変化なし */
export declare const CHART_COLOR_VARIATIONS: readonly {
    lumMod?: number;
    lumOff?: number;
}[];
/**
 * 既定の系列（・点）の色（C3）：位置 `index`（系列は `c:idx`・点は点の位置）に対し、テーマの `accent((index mod 6)+1)` に
 * 周回 k＝floor(index / 6) の変化（{@link CHART_COLOR_VARIATIONS}。9 周で繰り返す）を掛ける。
 */
export declare function autoColor(theme: Theme, index: number): Rgba;
/** 文字の幅の推定（px）。`sizePx` は文字の大きさ（px） */
export declare function estimateTextWidth(text: string, sizePx: number): number;
/** 1 行を最大 `max` 文字（既定 255）に切る（超えたら末尾を `…` にする） */
export declare function truncateLine(text: string, max?: number): string;
/** 項目軸の文字の形（ISSUE-107。`rotate45` は右上がり＝−45°・`rotate90` は下から上＝−90°） */
export type CategoryLabelForm = 'horizontal' | 'wrap' | 'rotate45' | 'rotate90';
/**
 * 1 つの項目名を、幅 `slotPx` に収まるよう空白で折り返した行にする（語を左から詰める）。
 * 1 行に収まれば `[text]`。どれかの語の幅が `slotPx` を超える（空白が無い・語が長い）時は折り返せないので `null`。
 */
export declare function wrapCategoryLabel(text: string, sizePx: number, slotPx: number): string[] | null;
/**
 * 縦棒・折れ線の項目軸の文字の形を選ぶ（ISSUE-107・ADR-0031 §決定 5。`a:bodyPr rot` の明示が無い時だけ呼ぶ）。
 * 上から順に判定し、最初に当たった形にする。
 * 1. 水平：どの項目名の幅（推定）も `slotPx`（項目の間隔）以下。
 * 2. 折り返し：幅が超える項目名が全部、空白を含み、どの語も `slotPx` に入る（`rows` は行数の最大）。
 * 3. 45°：`slotPx` が 45° の境目（{@link CHART_LAYOUT.catRotateMinSlotRatio} × 文字の大きさ）以上。
 * 4. 90°：それより狭い（呼び出し側が入らない分を間引く）。
 */
export declare function planCategoryLabelForm(labels: readonly string[], sizePx: number, slotPx: number): {
    form: CategoryLabelForm;
    rows: number;
};
/**
 * グラフの場面を計算する（ISSUE-102・ADR-0031 §決定 3・6）。
 * @param model `parseChart` の結果のグラフ
 * @param widthPx グラフの面の幅（px）
 * @param heightPx グラフの面の高さ（px）
 * @param theme テーマの色（既定の色に使う）
 * @param opts `thin`＝点の多いグラフ（項目の間隔 1px 未満）の幾何を間引くか（既定 真。ADR-0031 §決定 11.1。偽は照合・試験で間引きの有無を比べる用）
 * @returns 場面。大きさが正の有限値でない・場面の要素が {@link MAX_CHART_SCENE_ITEMS} を超えるなら null（表示側は破線枠）
 */
export declare function layoutChart(model: ChartModel, widthPx: number, heightPx: number, theme?: Theme, opts?: {
    thin?: boolean;
}): ChartScene | null;
export {};
