import { type XlsxLimits } from './limits';
import type { Theme } from './theme';
import type { ChartResult } from './types';
/**
 * グラフの部品（`chartN.xml` の全文）を `ChartModel` にする。
 * G1 の外・上限超え・壊れ・描ける系列なしは、例外にせず理由つきの `ok: false`（表示側はどの理由でも破線枠）。
 * @param theme ブックのテーマの 12 色（色の解決用）
 * @param limits `maxChartSeries`・`maxChartPoints`・`maxChartTextChars` を使う
 * @param date1904 ブックが 1904 年起点か（日付の項目の書式化用）
 */
export declare function parseChart(xml: string, theme: Theme, limits: XlsxLimits, date1904: boolean): ChartResult;
