import type { XlsxLimits } from './limits';
import type { Theme } from './theme';
import type { CellBorder } from './types';
/** 罫線の表 */
export interface BorderTable {
    /** 罫線の番号 → 罫線（線のある罫線だけ。`SheetRow.borders` の値・`XlsxSheet.borders` の位置） */
    readonly borders: readonly CellBorder[];
    /** cellXfs の位置（セルの `s`）→ 罫線の番号。罫線なし・不正・上限の外は -1 */
    readonly xfBorder: Int32Array;
}
/** 罫線の無い表（styles.xml が無い・壊れているとき。共有する・変更しない） */
export declare const EMPTY_BORDER_TABLE: BorderTable;
/**
 * styles.xml から罫線の表を作る（ブックで 1 回）。
 * @param xml styles.xml（無ければ null＝罫線なし）
 * @param theme ブックのテーマ（色の `theme` を解く）
 * @param limits `maxBorders` を超えた `<border>` は読まない（それを参照するセルは罫線なし）
 */
export declare function parseBorders(xml: string | null, theme: Theme, limits: XlsxLimits): BorderTable;
