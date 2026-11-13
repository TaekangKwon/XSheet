/**
 * styles.xml の書式の読み取り：`<fills>`（塗り）・`<fonts>`（文字の装飾）と、cellXfs の `fontId`・`fillId`。
 * ブックで 1 回だけ読み、セルの `s` 属性（cellXfs の位置）→ 書式の番号 → 塗りと文字の装飾、と引けるようにする（`FormatTable`）。
 * - 見える書式（塗り・太字・斜体・下線・取り消し線・自動でない色・比が 1 でない文字サイズのどれか）が無い xf は「書式なし」にまとめる。
 * - 同じ見た目の書式は 1 つにまとめる（キーは数値だけ。ファイルの文字列は持たない）。まとめた後の数は `maxCellFormats` まで。
 * - 塗りは 1 色に畳む：`solid`＝`fgColor`、ほかの模様＝前景と背景の混色、グラデーション＝50% の位置の色。
 * - `applyFill`・`applyFont` は見ない（F5）。cellStyleXfs は読まない。フォント名は読まない。
 * - 色は SpreadsheetML の `<color>`（cell-color.ts）をテーマで解いてから持つ。解けない値は例外にせず既定の色・書式なしに倒す。
 * - 走査は線形（開きタグは `[^<>]*>`、閉じは `indexOf`）。
 */
import { type Rgba } from './color';
import { type XlsxLimits } from './limits';
import type { Theme } from './theme';
import type { CellFormat } from './types';
/** 書式の表 */
export interface FormatTable {
    /** 書式の番号 → 書式（見える書式だけ・重複なし。`SheetRow.formats` の値・`XlsxSheet.formats` の位置） */
    readonly formats: readonly CellFormat[];
    /** cellXfs の位置（セルの `s`）→ 書式の番号。書式なし・不正・上限の外は -1 */
    readonly xfFormat: Int32Array;
}
/** 書式の無い表（styles.xml が無い・壊れているとき。共有する・変更しない） */
export declare const EMPTY_FORMAT_TABLE: FormatTable;
/**
 * styles.xml から書式の表を作る（ブックで 1 回）。
 * @param xml styles.xml（無ければ null＝書式なし）
 * @param theme ブックのテーマ（色の `theme` を解く）
 * @param limits `maxFonts`・`maxFills` を超えた要素は読まない（それを参照するセルは装飾なし・塗りなし）。
 *   重複を除いた書式の数が `maxCellFormats` を超えた書式のセルは書式なし
 */
export declare function parseCellFormats(xml: string | null, theme: Theme, limits: XlsxLimits): FormatTable;
/**
 * 「紙」の塗りか：解決後の塗りが不透明な `#FFFFFF`（`rgb`・`theme`・`indexed` など出どころは問わない）。
 * 紙は塗りなしと同じに描く（ADR-0030 1.1）ため、表の範囲を広げる判定（sheet.ts）と描き分け（viewer）の両方がこれを使う。
 * `#FEFEFE`・`#F2F2F2`・半透明の白は紙ではない（alpha < 1）。
 */
export declare function isPaperFill(fill: Rgba | undefined): boolean;
/**
 * 書式の色に `theme` の指定があるか（`<fonts>`・`<fills>`・`<borders>` の中だけを見る）。無ければテーマを読まなくても
 * 色は決まるので、呼び出し側（index.ts）はテーマの読み取りを省ける（図の無いブックでテーマの読み取りを増やさない。ISSUE-94）。
 */
export declare function stylesUseTheme(xml: string | null): boolean;
