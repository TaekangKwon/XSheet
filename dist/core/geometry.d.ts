import type { SheetLayout } from './types';
/** 行高 pt → px。`round(pt × 96 / 72)`（15pt＝20px、18.75pt＝25px） */
export declare function ptToPx(pt: number): number;
/**
 * 列幅（`<col width>` の文字数・余白を含む値）→ px。
 * ECMA-376 Part 1 の `col` 要素の換算式 `trunc(((256·w + trunc(128/MDW)) / 256) · MDW)`。
 * @param w 列幅（文字数）
 * @param mdw 最大桁幅 MDW（px）
 */
export declare function colWidthToPx(w: number, mdw: number): number;
/**
 * 既定の列幅（`defaultColWidth` が無いとき）→ px。`ceil((baseColWidth·MDW + 余白) / 8) · 8`
 * （MDW 7 で 64px、MDW 8 で 72px）。余白は `2·ceil(MDW/4) + 1`（MDW 5〜8 では 5px）。
 * 根拠: Excel 16 の実ファイルの `<col width>` の端数から余白を逆算（2026-10-27）。
 */
export declare function defaultColWidthToPx(baseColWidth: number, mdw: number): number;
/** 換算した座標（px・すべて整数）。番号は 1 始まり */
export interface Geometry {
    /** 列 c の左端の x（c ≤ 1 は 0）。cols より右の列も既定の幅で延ばして返す */
    colLeft(c: number): number;
    /** 列 c の幅（非表示は 0。c < 1 は 0） */
    colWidth(c: number): number;
    /** 行 r の上端の y（r ≤ 1 は 0）。rows より下の行も既定の高さで延ばして返す */
    rowTop(r: number): number;
    /** 行 r の高さ（非表示は 0。r < 1 は 0） */
    rowHeight(r: number): number;
    /**
     * y を含む行（1〜rows）。負・NaN は先頭、下端以降は最後の行へ丸める。rows が 0 なら 0。
     * 非表示の行（高さ 0）は、同じ y の次の表示行より前になるため返らない（末尾だけが非表示のときを除く）。
     */
    rowAt(y: number): number;
    /** x を含む列（1〜cols）。丸め方は `rowAt` と同じ */
    colAt(x: number): number;
    /** 表の幅（1〜cols の列幅の合計） */
    readonly totalWidth: number;
    /** 表の高さ（1〜rows の行高の合計） */
    readonly totalHeight: number;
}
/**
 * layout から座標（px）の計算器を作る。
 * - 列幅は `<col width>` を `colWidthToPx` で換算し、指定の無い列は既定の列幅（`defaultColWidth` があれば同じ式で換算、
 *   無ければ `defaultColWidthToPx`）。非表示の列は 0。
 * - 行高：シートの既定の行高は、`customHeight="1"` なら `defaultRowHeight`（pt）の換算値、
 *   そうでなければフォントの既定の行高（`fontRowPx`。実測表）を優先し、無ければ `defaultRowHeight` の換算値。
 *   （Excel は `customHeight` の無い `defaultRowHeight`・`<row ht>` を、その PC のフォントから計算した行高で置き換えるため）
 *   行は、`customHeight="1"` の `ht` は `ptToPx(ht)`。`customHeight` の無い `ht` は、ファイルの
 *   `defaultRowHeight` と ±0.01 で等しければシートの既定、違えば保存時の自動調整値として `ptToPx(ht)` とする。
 *   `ht` が無ければシートの既定。非表示の行は 0。
 *   `zeroHeight` のときは、`<row>` の無い行を既定で高さ 0 とする（`<row>` のある行は表示）。
 * - 列幅・行高は Excel の上限（255 文字・409.5pt）に丸める。
 * @param layout `XlsxSheet.layout`
 * @param rows 表の行数（`XlsxSheet.rowCount`）。`totalHeight`・`rowAt` の範囲に使う
 * @param cols 表の列数（`XlsxSheet.colCount`）。`totalWidth`・`colAt` の範囲に使う
 */
export declare function createGeometry(layout: SheetLayout, rows: number, cols: number): Geometry;
