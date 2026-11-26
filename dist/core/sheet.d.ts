/**
 * シート 1 枚（sheetN.xml）の読み取り：行の位置の索引と、範囲の行だけの展開・結合セル。
 * - 開くときは `<row>` の位置の索引と、最後の行・列、結合セルだけを作る。セルは展開しない。
 * - 展開（getRows）は呼ばれた範囲の行だけを作って返し、展開した行を保持しない。
 * - 寸法（`<cols>`・`<sheetFormatPr>`・行の `ht`／`hidden`）は `layout` に疎な形で持つ（px への換算は geometry.ts）。
 *   行の寸法は、行の位置の索引を作る走査で同時に拾う（XML を 2 度読まない）。
 * - 名前空間の接頭辞付き XML（`<x:c>`）・`r` 属性の省略・`t="d"`・inlineStr・真偽・エラー・数式の保存済みの値を読む。
 */
import type { BorderTable } from './borders';
import { type FormatTable } from './cell-format';
import type { CellFormatter } from './format';
import { type XlsxLimits } from './limits';
import { type SharedStrings } from './shared-strings';
import { type XlsxSheet } from './types';
/** シートを読むのに要るもの（ブック全体で 1 つ） */
export interface SheetContext {
    strings: SharedStrings;
    formatter: CellFormatter;
    date1904: boolean;
    limits: XlsxLimits;
    /** 最大桁幅 MDW（px。styles.xml の先頭の font から。寸法の換算用） */
    mdw: number;
    /** 先頭の font の既定の行高（px・96dpi。実測表に無ければ null。寸法の換算用） */
    fontRowPx: number | null;
    /** 罫線の表（ブックで 1 つ。styles.xml の cellXfs の位置 → 罫線の番号。ADR-0029） */
    borderTable: BorderTable;
    /** 書式の表（ブックで 1 つ。styles.xml の cellXfs の位置 → 書式の番号。ADR-0030） */
    formatTable: FormatTable;
}
/** 開いたシート（getRows のほか、メモリを放す release を持つ） */
export interface ReleasableSheet extends XlsxSheet {
    release(): void;
}
/** 空のシート（グラフシートなど、セルを持たないもの） */
export declare function createEmptySheet(): ReleasableSheet;
/**
 * シートの XML から、シートを開く。
 * @throws XlsxError Strict 形式は unsupported、`<sheetData>` の閉じが無ければ broken、行・結合が多すぎれば too-large
 */
export declare function createSheet(xml: string, ctx: SheetContext): ReleasableSheet;
