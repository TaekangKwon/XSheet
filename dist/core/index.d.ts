import type { BuiltinDates } from './format-ja';
import { type XlsxLimits } from './limits';
import { type XlsxBook } from './types';
export { XlsxError } from './types';
export type { AnchorPoint, ArrowSize, ArrowType, BorderEdge, BorderStyle, CellBorder, CellFormat, CellKind, CellView, ChartAxis, ChartBox, ChartFailReason, ChartLegend, ChartLine, ChartMarker, ChartMarkerSymbol, ChartModel, ChartPaint, ChartResult, ChartSeries, ChartTextStyle, ChartTickMark, ChartTitle, ColLayout, CustGeomDef, CustGeomPath, DrawingAnchor, DrawingColor, DrawingGroup, DrawingKind, DrawingLine, DrawingNotice, DrawingText, DrawingTextParagraph, DrawingTextRun, GeomArg, GeomPathFill, GroupLevel, GroupRect, ImageFailReason, ImageFormat, ImageResult, LineDash, LineEnd, LineGeom, ManualLayout, RowLayout, ShapeGeom, ShapeOutline, SheetDrawing, SheetDrawings, SheetInfo, SheetLayout, SheetRow, TextAlign, TextAnchor, TextVert, XlsxBook, XlsxErrorKind, XlsxSheet, } from './types';
export type { Rgba } from './color';
export { createGeometry, colWidthToPx, defaultColWidthToPx, ptToPx } from './geometry';
export { fontDefaultRowPx, lookupMdw } from './font-metrics';
export type { Geometry } from './geometry';
export type { XlsxLimits } from './limits';
export { DEFAULT_LIMITS } from './limits';
export { colName } from './xml';
export { isPaperFill } from './cell-format';
export { findText } from './find-text';
export type { FindTextOptions, FindTextResult, TextMatch } from './find-text';
export type { BuiltinDates } from './format-ja';
/** openXlsx の追加の指定 */
export interface OpenXlsxOptions {
    /** 上限を替える（既定は DEFAULT_LIMITS。試験で小さい値を渡す用） */
    limits?: Partial<XlsxLimits>;
    /**
     * 組み込みの日付の書式 27〜36・50〜58 の見せ方（既定 `'default'`＝西暦の表のまま）。
     * `'japanese-era'` で日本語版 Excel と同じ表（28 は `令和8年11月11日`、27 は `R8.11.11` など）にする。
     * 書式に直接書かれた和暦（`ggg`・`e` など）と曜日（`aaa`・`aaaa`）は、指定しなくてもいつも出す。
     * XSheet: ADR-0007 §6（M7-8）
     */
    builtinDates?: BuiltinDates;
}
/**
 * xlsx を開く（シートの中身は openSheet で開くときに初めて展開する）。
 * @throws XlsxError zip でない（not-zip）・暗号化（encrypted）・壊れている（broken）・大きすぎる（too-large）・読めない形式（unsupported）
 */
export declare function openXlsx(data: ArrayBuffer | Uint8Array, options?: OpenXlsxOptions): Promise<XlsxBook>;
