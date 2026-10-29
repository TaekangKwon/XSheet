// 正: docs/02_architecture.md §3
/**
 * render の入口（`xsheet/render`）。DOM の表・罫線・図の層。
 * InterHub `viewers/xlsx-*.ts`・`spreadsheet.ts` の .xlsx 部分から作る。`renderSheet` は M2-3、`renderWorkbook` は M2-4。
 */
export { MESSAGES_EN, MESSAGES_JA, type XSheetMessages } from './messages';
export { renderSheet, type RenderOptions, type SheetView } from './render-sheet';
export { renderWorkbook, type WorkbookOptions, type WorkbookView } from './render-workbook';
