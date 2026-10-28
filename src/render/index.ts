// 正: docs/02_architecture.md §3
/**
 * render の入口（`xsheet/render`）。DOM の表・罫線・図の層。
 * M2 で InterHub `viewers/xlsx-*.ts`・`spreadsheet.ts` の .xlsx 部分から `renderSheet`・`renderWorkbook` を作る。
 */
export { MESSAGES_EN, MESSAGES_JA, type XSheetMessages } from './messages';
