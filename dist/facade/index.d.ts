/**
 * 入口の層（M6-2）。`view`（1 行で表示）・`read`（DOM なしで開く）。
 * core・render の上の薄い層で、InterHub から写したファイルとは別に置く（ADR-0003 §4-1 の 3）。
 */
export type { XSheetSource } from './source';
export type { XSheetTheme } from './style';
export { read, view, type ReadOptions, type Viewer, type ViewOptions, type XSheetLang, } from './view';
