import type { CellRange } from './types';
/**
 * 範囲 1 つの文字（シート名を除いた部分）を読む。`$` は無くてもよい。
 * `A1:C3`・`A1`・`A:C`（行の全体）・`1:3`（列の全体）。読めない・範囲外は null
 */
export declare function parseAreaRef(ref: string): CellRange | null;
/**
 * 定義の中身（実体参照を戻した後）を、範囲ごとの文字（シート名の部分を除く）に分ける。
 * シート名の引用（`'a,b'!A1`・`''` は `'`）の中のカンマ・`!` では分けない。走査は 1 文字ずつの線形
 */
export declare function splitAreas(formula: string): string[];
/**
 * workbook.xml の印刷範囲を、シートの位置（0 始まり・`<sheets>` の順＝`localSheetId`）ごとに返す。
 * 同じシートの定義が 2 つあれば先の物を使う。範囲が 1 つも読めないシートは入れない。
 */
export declare function readPrintAreas(wbXml: string, sheetCount: number, maxPrintAreas: number): Map<number, readonly CellRange[]>;
