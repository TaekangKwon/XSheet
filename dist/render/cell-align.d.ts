/**
 * セルの配置と折り返しの描き方（M7-13・Moobius の依頼 R9）。
 * - 入れ物（`div.xs-cell`）に配置のクラスを付ける（{@link alignClasses}）。規則は `style.css` の `is-h-*`・`is-v-*`・`is-wrap`。
 * - 右寄せの文字列は左へ、中央の文字列は両側へはみ出す（{@link spillSides}・{@link spillAround}）。左寄せ・標準の右へのはみ出しは
 *   今の `grid-table.ts`（`spillColumns`・`spillOut`）のまま。折り返すセルははみ出さない。
 * - 位置の判定（`navigate.ts`）も {@link spillSides} で、左右どちらから届くはみ出しかを引く。
 * `grid-table.ts`（InterHub の写し）には `// XSheet:` の印で呼び出しの口だけを足す。
 * セル値は `textContent` だけで入れる（信頼できない入力）。
 */
import type { CellFormat, CellKind, Geometry } from '../core/index';
import { type GridCell, type GridRow } from './grid';
/** 横の配置（標準は型で決めたもの） */
export type HAlign = 'left' | 'center' | 'right';
/**
 * 実際の横の配置。指定（`hAlign`）が無ければ型で決める（数値・日付は右、真偽・エラーは中央、ほかは左）。
 */
export declare function resolveHAlign(kind: CellKind | undefined, format?: CellFormat): HAlign;
/** 折り返すセルか（文字列で `wrap` の指定がある） */
export declare function isWrapped(kind: CellKind | undefined, format?: CellFormat): boolean;
/**
 * 入れ物に足すクラス。指定の無いセルは空（型ごとの既定は `grid.ts` の `cellAlignClass`）。
 * - 横：文字列は `is-h-center`・`is-h-right`（`safe` を付けない）。文字列でないセル（数値など）は今の `is-center`・`is-right`
 *   （`safe`：収まらない `###` は左から）。左は付けない。
 * - 縦：`is-v-top`・`is-v-center`。折り返し：`is-wrap`（文字列だけ）。
 */
export declare function alignClasses(kind: CellKind | undefined, format?: CellFormat): string[];
/** 左右へはみ出す列数（自分の列を含まない） */
export interface SpillSides {
    left: number;
    right: number;
}
/**
 * 文字列が左右へはみ出して使える列数。左寄せ・標準は右だけ、右寄せは左だけ、中央は両側。
 * はみ出すのは型が文字列で値のある、結合の無い、折り返さないセルだけ。値のあるセル・結合で隠れるセル・結合の手前で止める。
 *
 * @param cells 行のセル（`GridRow.cells`）
 * @param k 対象の位置（`cells` の添字）
 * @param format 対象のセルの書式
 * @param lo 左へ数える下限（含む。描く列の左端）
 * @param hi 右へ数える上限（含まない。描く列の右端の次・表の列数）
 */
export declare function spillSides(cells: readonly (GridCell | null | undefined)[], k: number, format: CellFormat | undefined, lo?: number, hi?: number): SpillSides;
/**
 * 右寄せ・中央の文字列のはみ出しを描く（`grid-table.ts` から呼ぶ）。左寄せ・標準・折り返しのセルは何もせず false
 * （左寄せ・標準は呼び出し側の今の右へのはみ出しで描き、折り返しははみ出さない）。
 * - 入れ物を `margin-left` で左へ広げ、幅を左右の列幅の分だけ足す。中央は左右の余白の差で、文字の中心を元のセルの中心に合わせる。
 * - 文字を `span.xs-text` で包む（`style.css` の `.is-spill .xs-text` が文字の範囲だけ罫線を覆う）。
 * - 白以外の塗りが絡むなら `is-over-fill`（文字の下に面の色を敷かない）。列ごとの文字色の塗り分けはしない（ADR-0008 §3）。
 *
 * @param box セルの入れ物（`cellBox` の結果。文字は直下のテキスト）
 * @param row 描く行（範囲で切ったもの。`cells[k]` が対象）
 * @param k 対象の位置（`row.cells` の添字）
 * @param n 描く列の数（右へ数える上限）
 * @param geo 寸法
 * @param c 対象の列（0 始まり。geometry は c + 1）
 * @param table 書式の表（`XlsxSheet.formats`）
 * @param text セルの文字
 * @returns 呼び出し側の右へのはみ出しを止めるなら true（右寄せ・中央・折り返し）
 */
export declare function spillAround(box: HTMLDivElement, row: GridRow, k: number, n: number, geo: Geometry, c: number, table: readonly CellFormat[], text: string): boolean;
