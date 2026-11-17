/**
 * リッチテキストの区切りを描く（M7-17）。
 * - 区切りの文字を `span.xs-runs` 1 つで包む（`.xs-cell`・`.xs-text` は flex なので、包まずに分けると各区切りが別の
 *   flex 項目になり、端の空白・折り返しが崩れる。`highlight.ts` の `span.xs-hits` と同じ）。
 * - `font` の無い区切りはテキストノード、ある区切りは span に CSSOM で書式を当てる。
 *   値は整数・真偽から組んだ固定の語だけ（ファイルの文字列は入れない）。文字は `createTextNode` だけで入れる。
 */
import { type CellFormat, type TextRun } from '../core/index';
/** 文字の [from, to)（`CellView.text` 上の位置）を、区切りの書式で分けたノードにする */
export type RunSlicer = (from: number, to: number) => Node[];
/**
 * 区切りで分けるノードの作り手。
 * @param text セルの文字（`CellView.text`。区切りの連結と同じ）
 * @param runs `CellView.runs`
 * @param cell セルの書式（区切りの大きさの比の基準・自動の色）
 */
export declare function runSlicer(text: string, runs: readonly TextRun[], cell: CellFormat | undefined): RunSlicer;
/**
 * セルの入れ物の文字を区切りで描く（入れ物・はみ出しの span を作った後、検索の一致を付ける前に呼ぶ）。
 * @param box セルの入れ物（文字は直下のテキストか、はみ出しなら `span.xs-text` の中）
 * @param shown 表示した文字（文字列なので `text` と同じ。違えば描かない）
 * @returns 区切りで描いたときの作り手（検索の一致の片の中も区切りで分けるため）。描かなければ undefined
 */
export declare function applyRuns(box: HTMLElement, text: string, shown: string, runs: readonly TextRun[] | undefined, cell: CellFormat | undefined): RunSlicer | undefined;
