/**
 * 検索の一致のハイライト（M7-5・Moobius の依頼 R2）。
 * - 一致の一覧（`findText` の `matches`）を「行 → 列 → 範囲」の 2 段の表にして表示が持ち、
 *   表を描くたびに描く行・列の一致だけを付け直す（仮想スクロールで描き直しても残る）。
 * - セルの文字を一致の範囲で分け、**分けた全体を `span.xs-hits` 1 つで包む**。`.xs-cell`・`.xs-text` は flex なので、
 *   包まずに分けると各片が別の flex 項目になり、片の端の空白が消える。
 * - 一致の片は `span.xs-hit`（`data-xs-hit` に一覧の番号）。今のヒットは `is-current`。色は CSS 変数
 *   `--xs-hit-bg`・`--xs-hit-current-bg` で、背景だけを付ける（行の高さ・列の幅を変えない）。
 * - 自動の色の文字が塗りをまたぐはみ出し（`background-clip: text` で描く）は、片に `is-ink`（`--xs-fg` の文字の色）を付ける。
 *   不透明なハイライトが、親の背景で描く文字を隠さないようにするため。
 * - `###` で出すセル（数値が列幅に収まらない）は、表示の文字と一致の位置が合わないので、セル全体を光らせる。
 * 文字は `textContent`（テキストノード）だけで入れる（セル値は信頼できない入力）。
 */
import type { TextMatch } from '../core/index';
/** 1 セルの中の一致 */
export interface CellHit {
    /** `CellView.text` 上の位置（UTF-16） */
    start: number;
    /** 終わり（含まない） */
    end: number;
    /** 一致の一覧（`setHighlights` に渡した配列）の番号 */
    index: number;
}
/** 表を描くときに引くハイライト */
export interface GridHits {
    /** 今のヒットの番号（無ければ -1）。`focusHighlight` が替える（M7-6） */
    current: number;
    /** 一致の一覧の数（`setHighlights` に渡した配列の長さ） */
    readonly count: number;
    /** 番号 index の一致（番号の正しくない一致・範囲外は undefined） */
    match(index: number): TextMatch | undefined;
    /** 行 r・列 c（1 始まり）の一致（無ければ undefined） */
    at(r: number, c: number): readonly CellHit[] | undefined;
}
/**
 * 一致の一覧から、表を描くときに引く表を作る。null・空の配列は null（ハイライトなし）。
 * 番号の正しくない一致（整数でない行・列、範囲の逆転）は飛ばす。同じセルの一致は始まりの順に並べる。
 */
export declare function createGridHits(matches: readonly TextMatch[] | null | undefined, current?: number): GridHits | null;
/**
 * 描いた表の中の片の `is-current` を、今のヒットに合わせて付け替える（表を描き直さない。ADR-0007 §3）。
 * @returns 今のヒットの片が描いた表の中にあったか
 */
export declare function markCurrent(table: ParentNode, current: number): boolean;
/** 片の番号（`data-xs-hit`）。片でなければ -1 */
export declare function hitIndexOf(el: Element | null): number;
/**
 * セルの入れ物（`div.xs-cell`）の文字に一致を付ける。`grid-table.ts` が入れ物を作った後に呼ぶ。
 * @param box セルの入れ物（文字は直下のテキストか、はみ出しなら `span.xs-text` の中）
 * @param text セルの文字（`CellView.text`）
 * @param shown 表示した文字（収まらない数値は `###`）
 * @param hits このセルの一致（無ければ何もしない）
 * @param current 今のヒットの番号
 */
export declare function decorateCell(box: HTMLElement, text: string, shown: string, hits: readonly CellHit[] | undefined, current: number): void;
