/**
 * `.xlsx` のセルの書式（塗り・文字の装飾）の CSS を作る純関数（ISSUE-100・DOM 非依存）。
 * - 書式ごとに規則 1 行：`.xs-sheet[data-xs-fmt="<連番>"] td.xs-xf-<番号> { … }`。
 *   連番の属性で絞るため、既存の `.xs-grid td`（グリッド線）より詳細度が強く、同じ文書の別の表示と混ざらない。
 * - 文字列は固定の語・整数・整数から組んだ `#rrggbb`・小数（文字サイズの比）・固定のテーマ変数の参照
 *   （白の塗りの `var(--xs-cell-bg)` だけ）で組む。ファイルの文字列は入れない。
 * - ダーク表示の描き分けは規則に焼き込む（テーマ変数は増やさず、切替で描き直さない）：
 *   白以外の塗り＝Excel の色の面と黒の文字（明示色はそのまま）、塗りなし・白の塗り（紙）＝画面の面と文字色（明示色だけ書く）。
 */
import { type CellFormat } from '../core/index';
/**
 * 書式の番号 → td に足すクラス名（`xs-xf-<番号>`）。
 * @param index `XlsxSheet.formats` の位置
 */
export declare function formatClassName(index: number): string;
/**
 * 「紙でない」塗りがあるか。はみ出す文字の下の面（グリッド線隠し）を敷くかどうかの判定に使う（ADR-0030 §決定 9）。
 * 白の塗り・塗りなし・書式なし（undefined）は false。
 */
export declare function hasOverFill(format: CellFormat | undefined): boolean;
/**
 * はみ出した「自動」の文字色を、列ごとの区間で塗り分けるグラデーション（`linear-gradient`）を作る（ADR-0030 §決定 9）。
 * 塗りの上の自動色は黒、塗りなし・紙の上は画面の文字色で、はみ出した文字が塗りの状態の違う列をまたぐと、
 * 単色ではダーク表示でどちらかの面の上に読めない色が残る（黄塗りの黒文字が暗い面へ・明るい文字が薄青の面へ）ため、列の面に合わせる。
 * - 呼び出し側は `background-image` にこれを入れ、`background-clip: text; color: transparent` を足す（span の左端＝元の左端）。
 * - 元の文字が明示色（`CellFormat.color`。黒は自動に寄せ済み）なら null（Excel と同じく単色）。
 * - 区間は元の列から右へ。隣り合う同じ色は 1 区間にまとめ、幅 0（非表示の列）は数えない。区間が 1 つなら null（塗り分け不要）。
 * - 文字列は固定の語・整数の px・`#000000`・`var(--xs-fg)` だけ（ファイルの文字列は入らない）。
 *
 * @param table 書式の表（`XlsxSheet.formats`）
 * @param formats 行の書式の番号（`GridRow.formats`。描く列に揃えたもの。無ければ全列「塗りなし」）
 * @param c はみ出す元の列（0 始まり。`formats` の添字）
 * @param widths 元の列から右へ数えた列幅（px）。`spillColumns` の列数ぶん。不正な値は 0 とみなす
 * @returns グラデーションの文字列。塗り分けが要らなければ null
 */
export declare function spillInkGradient(table: readonly CellFormat[], formats: readonly (number | undefined)[] | undefined, c: number, widths: readonly number[]): string | null;
/**
 * 書式の表から CSS の文字列を作る（ブックを開いた時に 1 回）。
 * 書式の番号は `formats` の位置のまま（宣言が 1 つも出ない書式は規則を作らないが、後ろの番号はずれない）。
 *
 * @param formats `XlsxSheet.formats`（ブックで共有する書式の表）
 * @param seq 表示ごとの連番（`.xs-sheet` の `data-xs-fmt`）。整数でなければ 0
 * @returns 規則を改行でつないだ文字列。書式が無ければ空の文字列
 */
export declare function buildFormatCss(formats: readonly CellFormat[], seq: number): string;
