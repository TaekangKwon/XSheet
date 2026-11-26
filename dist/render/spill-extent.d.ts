/**
 * 最後の列からはみ出す文字の分だけ、表の列数を広げる（M7-11・Moobius の依頼 R8）。
 * - 行ごとに「値のある最後のセル」が文字列ではみ出せるときだけ、`列の左端 + 文字の幅の見積もり` を含む列を数える。
 * - 文字の幅は DOM で測らず、文字数から多めに見積もる（多めに出た分は空の列が増えるだけで、文字は切れない）。
 * - 調べるのは描いた行と移動先の行だけ（シートを開くときに全行を読まない）。広げた列数は表示（`sheet-view.ts`）が持つ。
 * DOM 非依存の純関数。
 */
import type { CellFormat, Geometry, TextRun } from '../core/index';
import type { GridRow } from './grid';
/**
 * 文字の幅の見積もり（px・倍率 1）。多めに出す。
 * @param text 表示する文字（1 行。改行は空白として数える）
 * @param format セルの書式（文字の大きさの比・太字・斜体）
 * @param runs リッチテキストの区切り（`<rPr>` のある区切りはその書式で数える。ADR-0009 §3）
 */
export declare function estimateCellTextWidth(text: string, format?: CellFormat, runs?: readonly TextRun[]): number;
/**
 * 行の文字がはみ出して届く列（1 始まり）の最大。はみ出す文字の無い行だけなら 0。
 * 行の「値のある最後のセル」だけを見る（それより左のセルは、右に値があるので最後の列を越えない）。
 * 右寄せ・折り返しのセルは右へはみ出さない。中央はセルの中心から文字の幅の半分だけ右へ届く（ADR-0008 §3）。
 *
 * @param geo 倍率 1 の寸法（表の範囲より右の列も既定の幅で延びるもの）
 * @param rows 調べる行（`GridRow`。`cells[i]` は列 i + 1）
 * @param formats 書式の表（`XlsxSheet.formats`）
 */
export declare function spillReachCols(geo: Pick<Geometry, 'colLeft' | 'colAt' | 'colWidth'>, rows: readonly Pick<GridRow, 'cells' | 'formats'>[], formats: readonly CellFormat[]): number;
