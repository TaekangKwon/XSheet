/**
 * `.xlsx` グリッド（`spreadsheet.ts`）の純ロジック（DOM 非依存・ISSUE-84）。
 * 表を Excel の寸法（`lib/xlsx/geometry.ts`）で固定して描くための計算だけを切り出す。
 * - 描く行の範囲・仮想化の上下の余白・セルの内側の高さは、すべて geometry から出す（DOM の実測はしない）。
 * - 非表示の行（高さ 0）は DOM に出さない。描く範囲と前後の余分は「見える行の数」で数え、結合の rowSpan は
 *   描く見える行の数へ縮める（ADR-0026 §決定 5 の予備案。列の畳みは `<col>` の `visibility: collapse` のまま）。
 * - 表の範囲は、データの行・列より広げて、スクロール枠を埋める（Excel のシートのように空セルが続く）。
 * - 文字列のはみ出し（右隣が空なら、その列幅の合計まで文字を出す）は、使える列数と幅の純関数にする。
 */
import { type CellFormat, type CellKind, type CellView, type SheetLayout, type SheetRow, type Geometry } from '../core/index';
/** 列記号の見出し行の高さ（px）。`style.css` の `.xs-grid thead th` の `height` と同じ値にする */
export declare const GRID_HEADER_HEIGHT = 24;
/** Excel の最大行数・最大列数（表を画面いっぱいに広げるときの上限） */
export declare const MAX_GRID_ROWS = 1048576;
export declare const MAX_GRID_COLS = 16384;
/** 隣り合う描く行の間に非表示の行が何行までなら、まとめて 1 回で読むか（結合の途中を正しく扱うため）。これを超える隙間は別々に読む */
export declare const GAP_READ_MAX = 200;
/** セルの内側の要素の左右の余白の合計（px）。`style.css` の `.xs-cell` の左右 padding（3px ずつ）と同じ値 */
export declare const CELL_PAD_X = 6;
/**
 * 行番号の列の幅（px）を行数の桁数から出す。1 桁 8px＋左右の余白 16px、最小 32px。
 * 列幅は固定レイアウトなので、最大の行番号（`rowCount`）が収まる幅を先に決める。
 *
 * @param rowCount シートの総行数
 */
export declare function rowNumWidth(rowCount: number): number;
/** 描く行の範囲（番号はすべて 1 始まり・両端を含む。見える行（高さ > 0）だけを数える。見える行が無いときはすべて 0） */
export interface RowRange {
    /** 見えている最初の行 */
    top: number;
    /** 見えている最後の行 */
    bottom: number;
    /** 描く最初の行（`top` の手前の余分を含む。見える行） */
    first: number;
    /** 描く最後の行（`bottom` の先の余分を含む。見える行） */
    last: number;
}
/**
 * 行 r の次の見える行（高さ > 0）。無ければ 0。非表示の行が続いても `rowAt`（二分探索）で跳ぶ。
 * r の下端の y を含む行は、同じ y に並ぶ非表示の行より後ろの、次の見える行になる（geometry の `rowAt` の仕様）。
 *
 * @param total 行数の上限
 */
export declare function nextVisibleRow(geo: Pick<Geometry, 'rowTop' | 'rowHeight' | 'rowAt'>, r: number, total: number): number;
/** 行 r の前の見える行（高さ > 0）。無ければ 0 */
export declare function prevVisibleRow(geo: Pick<Geometry, 'rowTop' | 'rowHeight' | 'rowAt'>, r: number): number;
/**
 * スクロール位置から、見えている行と、描く行（前後に「見える行」で overscan 行ぶん足した範囲）を出す。
 * 行の高さは可変（Excel の行高・非表示）なので、`rowAt`（geometry の二分探索）で引く。
 * 非表示の行は数えない（フィルターで大半が畳まれていても、描く見える行は一定数）。
 *
 * @param geo `createGeometry` の結果
 * @param scrollTop スクロール枠の `scrollTop`（表の最初の行の上端を 0 とみなす。見出し行は sticky で重なるため無視してよい）
 * @param clientHeight スクロール枠の表示高さ（px）
 * @param total 行数
 * @param overscan 見えている範囲の前後に足す見える行の数
 */
export declare function visibleRowRange(geo: Pick<Geometry, 'rowAt' | 'rowTop' | 'rowHeight'>, scrollTop: number, clientHeight: number, total: number, overscan: number): RowRange;
/**
 * 仮想化で描かない上下の行ぶんの高さ（px）。上＝`rowTop(first)`、下＝`totalHeight − rowTop(last + 1)`。
 * 上の余白＋描く行の高さ＋下の余白が、いつも `totalHeight` に一致する。
 */
export declare function spacerHeights(geo: Pick<Geometry, 'rowTop' | 'totalHeight'>, first: number, last: number): {
    top: number;
    bottom: number;
};
/** 枠の幅が分からない（0＝まだレイアウト前）とき、描く列の範囲に使う代わりの幅（px） */
export declare const COL_FALLBACK_VIEW_WIDTH = 1280;
/** 描く列の範囲（番号はすべて 1 始まり・両端を含む。`RowRange` の列版） */
export interface ColRange {
    /** 見えている最初の列 */
    left: number;
    /** 見えている最後の列 */
    right: number;
    /** 描く最初の列（`left` の手前に 1 画面分の余白を含む） */
    first: number;
    /** 描く最後の列（`right` の先に 1 画面分の余白を含む） */
    last: number;
}
/**
 * スクロール位置から、見えている列と、描く列（左右に 1 画面分の余白を足した範囲）を出す（列の仮想化）。
 * 列幅は可変（Excel の列幅・非表示）なので、`colAt`（geometry の二分探索）で引く。範囲は px で決めるため、
 * 列が極端に狭くても広くても描く列数は画面幅に比例する（遠い列に図や値があっても、16,384 列を描かない）。
 *
 * @param geo `createGeometry` の結果（列数は表の列数）
 * @param scrollLeft スクロール枠の `scrollLeft`
 * @param clientWidth スクロール枠の表示幅（px）。0 以下ならレイアウト前とみなして代わりの幅を使う
 * @param originX 表の A1 の左上の x（行番号の列の幅。スクロール枠の内容の左端から）
 * @param total 列数
 */
export declare function visibleColRange(geo: Pick<Geometry, 'colAt'>, scrollLeft: number, clientWidth: number, originX: number, total: number): ColRange;
/**
 * 描かない左右の列ぶんの幅（px）。左＝`colLeft(first)`、右＝`totalWidth − colLeft(last + 1)`。
 * 左の余白＋描く列の幅＋右の余白が、いつも `totalWidth` に一致する（`spacerHeights` の列版）。
 */
export declare function colSpacerWidths(geo: Pick<Geometry, 'colLeft' | 'totalWidth'>, first: number, last: number): {
    left: number;
    right: number;
};
/**
 * 行のセルを、描く列 first〜last（1 始まり）だけにする。返す行の `cells[i]` は列 `first + i`。
 * - 描く範囲の左端が結合セルの途中（結合の起点が範囲の左外）なら、起点のセルを範囲の左端へ移し、`colSpan` を範囲内へ縮める。
 *   起点が上の行にある（`rowSpan` で覆われる）セルは、これまでどおり null のまま（td を作らない）。
 * - 右端で切れる結合セルは、範囲の右端までに `colSpan` を縮める。縮めたセルは結合のまま扱う（はみ出さない）。
 *
 * @param row 全列を持つ行
 * @param first 描く最初の列
 * @param last 描く最後の列
 */
export declare function clipRowToColumns(row: GridRow, first: number, last: number): GridRow;
/**
 * 表の範囲（データの行・列）の外にある余白の位置へ、行の書式 → 列の書式の順で書式の番号を当てる（ADR-0030 §決定 7 の余白への拡張）。
 * 読み取り（`getRows`）は表の範囲の中だけを解くため、余白の行（`r > dataRows`）と余白の列（`列 > dataCols`）は、
 * Excel で行・列全体に続く塗りがここで抜ける。表の範囲は広げない（描く余白の位置に番号を足すだけ）。
 * - 行の書式（`rowFormat`）があればそれ（-1＝書式なしは何も当てず、列の書式も打ち消す）。無ければ列の書式（`colFormat`・-1 は無し）。
 * - 描く列 first〜last のうち、余白の位置だけを見る。表の中の位置（読み取りが解いた番号）は変えない。
 * - 余白に触れる行にだけ `rowFormat` を 1 回呼ぶ。`colFormat` は列ごとに 1 回（結果を覚える）。入力の行は書き換えず、
 *   足した行だけ新しい `GridRow` にする（足すものが無い行は同じ行をそのまま返す）。
 *
 * @param rows 描く行（`collectVisibleRows` の結果）
 * @param first 描く最初の列（1 始まり）
 * @param last 描く最後の列
 * @param dataRows 表のデータの行数（`XlsxSheet.rowCount`）
 * @param dataCols 表のデータの列数（`XlsxSheet.colCount`）
 * @param rowFormat 行 r の行の書式の番号（`XlsxSheet.rowFormat`）
 * @param colFormat 列 c（1 始まり）の列の書式の番号（`XlsxSheet.colFormat`）
 */
export declare function applyMarginFormats(rows: readonly GridRow[], first: number, last: number, dataRows: number, dataCols: number, rowFormat: (r: number) => number | undefined, colFormat: (c: number) => number): GridRow[];
/**
 * セルの内側の要素の高さ（px）。セル（td）の高さ（行高、結合は描く行の高さの合計）から下罫線 1px を引く。0 を下回らない。
 */
export declare function cellInnerHeight(cellHeight: number): number;
/**
 * 描く行のうち、i 番目の行から span 行ぶん（結合した、描く見える行の数）の高さの合計（px）。
 *
 * @param rows 描く行（見える行だけ。上から順）
 */
export declare function spanHeight(geo: Pick<Geometry, 'rowHeight'>, rows: readonly {
    r: number;
}[], i: number, span: number): number;
/** 描くセル。`CellView` に、結合セルかどうかを足したもの */
export interface GridCell extends CellView {
    /** 結合セル。行の畳みで `rowSpan` が 1 になっても結合のまま扱う（はみ出さない） */
    merged?: boolean;
}
/** 描く行（`SheetRow` の上位互換。見える行だけ） */
export interface GridRow {
    /** 1 始まりの行番号 */
    r: number;
    /** 長さは列数以下。範囲より右・範囲外の列は `undefined`（空のセル）、null＝結合で隠れる */
    cells: (GridCell | null | undefined)[];
    /**
     * 罫線の番号（`SheetRow.borders` を写したもの。長さは `cells` と同じ・値は `XlsxSheet.borders` の位置）。
     * 罫線のあるセルが無い行は持たない。結合で隠れるセル（`cells` が null）にも番号が入る。
     */
    borders?: (number | undefined)[];
    /**
     * 書式の番号（`SheetRow.formats` を写したもの。長さは `cells` と同じ・値は `XlsxSheet.formats` の位置）。
     * 書式のあるセルが無い行は持たない。結合で隠れるセル（`cells` が null）にも番号が入る。
     */
    formats?: (number | undefined)[];
}
/**
 * 文字列がはみ出して使える列数（自分の列を含む）を決める。Excel の「文字列は右隣が空ならはみ出し、
 * 値のあるセルの手前で切る」を表す。td は 1 セル 1 個のまま（罫線を残す）で、文字だけをはみ出させる。
 * - 文字列のとき（`kind === 's'` で非空、結合の無いセル）だけ右隣の空セルを数える。型で判定する
 *   （数字だけの文字列・「▲1,234」は文字列なのではみ出す。数値・真偽・エラー・日付、型の無いセルは自分の列で切る）。
 * - 結合セル・空のセルは 1（自分の列だけ）。
 * - 値のあるセル（数式の結果が "" の文字列セルを含む）・結合で隠れるセル（null）・結合セルの手前で止める。HTML の `colspan` 上限と揃えて 1000 列まで。
 *
 * @param cells 行のセル（`GridRow.cells`）
 * @param c 対象の列（0 始まり）
 * @param cols 表の列数。`cells` より右の列（足した空の列）も空のセルとして数える
 * @returns 使える列数（1 以上）。1 ならはみ出さない
 */
export declare function spillColumns(cells: readonly (GridCell | null | undefined)[], c: number, cols?: number): number;
/**
 * はみ出しに「紙でない塗り」が絡むか（ADR-0030 §決定 9）。はみ出す元のセル（列 c）と、`spillColumns` が数えた右の列のどれかに
 * 白以外の塗りがあれば true（入れ物へ `is-over-fill` を付け、文字の下に面の色を敷かない）。白の塗りは塗りなしとして数える。
 * 追加の走査はせず、数えた列の `formats` を見るだけ。
 *
 * @param table 書式の表（`XlsxSheet.formats`）
 * @param formats 行の書式の番号（`GridRow.formats`。描く列に揃えたもの。無ければ false）
 * @param c はみ出す元の列（0 始まり）
 * @param columns `spillColumns` の結果（自分の列を含む列数）
 */
export declare function spillOverFill(table: readonly CellFormat[], formats: readonly (number | undefined)[] | undefined, c: number, columns: number): boolean;
/**
 * はみ出す文字の入れ物の幅（px）。自分の列から `columns` 列ぶんの列幅の合計から、右端の罫線 1px を引く。
 * 非表示の列（幅 0）は合計に入らない。自分の列が非表示（幅 0）なら 0（隠れたセルの文字を出さない）。
 *
 * @param geo `createGeometry` の結果
 * @param c 対象の列（0 始まり）
 * @param columns `spillColumns` の結果（自分の列を含む列数）
 */
export declare function spillWidth(geo: Pick<Geometry, 'colLeft' | 'colWidth'>, c: number, columns: number): number;
/**
 * 表を画面いっぱいに広げたときの行数・列数。Excel のシートのように、データの先にも空のセル
 * （既定の幅と高さ。geometry は範囲外の番号も既定の寸法で延ばす）が続くようにする。
 * 行数＝max(データの行数, `minRows`, スクロール枠の高さを埋める行数)、列数も同じ。上限は Excel の最大（1,048,576 行・16,384 列）。
 * 後続の画像（図のある範囲まで表を広げる）は `minRows`・`minCols` に図の最下行・最右列を渡す。
 *
 * @param layout `XlsxSheet.layout`
 * @param dataRows データの行数（`XlsxSheet.rowCount`）
 * @param dataCols データの列数（`XlsxSheet.colCount`）
 * @param viewWidth スクロール枠の表示幅（px。`clientWidth`）
 * @param viewHeight スクロール枠の表示高さ（px。`clientHeight`）
 * @param minRows 表示する行数の最小（既定 0）
 * @param minCols 表示する列数の最小（既定 0）
 */
export declare function gridExtent(layout: SheetLayout, dataRows: number, dataCols: number, viewWidth: number, viewHeight: number, minRows?: number, minCols?: number): {
    rows: number;
    cols: number;
};
/**
 * 行 first〜last のうち、見える行（高さ > 0）だけを DOM 用の行にして返す。
 * 非表示の行は返さない。結合セル（`rowSpan` > 1）は、描く見える行の数へ縮める：
 * - 先頭の行が非表示の結合は、結合内の最初に見える行へセルを移す（値はその行の td に出る）。
 * - 非表示の行を跨ぐ結合の `rowSpan` は、見える行の数にする。
 * 隣り合う見える行の間の非表示の行が `gapMax` 行までなら、まとめて 1 回で読む（結合の途中を正しく扱うため）。
 * これを超える隙間（フィルターで大半を畳んだシートなど）は別々に読み、隙間の向こうへ続く結合は、
 * 次の見える行に（同じ文字の）結合セルが現れる。
 *
 * @param geo `createGeometry` の結果
 * @param first 描く最初の行
 * @param last 描く最後の行
 * @param total 行数
 * @param readRows 行 a〜b（連続）を読む関数（`XlsxSheet.getRows` と同じ形・範囲外の行は呼び出し側で空の行にする）
 * @param gapMax まとめて読む隙間の上限
 */
export declare function collectVisibleRows(geo: Pick<Geometry, 'rowTop' | 'rowHeight' | 'rowAt'>, first: number, last: number, total: number, readRows: (a: number, b: number) => SheetRow[], gapMax?: number): GridRow[];
/**
 * 数値の表示文字の推定幅（px）。DOM は測らず「文字数 × MDW」で見積もる。
 * 数字・符号・E・％・通貨記号などは MDW、小数点・桁区切り・コロン・スラッシュ（日付）・空白は MDW の半分とする。
 *
 * @param text 表示文字
 * @param mdw 最大桁幅 MDW（`SheetLayout.mdw`）
 */
export declare function estimateNumberWidth(text: string, mdw: number): number;
/**
 * 収まらない数値を `###` にする（Excel と同じ）。数値・日付（`kind` が `n`・`d`）の表示文字の推定幅が
 * セルの内側の幅（セル幅 − 右罫線 1px − 左右の余白）を超えたら、`#` の繰り返し（内側の幅 ÷ MDW ÷ 文字サイズの比（1 以上）の切り捨て個数、最低 1）を返す。
 * 収まるときは表示文字のまま。文字列・真偽・エラーは対象外（そのまま）。
 *
 * @param text 表示文字
 * @param kind セルの型
 * @param cellWidth セルの幅（px。結合は結合範囲の列幅の合計）
 * @param mdw 最大桁幅 MDW
 * @param sizeRatio 文字サイズの比（`CellFormat.sizeRatio`。推定幅に掛ける。既定 1。太字は掛けない・ADR-0030 §決定 10）
 */
export declare function fitNumberText(text: string, kind: CellKind | undefined, cellWidth: number, mdw: number, sizeRatio?: number): string;
/**
 * セルの型ごとの既定の配置（Excel の「標準」。セルの配置指定は読まない）。
 * 数値・日付は右、真偽・エラーは中央、文字列・型なしは左（クラス無し）。
 */
export declare function cellAlignClass(kind: CellKind | undefined): 'is-right' | 'is-center' | '';
