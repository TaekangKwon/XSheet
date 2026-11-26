/**
 * 印刷の描画（`renderForPrint`・XSheet の追加。M7-20）。
 * 表示中のシートのデータの範囲を、用紙の中身の大きさのページに切り分けて、仮想化せずに描く。今の表示（倍率・位置）には触らない。
 * - ページは「下へ、それから右へ」の順（Excel の既定）。各ページは、その行・列の範囲を窓にした固定の表示：
 *   表（`buildGridTable`）・罫線の層（`svg.xs-borders` 1 枚）・図の層（範囲で切る入れ物の中）を重ね、ページの端で切る。
 *   ブラウザの改ページには任せない（幅のはみ出しは次のページへ送られず切れるため）。
 * - 範囲は、データ・図の右下・はみ出しの届く列（全部の行から）。枠を埋める空のセルは入れない。何も描く物の無いページは出さない。
 *   ファイルに印刷範囲（`_xlnm.Print_Area`）があれば、その範囲を定義の順に、範囲ごとにページを分けて描く（§3・M7-23）。
 * - 倍率はいつも 100%。`fit: 'width'` のときだけ、全部の列が 1 ページの幅に収まるまで縮める。
 * - 描くセルの数が上限を超えたら、超える手前の行で打ち切る（列は打ち切らない）。
 * - 画像・グラフを描き終えてから解決する。中断（`signal`・表示の destroy・シートの切り替え）では、足した物を片付けてから AbortError。
 * ユーザーのファイル内容は `textContent` だけで入れる（表・図の層の部品と同じ）。例外のメッセージにセル値・シート名を含めない。
 */
import type { CellRange, colName, createGeometry, Geometry, SheetDrawing, XlsxBook, XlsxSheet } from '../core/index';
import type { createBorderLayer } from './border-layer';
import type { buildBorderPaths } from './borders';
import type { drawingExtent, DrawingBox, rectsIntersect } from './drawings';
import type { buildFormatCss } from './format-css';
import type { applyMarginFormats, collectVisibleRows, rowNumWidth } from './grid';
import type { buildGridTable } from './grid-table';
import type { createImageLayer } from './image-layer';
import type { XSheetMessages } from './messages';
import type { spillReachCols } from './spill-extent';
import type { zoomGeometry } from './zoom';
/**
 * 印刷の描画が使う、表示と共有の部品（print-gate.ts の `PRINT_KIT`）。
 * このファイルは遅れて読む部品なので、表示の部品は import せずに受け取る（単独版の xsheet-print.mjs に表示の部品を二重に入れないため。ADR-0010 §2）。
 */
export interface PrintKit {
    colName: typeof colName;
    createGeometry: typeof createGeometry;
    createBorderLayer: typeof createBorderLayer;
    buildBorderPaths: typeof buildBorderPaths;
    drawingExtent: typeof drawingExtent;
    rectsIntersect: typeof rectsIntersect;
    buildFormatCss: typeof buildFormatCss;
    applyMarginFormats: typeof applyMarginFormats;
    collectVisibleRows: typeof collectVisibleRows;
    rowNumWidth: typeof rowNumWidth;
    buildGridTable: typeof buildGridTable;
    createImageLayer: typeof createImageLayer;
    spillReachCols: typeof spillReachCols;
    zoomGeometry: typeof zoomGeometry;
    /** 見出し行の高さ（grid.ts の `GRID_HEADER_HEIGHT`） */
    GRID_HEADER_HEIGHT: number;
    /** Excel の最大の行数・列数（grid.ts の `MAX_GRID_ROWS`・`MAX_GRID_COLS`） */
    MAX_GRID_ROWS: number;
    MAX_GRID_COLS: number;
}
/** 印刷の指定（ADR-0010 §2-1） */
export interface PrintOptions {
    /**
     * 1 ページの中身の大きさ（CSS px。96px = 1in）。既定は A4 縦（210×297mm）から Excel の「標準」の余白
     * （左右 0.7in・上下 0.75in）を引いた 659×978。使う側の `@page` と同じ大きさにすること。
     */
    page?: {
        width: number;
        height: number;
    };
    /** `'width'` は全部の列が 1 ページの幅に収まるように縮める（縦だけページを分ける）。既定は縮めない（100%） */
    fit?: 'width';
    /** 列記号・行番号を各ページに出す（既定 false） */
    headers?: boolean;
    /** 目盛り線（セルの境の薄い線）を出す（既定 false）。罫線はいつも出す */
    gridlines?: boolean;
    /** 描くセルの数の上限（行 × 列。非表示の行・列は数えない）。既定 200000 */
    maxCells?: number;
    /**
     * ファイルの印刷範囲（`_xlnm.Print_Area`）を使わず、データの範囲を印刷する（Excel の「印刷範囲を無視」）。
     * 既定 false（印刷範囲があれば、その範囲だけを範囲ごとにページを分けて描く）
     */
    ignorePrintArea?: boolean;
    /** 中断する（描いている間・画像を待つ間に見る） */
    signal?: AbortSignal;
}
/** 印刷の描画の結果（ADR-0010 §2-1） */
export interface PrintResult {
    /** 描いたページの数（表の無い表示・シートを開けていない・描く物の無いシートは 0） */
    readonly pages: number;
    /** 上限（`maxCells`）で行を打ち切った */
    readonly truncated: boolean;
    /** `target` に足した要素を外し、blob URL を放す。何度呼んでもよい */
    destroy(): void;
}
/** 既定の 1 ページの中身の大きさ（A4 縦・Excel の「標準」の余白。CSS px） */
export declare const PRINT_PAGE_DEFAULT: Readonly<{
    width: number;
    height: number;
}>;
/** 描くセルの数の上限の既定（仮の値。ADR-0010 §2-5） */
export declare const PRINT_MAX_CELLS_DEFAULT = 200000;
/** 解いた印刷の指定 */
export interface ResolvedPrintOptions {
    page: {
        width: number;
        height: number;
    };
    fit: boolean;
    headers: boolean;
    gridlines: boolean;
    maxCells: number;
    ignorePrintArea: boolean;
}
/**
 * 印刷の指定を解く。
 * @throws RangeError 用紙の大きさが正の有限の数でない・`fit` が `'width'` でない・`maxCells` が 1 以上の数でない
 */
export declare function resolvePrintOptions(options?: PrintOptions): ResolvedPrintOptions;
/** 1 ページの範囲（行・列は 1 始まり・両端を含む） */
export interface PrintPage {
    r1: number;
    r2: number;
    c1: number;
    c2: number;
}
/** 1 つの範囲のページの切れ目 */
export interface PrintAreaBands {
    /** 行のページの切れ目（上から） */
    rowBands: readonly [number, number][];
    /** 列のページの切れ目（左から） */
    colBands: readonly [number, number][];
}
/** 印刷の割り付け（DOM を使わない計算の結果） */
export interface PrintPlan {
    /** 描く最後の行（全部の範囲で一番下） */
    rows: number;
    /** 描く最後の列（全部の範囲で一番右） */
    cols: number;
    /** 倍率（`fit: 'width'` のときだけ 1 未満） */
    zoom: number;
    truncated: boolean;
    /** 範囲ごとの切れ目（印刷範囲の順。印刷範囲の無いシートはデータの範囲 1 つ）。範囲ごとにページを分ける */
    areas: readonly PrintAreaBands[];
}
/** 割り付けに使うシートの部分 */
export type PrintSheet = Pick<XlsxSheet, 'layout' | 'rowCount' | 'colCount' | 'getRows' | 'formats'>;
/**
 * 番号 start〜end を、長さ limit に収まる区切りに分ける（ページの切れ目）。切れ目は行・列の境。
 * 1 つだけで limit を超える番号は、その番号だけの区切りにする。長さ 0 の番号（非表示）だけの区切りは返さない。
 *
 * @param edge 番号 i の始まりの位置（`rowTop`・`colLeft`）
 * @param at 位置を含む番号（`rowAt`・`colAt`）
 */
export declare function splitBands(start: number, end: number, edge: (i: number) => number, at: (v: number) => number, limit: number): [number, number][];
/**
 * 印刷の範囲・倍率・ページの切れ目を決める（ADR-0010 §2-2・§2-5・§3）。
 * - 印刷範囲が無ければ、行は データ・図の最下行まで、列は データ・図の最右列・はみ出しの届く列（見える全部の行から）まで。
 * - 印刷範囲があれば、その範囲を定義の順に使う（範囲の外へのはみ出しは広げない）。列だけ（`$A:$C`）の範囲は
 *   データ・図の最下行まで、行だけ（`$1:$5`）の範囲はデータ・図の最右列・はみ出しの届く列までに縮める。
 * - 見える行を上から数え、全部の範囲を通して見える行 × 見える列が `maxCells` を超える手前で打ち切る
 *   （最初の 1 行は必ず描く。打ち切った後の範囲は描かない）。
 */
export declare function planPrint(kit: PrintKit, sheet: PrintSheet, boxes: readonly DrawingBox[], opts: ResolvedPrintOptions, printArea?: readonly CellRange[]): PrintPlan;
/**
 * 寸法を 1 ページの窓にする：範囲の外の行・列は高さ・幅 0、位置は範囲の左上を 0 とする。
 * 表・罫線の部品にそのまま渡すと、範囲の外の余白を作らずに範囲だけを描く。
 */
export declare function windowGeometry(g: Geometry, p: PrintPage): Geometry;
/** 印刷の描画に使う、表示中のシートの状態（sheet-view.ts が渡す） */
export interface PrintSource {
    /** `sheets` は印刷範囲（`SheetInfo.printArea`）を引くのに使う */
    book: Pick<XlsxBook, 'readImage' | 'readChart' | 'sheets'>;
    /** シートの番号（`SheetInfo.index`） */
    index: number;
    sheet: XlsxSheet;
    /** 図（図を出さない指定なら空） */
    drawings: readonly SheetDrawing[];
    /** 配置済みの図（A1 の左上基準・100% の px） */
    boxes: readonly DrawingBox[];
    messages: Readonly<XSheetMessages>;
    nonce?: string;
}
/**
 * 表示中のシートを、印刷のためにページに分けて `target` の中に描く（ADR-0010 §2）。
 * @throws RangeError 正しくない指定
 * @throws DOMException AbortError（中断。足した物は片付けてから拒否する）
 */
export declare function renderPrint(kit: PrintKit, target: HTMLElement, src: PrintSource | null, options?: PrintOptions): Promise<PrintResult>;
