/**
 * `.xlsx` のグラフの読み込み結果を、シート 1 枚ぶん管理する（DOM 非依存・ISSUE-103）。`xlsx-image-store.ts` と同じ流儀。
 * - 部品のパス → 読み込み結果（`ChartResult`）の `Map`。同じグラフを複数のアンカーで共有し、`readChart` は 1 パスにつき 1 回だけ呼ぶ。
 * - 読み込みは 1 本ずつ待ち行列で行う（シートの点の合計を呼んだ順に数えるため。ADR-0031 §決定 11）。`read` を呼ぶ直前に破棄済みかを確かめる。
 * - `dispose()`（シートの切替・ビューの破棄）で未開始の読み込みを捨て、結果を持たない。破棄後に読み込みが終わっても結果は返さない（null）
 *   ＝古いシートの結果が新しい表示に入らない。
 * グラフのバイト・パス・名前は console に出さない。
 */
import type { ChartResult } from '../core/index';
/** グラフの読み込みの外部依存 */
export interface ChartStoreDeps {
    /** グラフを読む（`XlsxBook.readChart` をシート番号込みで束ねたもの） */
    read: (path: string) => Promise<ChartResult>;
}
/** シート 1 枚のグラフの読み込み結果 */
export interface ChartStore {
    /** 読み込み済みの結果（未読み込み・読み込み中は undefined） */
    peek(path: string): ChartResult | undefined;
    /**
     * グラフを読む。読み込み済み・読み込み中のパスは共有する。破棄済み（または読み込み中に破棄された）なら null。
     * 例外にはしない（読めなければ `reason: 'broken'` の失敗）。
     */
    load(path: string): Promise<ChartResult | null>;
    /** 結果を捨て、未開始の読み込みを取り消し、以後の結果を捨てる */
    dispose(): void;
}
/** グラフの読み込みの管理を作る */
export declare function createChartStore(deps: ChartStoreDeps): ChartStore;
