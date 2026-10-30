// 正: InterHub docs/plan/adr/0031-xlsx-charts.md（§決定 11・P1）
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
export function createChartStore(deps: ChartStoreDeps): ChartStore {
  const done = new Map<string, ChartResult>();
  const pending = new Map<string, Promise<ChartResult | null>>();
  let disposed = false;

  const queue: { path: string; resolve: (r: ChartResult | null) => void }[] = [];
  let busy = false;

  const start = async (path: string): Promise<ChartResult | null> => {
    if (disposed) {
      return null;
    }
    let result: ChartResult;
    try {
      result = await deps.read(path);
    } catch {
      // 読めなかった（ブックが閉じられた等）。理由の詳細は出さず、描けないグラフとして扱う
      result = { ok: false, reason: 'broken' };
    }
    if (disposed) {
      return null;
    }
    done.set(path, result);
    return result;
  };

  /** 待ち行列を 1 本ずつ処理する。破棄済みなら `read` を呼ばず null で終える */
  const pump = async (): Promise<void> => {
    if (busy) {
      return;
    }
    busy = true;
    try {
      for (let job = queue.shift(); job; job = queue.shift()) {
        job.resolve(disposed ? null : await start(job.path).catch(() => null));
      }
    } finally {
      busy = false;
    }
  };

  return {
    peek: (path) => done.get(path),
    load(path) {
      if (disposed) {
        return Promise.resolve(null);
      }
      const known = done.get(path);
      if (known) {
        return Promise.resolve(known);
      }
      let p = pending.get(path);
      if (!p) {
        p = new Promise<ChartResult | null>((resolve) => {
          queue.push({ path, resolve });
          void pump();
        }).finally(() => pending.delete(path));
        pending.set(path, p);
      }
      return p;
    },
    dispose() {
      if (disposed) {
        return;
      }
      disposed = true;
      for (const job of queue.splice(0)) {
        job.resolve(null); // 未開始の読み込みは捨てる（read を呼ばない）
      }
      done.clear();
      pending.clear();
    },
  };
}
