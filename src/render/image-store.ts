// 正: InterHub docs/plan/adr/0026-xlsx-image-overlay.md（§決定 9）
/**
 * `.xlsx` の画像の blob URL を、シート 1 枚ぶん管理する（DOM 非依存・ISSUE-86）。
 * - media のパス → 読み込み結果の `Map`。同じ画像を複数のアンカーで共有し、`readImage` は 1 パスにつき 1 回だけ呼ぶ。
 * - 判定済みの MIME を明示した Blob から URL を作る（`<img>` の `src` にだけ使う前提。拡張子・Content_Types は信じない）。
 * - 読み込み（展開）は 1 本ずつ待ち行列で行い、`read` を呼ぶ直前に破棄済みかを確かめる。`dispose()` で未開始の読み込みは捨てる
 *   （シートを切り替えた後も前のシートの画像の展開が走り続けて、新しいシートの画像が待たされるのを防ぐ）。
 * - `dispose()`（タブの切替・ビューの破棄）ですべて revoke する。破棄後に読み込みが終わっても URL は作らない
 *   （読み込み中にタブやファイルが切り替わった場合の古い結果の抑止。`spreadsheet.ts` の `openSeq` と同じ目的を、シート単位の破棄で担う）。
 * 画像のバイト・パスは console に出さない。
 */
import type { ImageFailReason, ImageResult } from '../core/index';

/** 読み込み結果。成功は blob URL、失敗は理由（表示側が破線枠にする）。`label` は EMF など形式名（分かれば） */
export type ImageEntry =
  { ok: true; url: string } | { ok: false; reason: ImageFailReason; label: string | null };

/** 画像ストアの外部依存（テストで差し替える） */
export interface ImageStoreDeps {
  /** 画像を読む（`XlsxBook.readImage` をシート番号込みで束ねたもの） */
  read: (path: string) => Promise<ImageResult>;
  /** Blob から URL を作る（既定 `URL.createObjectURL`） */
  createUrl?: (blob: Blob) => string;
  /** URL を解放する（既定 `URL.revokeObjectURL`） */
  revokeUrl?: (url: string) => void;
  /** 1 シートの画像の合計の上限を超えた（`over-budget`）ことを初めて知ったとき 1 回だけ呼ぶ（バナー用） */
  onOverBudget?: () => void;
}

/** シート 1 枚の画像ストア */
export interface ImageStore {
  /** 読み込み済みの結果（未読み込み・読み込み中は undefined） */
  peek(path: string): ImageEntry | undefined;
  /**
   * 画像を読む。読み込み済み・読み込み中のパスは共有する。破棄済み（または読み込み中に破棄された）なら null。
   * 例外にはしない（読めなければ `reason: 'broken'` の失敗）。
   */
  load(path: string): Promise<ImageEntry | null>;
  /** 作った URL をすべて revoke し、以後の読み込み結果を捨てる */
  dispose(): void;
  /** 保持している blob URL の数（テスト用） */
  readonly urlCount: number;
}

/** 画像ストアを作る */
export function createImageStore(deps: ImageStoreDeps): ImageStore {
  const createUrl = deps.createUrl ?? ((blob: Blob) => URL.createObjectURL(blob));
  const revokeUrl = deps.revokeUrl ?? ((url: string) => URL.revokeObjectURL(url));
  const done = new Map<string, ImageEntry>();
  const pending = new Map<string, Promise<ImageEntry | null>>();
  let disposed = false;
  let overBudgetNotified = false;

  /** 未開始の読み込み（先頭から 1 本ずつ実行する）。dispose で破棄する */
  const queue: { path: string; resolve: (entry: ImageEntry | null) => void }[] = [];
  let busy = false;

  const start = async (path: string): Promise<ImageEntry | null> => {
    if (disposed) {
      return null;
    }
    let result: ImageResult;
    try {
      result = await deps.read(path);
    } catch {
      // 読めなかった（ブックが閉じられた等）。理由の詳細は出さず、描けない画像として扱う
      result = { ok: false, reason: 'broken', label: null };
    }
    if (disposed) {
      return null;
    }
    let entry: ImageEntry;
    if (result.ok) {
      // bytes は readImage が作った独立のコピー。TS の Blob 型に合わせて ArrayBuffer へ寄せる
      const blob = new Blob([result.bytes as BlobPart], { type: result.mime });
      entry = { ok: true, url: createUrl(blob) };
    } else {
      entry = { ok: false, reason: result.reason, label: result.label };
      if (result.reason === 'over-budget' && !overBudgetNotified) {
        overBudgetNotified = true;
        deps.onOverBudget?.();
      }
    }
    done.set(path, entry);
    return entry;
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
        p = new Promise<ImageEntry | null>((resolve) => {
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
      for (const entry of done.values()) {
        if (entry.ok) {
          revokeUrl(entry.url);
        }
      }
      done.clear();
      pending.clear();
    },
    get urlCount() {
      let n = 0;
      for (const entry of done.values()) {
        if (entry.ok) n++;
      }
      return n;
    },
  };
}
