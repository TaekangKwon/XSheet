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
export type ImageEntry = {
    ok: true;
    url: string;
} | {
    ok: false;
    reason: ImageFailReason;
    label: string | null;
};
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
export declare function createImageStore(deps: ImageStoreDeps): ImageStore;
