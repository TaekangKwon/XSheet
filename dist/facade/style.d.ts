/** テーマ。`auto` は `prefers-color-scheme` に従う */
export type XSheetTheme = 'light' | 'dark' | 'auto';
/** container の高さが 0 のときに付ける高さ */
export declare const DEFAULT_HEIGHT = "480px";
/** 表示の CSS が無ければ入れる（読み込み済み・入れ済みなら何もしない） */
export declare function ensureStyle(container: HTMLElement, nonce?: string): void;
/** テーマを表示の根に付ける（`light` は何も付けず、使う側の `--xs-*` の上書きがそのまま効く） */
export declare function applyTheme(root: HTMLElement, theme: XSheetTheme | undefined): void;
/**
 * 表示の根の高さを決める。
 * - 数値は px、文字列は CSS の長さとしてそのまま付ける。
 * - `auto`（既定）は container の高さが 0 のときだけ {@link DEFAULT_HEIGHT} を付ける（それ以外は style.css の 100%）。
 *   container が描かれていないときは、描かれたときに決める。
 * @returns 後始末（描かれるのを待つ ResizeObserver を外す）
 */
export declare function applyHeight(container: HTMLElement, root: HTMLElement, height: number | string | undefined): () => void;
