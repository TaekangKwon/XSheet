/**
 * SpreadsheetML の色（styles.xml の `<color auto rgb theme tint indexed>`）の解釈（画面に依存しない純関数）。
 * DrawingML の色（color.ts・ADR-0028 §決定 6）とは書き方も `tint` の式も別なので、別の小さな解釈器にする。
 * 罫線（borders.ts）のほか、塗り・フォントの色（次の候補）でも使い回す。
 * - `auto="1"`・属性なし・`indexed="64"` は null（自動。描く側が既定の色にする）。解けない・不正な値も null。例外にしない。
 * - 値は整数（0〜0xFFFFFF）と不透明度で返し、ファイルの文字列を返さない。
 */
import type { Rgba } from './color';
import { type Theme } from './theme';
import { type Attrs } from './xml';
/** 旧 64 色の表（ECMA-376 §18.8.27。`<indexedColors>` が無いブックの `indexed` 0〜63） */
export declare const INDEXED_COLORS: readonly number[];
/**
 * styles.xml の `<colors><indexedColors>` を読み、64 色の表を返す（既定の表を、書いてある位置だけ上書きする）。
 * 無い・空なら null（既定の表を使う）。`<rgbColor rgb>` は 64 件まで。不正な値はその位置だけ既定のまま。
 */
export declare function parseIndexedColors(xml: string): readonly number[] | null;
/**
 * `<color>` の属性 → 色。null は自動（`auto="1"`・属性なし・`indexed="64"`）か解けない値。
 * @param attrs `parseAttrs` で読んだ `<color>` の属性（無ければ undefined＝自動）
 * @param theme ブックのテーマ（`theme` の番号を解く）
 * @param palette `parseIndexedColors` の結果（null なら既定の 64 色）
 */
export declare function resolveCellColor(attrs: Attrs | undefined, theme: Theme, palette: readonly number[] | null): Rgba | null;
