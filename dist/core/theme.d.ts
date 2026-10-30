/**
 * ブックのテーマの 12 色（`xl/theme/theme1.xml` の `a:clrScheme`）の読み取り。
 * テーマの書式（`a:fmtScheme`）・フォント（`a:fontScheme`）は読まない。
 * テーマが無い・読めない・上限を超える・色が欠けているときは、Office 2013〜2022 の既定の 12 色で補う
 * （`XlsxError` にもバナーにもしない）。テーマはブックに 1 つで、styles.xml と同じくブックで 1 回だけ読む（index.ts）。
 */
import type { XlsxLimits } from './limits';
import type { XlsxPackage } from './open';
import { type Attrs } from './xml';
/** テーマの色の位置（`a:clrScheme` の子の要素名）。この順で 12 色 */
export declare const THEME_SLOTS: readonly ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
export type ThemeSlot = (typeof THEME_SLOTS)[number];
/** テーマの 12 色（位置 → RGB の整数）。`Map` なので位置の名前が `__proto__` などでも汚染されない。常に 12 色そろう */
export type Theme = ReadonlyMap<ThemeSlot, number>;
/** 既定のテーマ（共有する。変更しない） */
export declare const DEFAULT_THEME: Theme;
/** 16 進 6 桁（`[0-9A-Fa-f]{6}`）だけを整数にする。それ以外・未指定は null */
export declare function parseHex6(s: string | undefined): number | null;
/** `sysClr` の色：`lastClr`（16 進 6 桁）。無い・不正なら `windowText`＝黒・`window`＝白、その他は null */
export declare function sysClrValue(attrs: Attrs): number | null;
/**
 * `theme1.xml` の `a:clrScheme` から 12 色を読む。欠けた色・不正な色はその色だけ既定で補う。
 * xml が null・`clrScheme` が無い・閉じが無いときは既定の 12 色（新しい `Map`）。
 */
export declare function parseTheme(xml: string | null): Theme;
/**
 * テーマを読む。`path` が無い・部品が zip に無い・上限（小さな部品）を超える・読めないときは既定の 12 色。
 * 画面を閉じた後の読み取り（`broken`・`closed`）だけは呼び出し側へ投げる（index.ts が閉じたかを確かめる）。
 */
export declare function loadTheme(pkg: XlsxPackage, path: string | null, limits: XlsxLimits): Promise<Theme>;
