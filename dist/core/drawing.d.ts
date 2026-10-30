import { type XlsxLimits } from './limits';
import type { XlsxPackage } from './open';
import { type Theme } from './theme';
import { type SheetDrawings } from './types';
/**
 * シートの XML（`<sheetData>` の後ろの部分）から `<drawing r:id>` の関係の Id を取り出す。無ければ null。
 * `legacyDrawing`（コメントの吹き出し・フォームコントロール）は別の要素なので拾わない。
 */
export declare function findDrawingRelId(sheetXml: string): string | null;
/**
 * シートの図を読む。`relId` はシートの `<drawing r:id>`。
 * 上限超え・読み取りの失敗は例外にせず、通知（`notices`）で返す。図の無いシート・zip に解決できない参照は空の一覧。
 * @param sheetPath シート XML の zip の中のパス（rels の場所と Target の基準）
 */
export declare function loadSheetDrawings(pkg: XlsxPackage, sheetPath: string, relId: string, limits: XlsxLimits, theme?: Theme): Promise<SheetDrawings>;
