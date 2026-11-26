/**
 * cellXfs の `<xf>` の子 `<alignment>`（横・縦の配置と折り返し）の読み取り（M7-12・Moobius の依頼 R9）。
 * - `cell-format.ts`（InterHub の写し）は、xf ごとの書式に {@link readXfAlignments} の結果を足し、まとめるキーに {@link alignKey} を足すだけ。
 * - `applyAlignment` は見ない（塗り・文字の装飾と同じ。F5）。cellStyleXfs は読まない。
 * - 走査は線形：`<xf>`・`</xf>`・`<alignment>` の開きタグを 1 本の正規表現で順に拾う。閉じの無い xf は次の xf で終わったとみなす。
 * DOM・Node 固有の API を使わない。
 */
import type { CellFormat } from './types';
/** 書式のうち配置の項目 */
export type CellAlign = Pick<CellFormat, 'hAlign' | 'vAlign' | 'wrap'>;
/**
 * cellXfs の中身 → xf の位置ごとの配置（配置の無い xf は undefined）。
 * 位置は `<xf` の開きタグの出現順（`cell-format.ts` の数え方と同じ）。
 * @param body `<cellXfs>` の中身
 */
export declare function readXfAlignments(body: string): (CellAlign | undefined)[];
/** 書式をまとめるキーの配置の部分（数値だけ：横 0〜3・縦 0〜2・折り返し 0〜1） */
export declare function alignKey(f: CellAlign): string;
