/**
 * workbook.xml の読み取り：シートの一覧（workbook の順）・非表示・1904 年起点・各部品の場所。
 * シートの順は非表示も数える（1 始まり・非表示を含む）。
 */
import type { XlsxLimits } from './limits';
import type { XlsxPackage } from './open';
/** workbook の中の 1 シート */
export interface WorkbookSheet {
    /** 1 始まり・workbook の順 */
    index: number;
    name: string;
    hidden: boolean;
    /** zip の中のパス。rels に見つからなければ null */
    path: string | null;
    /** 普通のワークシートか（グラフシートなどは false） */
    isWorksheet: boolean;
}
/** 読んだ workbook */
export interface Workbook {
    sheets: WorkbookSheet[];
    date1904: boolean;
    /** sharedStrings.xml のパス。無ければ null */
    sharedStringsPath: string | null;
    /** styles.xml のパス。無ければ null */
    stylesPath: string | null;
    /** テーマ（theme1.xml）のパス。rels に無ければ null（既定の 12 色を使う） */
    themePath: string | null;
}
/** 関係（rels）1 件 */
export interface Relationship {
    type: string;
    target: string;
    /** `TargetMode="External"`（zip の外を指す。読み取りでは辿らない） */
    external: boolean;
}
/** `rels` を Id → 関係 の Map にする（キーを Map に入れて `__proto__` を防ぐ） */
export declare function parseRels(xml: string): Map<string, Relationship>;
/** 関係の Target を zip のパスにする。`/` 始まりは絶対、それ以外は baseDir からの相対（`..` を解く） */
export declare function resolveTarget(baseDir: string, target: string): string;
/** パスのフォルダ部分（末尾の / を含む。無ければ空） */
export declare function dirOf(path: string): string;
/** パスのファイル名の部分 */
export declare function baseOf(path: string): string;
/** 型（URL の最後の語）が一致するか */
export declare const typeIs: (type: string, last: string) => boolean;
/**
 * workbook.xml を読む。
 * - 本体の場所は `_rels/.rels` の officeDocument から引く（無ければ `xl/workbook.xml`）。
 * - 読めない形式（Strict・xlsb）は unsupported。workbook が無ければ broken。
 */
export declare function loadWorkbook(pkg: XlsxPackage, limits: XlsxLimits): Promise<Workbook>;
