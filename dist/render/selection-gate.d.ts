import type { SelectionController, SelectionHost, SelectionMode } from './selection';
/** 表示 1 つの選択の口 */
export interface SelectionGate extends SelectionController {
    /** 中身を読み終える（読めなくても解決する）。選択の無い表示ではすぐ解決する */
    readonly ready: Promise<void>;
}
/** {@link SelectionGate} を作る。`mode` が無ければ何も読まない */
export declare function createSelectionGate(host: Omit<SelectionHost, 'colName' | 'createGeometry'>, mode: SelectionMode | undefined): SelectionGate;
