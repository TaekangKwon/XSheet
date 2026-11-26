import type { PrintKit, PrintOptions, PrintResult, PrintSource } from './print';
/** 印刷の描画に渡す、表示と共有の部品 */
export declare const PRINT_KIT: PrintKit;
/**
 * 印刷の描画（`print.ts` の `renderPrint`）を読んで呼ぶ。読めなかったとき（単独版で xsheet-print.mjs を置いていない等）は、
 * 読み込みの例外で拒否する。
 */
export declare function renderPrintLazy(target: HTMLElement, src: PrintSource | null, options?: PrintOptions): Promise<PrintResult>;
/** 表示 1 つの印刷の口 */
export interface PrintGate {
    run(target: HTMLElement, src: PrintSource | null, options?: PrintOptions): Promise<PrintResult>;
    /** 描いている途中の印刷をすべて中断する */
    abortAll(): void;
}
/** {@link PrintGate} を作る */
export declare function createPrintGate(): PrintGate;
