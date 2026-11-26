// 正: docs/adr/0010-scroll-element-and-print.md（§2）
/**
 * 印刷の描画の口（XSheet の追加。M7-20・M7-21）。中身（`print.ts`）は印刷するときだけ動的 import で読む
 * （render の大きさの上限のため。グラフの描画 `chart-draw` と同じ形。単独版の ESM では同じフォルダの `xsheet-print.mjs`）。
 * - `PRINT_KIT`：印刷が使う表示と共有の部品。`print.ts` はこれを受け取り、表示の部品を自分で import しない
 *   （単独版の部品は部品ごとに別にビルドするので、import すると xsheet-print.mjs に表示の部品が二重に入る）。
 * - `createPrintGate`：表示 1 つの口。表示の destroy・シートの切り替えで、描いている途中の印刷を中断する（sheet-view.ts が持つ）。
 */
import { colName, createGeometry } from '../core/index';
import { createBorderLayer } from './border-layer';
import { buildBorderPaths } from './borders';
import { drawingExtent, rectsIntersect } from './drawings';
import { buildFormatCss } from './format-css';
import {
  applyMarginFormats,
  collectVisibleRows,
  GRID_HEADER_HEIGHT,
  MAX_GRID_COLS,
  MAX_GRID_ROWS,
  rowNumWidth,
} from './grid';
import { buildGridTable } from './grid-table';
import { createImageLayer } from './image-layer';
import type { PrintKit, PrintOptions, PrintResult, PrintSource } from './print';
import { spillReachCols } from './spill-extent';
import { zoomGeometry } from './zoom';

/** 印刷の描画に渡す、表示と共有の部品 */
export const PRINT_KIT: PrintKit = {
  colName,
  createGeometry,
  createBorderLayer,
  buildBorderPaths,
  drawingExtent,
  rectsIntersect,
  buildFormatCss,
  applyMarginFormats,
  collectVisibleRows,
  rowNumWidth,
  buildGridTable,
  createImageLayer,
  spillReachCols,
  zoomGeometry,
  GRID_HEADER_HEIGHT,
  MAX_GRID_ROWS,
  MAX_GRID_COLS,
};

/**
 * 印刷の描画（`print.ts` の `renderPrint`）を読んで呼ぶ。読めなかったとき（単独版で xsheet-print.mjs を置いていない等）は、
 * 読み込みの例外で拒否する。
 */
export async function renderPrintLazy(
  target: HTMLElement,
  src: PrintSource | null,
  options?: PrintOptions,
): Promise<PrintResult> {
  const { renderPrint } = await import('./print');
  return renderPrint(PRINT_KIT, target, src, options);
}

/** 表示 1 つの印刷の口 */
export interface PrintGate {
  run(target: HTMLElement, src: PrintSource | null, options?: PrintOptions): Promise<PrintResult>;
  /** 描いている途中の印刷をすべて中断する */
  abortAll(): void;
}

/** {@link PrintGate} を作る */
export function createPrintGate(): PrintGate {
  const running = new Set<AbortController>();
  return {
    async run(target, src, options = {}) {
      const ctl = new AbortController();
      const user = options.signal;
      const relay = (): void => ctl.abort();
      if (user?.aborted) {
        ctl.abort();
      } else {
        user?.addEventListener('abort', relay, { once: true });
      }
      running.add(ctl);
      try {
        return await renderPrintLazy(target, src, { ...options, signal: ctl.signal });
      } finally {
        running.delete(ctl);
        user?.removeEventListener('abort', relay);
      }
    },
    abortAll() {
      for (const ctl of running) {
        ctl.abort();
      }
      running.clear();
    },
  };
}
