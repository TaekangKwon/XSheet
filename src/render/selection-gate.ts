// 正: docs/adr/0011-selection.md §3
/**
 * 選択の口（XSheet の追加。M7-24）。中身（`selection.ts`：帯・クリック・キーボード・コピー）は、`selection` を指定した表示を
 * 作るときだけ動的 import で読む（render の大きさの上限のため。印刷の `print-gate.ts` と同じ形。単独版の ESM では同じフォルダの
 * `xsheet-selection.mjs`）。表示は最初のシートを出す前に読み終えるのを待つので、`renderSheet`・`renderWorkbook` の約束が
 * 解決した後は、`getSelection`・`setSelection` がすぐ効く。
 * - `selection` を指定しない表示では何も読まず、選択は無い（`getSelection` は null、`setSelection` は何もしない）。
 * - 読めなかったとき（単独版で xsheet-selection.mjs を置いていない等）も、選択の無い表示として動く。
 */
import { colName, createGeometry } from '../core/index';
import type { SelectionController, SelectionHost, SelectionMode } from './selection';

/** 表示 1 つの選択の口 */
export interface SelectionGate extends SelectionController {
  /** 中身を読み終える（読めなくても解決する）。選択の無い表示ではすぐ解決する */
  readonly ready: Promise<void>;
}

/** {@link SelectionGate} を作る。`mode` が無ければ何も読まない */
export function createSelectionGate(
  host: Omit<SelectionHost, 'colName' | 'createGeometry'>,
  mode: SelectionMode | undefined,
): SelectionGate {
  let ctl: SelectionController | null = null;
  let disposed = false;
  const ready = mode
    ? import('./selection').then(
        (m) => {
          if (disposed) return;
          ctl = m.createSelection({ ...host, colName, createGeometry }, mode);
          ctl.layout();
        },
        () => undefined,
      )
    : Promise.resolve();
  return {
    ready,
    get: () => ctl?.get() ?? null,
    set: (input, origin) => ctl?.set(input, origin),
    layout: () => ctl?.layout(),
    text: () => ctl?.text() ?? null,
    clear: () => ctl?.clear(),
    dispose() {
      disposed = true;
      ctl?.dispose();
      ctl = null;
    },
  };
}
