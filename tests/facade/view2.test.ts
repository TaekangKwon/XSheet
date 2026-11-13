// @vitest-environment jsdom
import { afterEach, beforeAll, beforeEach, describe, expect, it } from 'vitest';
import { view2, type Viewer2 } from '../../src/facade/view2';
import { buildXlsx } from '../../src/testing/make-xlsx';
import { stubClientSize, wait } from '../render/render-sheet.helpers';

/**
 * 確認用の `view2`（src/facade/view2.ts・ADR-0007 の未決の 2 つ）。
 * 1. Viewer に 0.5.0 の操作（find・ハイライト・倍率・今のシート・onSheetChange）を足したもの。
 * 2. 組み込みの日付の書式の既定が日本語版 Excel の表（builtinDates の既定 'japanese-era'）。
 */

let book: Uint8Array;
const viewers: Viewer2[] = [];
let restoreSize: () => void;

beforeAll(async () => {
  book = await buildXlsx({
    sheets: [
      {
        name: 'One',
        rows: [['apple', 'banana'], ['Apple pie'], [{ v: 46337, s: 1 }]],
        cols: [{ min: 1, max: 1, width: 40 }],
      },
      { name: 'Two', rows: [['cherry apple']] },
    ],
    xfs: [0, 28],
  });
});
beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
});
afterEach(() => {
  viewers.splice(0).forEach((v) => v.destroy());
  document.body.replaceChildren();
  document.head.querySelectorAll('style').forEach((s) => s.remove());
  restoreSize();
});

async function open(
  options?: Parameters<typeof view2>[2],
): Promise<{ el: HTMLElement; v: Viewer2 }> {
  const el = document.createElement('div');
  document.body.append(el);
  const v = await view2(el, book, options);
  viewers.push(v);
  return { el, v };
}

/** 行 r（1 始まり）の A 列の文字（aria-rowindex は見出しの行を 1 と数える） */
const cellA = (el: HTMLElement, r: number): string | null | undefined =>
  el.querySelector(`tbody tr[aria-rowindex="${r + 1}"] td:not(.xs-rownum)`)?.textContent;

describe('view2：組み込みの日付の書式の既定（ADR-0007 §6）', () => {
  it('既定は日本語版 Excel の表（28 を和暦で描く）。default を渡すと view と同じ西暦', async () => {
    const era = await open();
    const plain = await open({ builtinDates: 'default' });
    await wait();
    expect(cellA(era.el, 3)).toBe('令和8年11月11日');
    expect(cellA(plain.el, 3)).toBe('2026/11/11');
  });
});

describe('view2：0.5.0 の操作（ADR-0007 §7）', () => {
  it('activeSheet は SheetInfo.index（1 始まり）。showSheet は番号・名前で替える', async () => {
    const { v } = await open();
    expect(v.sheets.map((s) => s.index)).toEqual([1, 2]);
    expect(v.activeSheet).toBe(1);
    await v.showSheet('Two');
    expect(v.activeSheet).toBe(2);
    await v.showSheet(1);
    expect(v.activeSheet).toBe(1);
    await expect(v.showSheet('Nope')).rejects.toThrow(RangeError);
  });

  it('find は表示中のシートを探してハイライトを付ける。空の文字で消す', async () => {
    const { el, v } = await open();
    const result = await v.find('apple');
    expect(result.matches.map((m) => [m.r, m.c])).toEqual([
      [1, 1],
      [2, 1],
    ]);
    await wait();
    expect(el.querySelectorAll('span.xs-hit').length).toBe(2);
    expect(el.querySelectorAll('span.xs-hit.is-current').length).toBe(1);
    await v.find('');
    await wait();
    expect(el.querySelectorAll('span.xs-hit').length).toBe(0);
  });

  it('find の指定（match・focus）を渡せる', async () => {
    const { el, v } = await open();
    const result = await v.find('apple', { match: 'cell', focus: false });
    expect(result.matches).toHaveLength(1);
    await wait();
    expect(el.querySelectorAll('span.xs-hit.is-current').length).toBe(0);
  });

  it('シートを替えると onSheetChange を呼ぶ。替えた後の find はそのシートを探す', async () => {
    const changed: number[] = [];
    const { v } = await open({ onSheetChange: (i) => changed.push(i) });
    await v.showSheet(2);
    expect(changed).toEqual([2]);
    const result = await v.find('apple');
    expect(result.matches.map((m) => [m.r, m.c, m.start])).toEqual([[1, 1, 7]]);
  });

  it('探している間に次の find・シートの切り替えが来たら AbortError', async () => {
    const { v } = await open();
    const first = v.find('apple');
    const second = v.find('banana');
    await expect(first).rejects.toMatchObject({ name: 'AbortError' });
    expect((await second).matches).toHaveLength(1);
    const third = v.find('apple');
    void v.showSheet(2);
    await expect(third).rejects.toMatchObject({ name: 'AbortError' });
  });

  it('外の signal で中断できる', async () => {
    const { v } = await open();
    const ac = new AbortController();
    ac.abort();
    await expect(v.find('apple', { signal: ac.signal })).rejects.toMatchObject({
      name: 'AbortError',
    });
  });

  it('倍率・見出しの指定を渡し、setZoom・zoom・cellAt・highlightAt を使える', async () => {
    const { v } = await open({ zoom: 1.5, headers: false });
    expect(v.zoom).toBe(1.5);
    v.setZoom(0.5);
    expect(v.zoom).toBe(0.5);
    expect(v.cellAt(-1000, -1000)).toBeNull();
    expect(v.highlightAt(-1000, -1000)).toBe(-1);
    expect(() => v.scrollToCell(0, 1)).toThrow(RangeError);
  });

  it('destroy の後の find は何も探さず、destroy は何度呼んでもよい', async () => {
    const { el, v } = await open();
    v.destroy();
    v.destroy();
    expect(el.querySelector('.xs-viewer')).toBeNull();
    expect((await v.find('apple')).matches).toEqual([]);
  });
});
