import { describe, expect, it, vi } from 'vitest';
import type { ChartResult } from '../../src/core/index';
import { createChartStore } from '../../src/render/chart-store';
import { barModel } from './chart-fixtures';

/**
 * グラフの読み込みの管理（ISSUE-103・ADR-0031 §決定 11）：パスごとに 1 回・1 つずつ順に・破棄後の結果は返さない。
 */

const ok = (): ChartResult => ({ ok: true, chart: barModel() });
const tick = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

function deferred<T>() {
  let resolve!: (v: T) => void;
  const promise = new Promise<T>((r) => (resolve = r));
  return { promise, resolve };
}

describe('TD-103-20 チャートストア', () => {
  it('同じパスは 1 回だけ読む（読み込み中も読み込み後も共有）。peek は読み込み後だけ返す', async () => {
    const read = vi.fn(async () => ok());
    const store = createChartStore({ read });
    const a = store.load('p');
    const b = store.load('p');
    expect(store.peek('p')).toBeUndefined();
    const [ra, rb] = await Promise.all([a, b]);
    expect(ra).toBe(rb);
    expect(store.peek('p')).toBe(ra);
    await store.load('p');
    expect(read).toHaveBeenCalledTimes(1);
  });

  it('1 つずつ順に読む（先頭の結果が返るまで次を呼ばない）', async () => {
    const d1 = deferred<ChartResult>();
    const order: string[] = [];
    const read = vi.fn((p: string) => {
      order.push(p);
      return p === 'a' ? d1.promise : Promise.resolve(ok());
    });
    const store = createChartStore({ read });
    const pa = store.load('a');
    const pb = store.load('b');
    await tick();
    expect(order).toEqual(['a']);
    d1.resolve(ok());
    await Promise.all([pa, pb]);
    expect(order).toEqual(['a', 'b']);
  });

  it('read が例外・reject でも、例外にせず broken の失敗として覚える', async () => {
    const store = createChartStore({ read: () => Promise.reject(new Error('closed')) });
    expect(await store.load('x')).toEqual({ ok: false, reason: 'broken' });
    expect(store.peek('x')).toEqual({ ok: false, reason: 'broken' });
  });

  it('dispose：未開始の読み込みは read を呼ばず null。読み込み中の結果は捨てて null。以後の load は null・peek は空', async () => {
    const d1 = deferred<ChartResult>();
    const read = vi.fn((p: string) => (p === 'a' ? d1.promise : Promise.resolve(ok())));
    const store = createChartStore({ read });
    const pa = store.load('a');
    const pb = store.load('b');
    await tick();
    store.dispose();
    d1.resolve(ok());
    expect(await pa).toBeNull();
    expect(await pb).toBeNull();
    expect(read).toHaveBeenCalledTimes(1); // b は呼ばれない
    expect(await store.load('c')).toBeNull();
    expect(store.peek('a')).toBeUndefined();
    store.dispose(); // 2 回目は何もしない
  });
});
