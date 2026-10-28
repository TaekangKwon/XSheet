import { describe, expect, it, vi } from 'vitest';
import type { ImageResult } from '../../src/core/index';
import { createImageStore } from '../../src/render/image-store';

/** 成功の読み取り結果 */
const okResult = (mime = 'image/png'): ImageResult => ({
  ok: true,
  format: 'png',
  mime,
  width: 40,
  height: 30,
  bytes: new Uint8Array([1, 2, 3]),
});

/** 外から解決できる Promise */
function deferred<T>() {
  let resolve!: (v: T) => void;
  const promise = new Promise<T>((r) => {
    resolve = r;
  });
  return { promise, resolve };
}

/** URL の作成・解放を数える依存 */
function urlSpy() {
  let n = 0;
  const created: Blob[] = [];
  const revoked: string[] = [];
  return {
    created,
    revoked,
    createUrl: (blob: Blob) => {
      created.push(blob);
      return `blob:test/${++n}`;
    },
    revokeUrl: (url: string) => {
      revoked.push(url);
    },
  };
}

describe('TD-86-4 画像ストア：blob URL の Map（共有・MIME・revoke）', () => {
  it('同じパスは 1 回だけ読み、同じ URL を共有する（複数のアンカーが同じ画像を使う）', async () => {
    const spy = urlSpy();
    const read = vi.fn(async () => okResult());
    const store = createImageStore({ read, ...spy });
    const [a, b] = await Promise.all([
      store.load('xl/media/image1.png'),
      store.load('xl/media/image1.png'),
    ]);
    const c = await store.load('xl/media/image1.png');
    expect(read).toHaveBeenCalledTimes(1);
    expect(a).toEqual({ ok: true, url: 'blob:test/1' });
    expect(b).toBe(a);
    expect(c).toBe(a);
    expect(spy.created).toHaveLength(1);
    expect(store.urlCount).toBe(1);
    expect(store.peek('xl/media/image1.png')).toBe(a);
  });

  it('Blob の type は readImage が判定した MIME（拡張子・Content_Types ではない）', async () => {
    const spy = urlSpy();
    const store = createImageStore({ read: async () => okResult('image/jpeg'), ...spy });
    await store.load('xl/media/whatever.png');
    expect(spy.created[0]!.type).toBe('image/jpeg');
    expect(spy.created[0]!.size).toBe(3);
  });

  it('パスが違えば別の URL（それぞれ 1 回ずつ読む）', async () => {
    const spy = urlSpy();
    const read = vi.fn(async () => okResult());
    const store = createImageStore({ read, ...spy });
    await store.load('a');
    await store.load('b');
    expect(read).toHaveBeenCalledTimes(2);
    expect(store.urlCount).toBe(2);
  });

  it('dispose：作った URL をすべて revoke する（二重呼び出しでも 1 回ずつ）。失敗の結果は URL を持たない', async () => {
    const spy = urlSpy();
    const store = createImageStore({
      read: async (p) =>
        p === 'bad' ? { ok: false, reason: 'unsupported', label: 'EMF' } : okResult(),
      ...spy,
    });
    await store.load('a');
    await store.load('b');
    const bad = await store.load('bad');
    expect(bad).toEqual({ ok: false, reason: 'unsupported', label: 'EMF' });
    store.dispose();
    store.dispose();
    expect(spy.revoked.sort()).toEqual(['blob:test/1', 'blob:test/2']);
    expect(store.urlCount).toBe(0);
    expect(store.peek('a')).toBeUndefined();
  });

  it('破棄後の load は読まず null（URL を作らない）', async () => {
    const spy = urlSpy();
    const read = vi.fn(async () => okResult());
    const store = createImageStore({ read, ...spy });
    store.dispose();
    expect(await store.load('a')).toBeNull();
    expect(read).not.toHaveBeenCalled();
    expect(spy.created).toHaveLength(0);
  });

  it('読み込み中に dispose されたら、終わっても URL を作らず null（シート・ファイルの切替で古い結果を捨てる）', async () => {
    const spy = urlSpy();
    const d = deferred<ImageResult>();
    const store = createImageStore({ read: () => d.promise, ...spy });
    const pending = store.load('a');
    store.dispose();
    d.resolve(okResult());
    expect(await pending).toBeNull();
    expect(spy.created).toHaveLength(0);
    expect(spy.revoked).toHaveLength(0);
    expect(store.urlCount).toBe(0);
  });

  it('readImage が例外を投げても例外にしない（broken の失敗）', async () => {
    const store = createImageStore({
      read: async () => {
        throw new Error('closed');
      },
      ...urlSpy(),
    });
    expect(await store.load('a')).toEqual({ ok: false, reason: 'broken', label: null });
  });

  it('over-budget は初めて知ったとき 1 回だけ通知する。破棄後の結果では通知しない', async () => {
    const onOverBudget = vi.fn();
    const store = createImageStore({
      read: async () => ({ ok: false, reason: 'over-budget', label: null }),
      onOverBudget,
      ...urlSpy(),
    });
    await store.load('a');
    await store.load('b');
    expect(onOverBudget).toHaveBeenCalledTimes(1);

    const late = vi.fn();
    const d = deferred<ImageResult>();
    const s2 = createImageStore({ read: () => d.promise, onOverBudget: late, ...urlSpy() });
    const p = s2.load('a');
    s2.dispose();
    d.resolve({ ok: false, reason: 'over-budget', label: null });
    await p;
    expect(late).not.toHaveBeenCalled();
  });

  it('失敗（too-large など）は over-budget の通知をしない', async () => {
    const onOverBudget = vi.fn();
    const store = createImageStore({
      read: async () => ({ ok: false, reason: 'too-large', label: null }),
      onOverBudget,
      ...urlSpy(),
    });
    expect(await store.load('a')).toEqual({ ok: false, reason: 'too-large', label: null });
    expect(onOverBudget).not.toHaveBeenCalled();
  });
});

describe('TD-86-16 画像ストア：1 本ずつの待ち行列と dispose での破棄', () => {
  it('読み込みは 1 本ずつ：前の展開が終わるまで次の read を始めない', async () => {
    const d1 = deferred<ImageResult>();
    const started: string[] = [];
    const store = createImageStore({
      read: (path) => {
        started.push(path);
        return path === 'a' ? d1.promise : Promise.resolve(okResult());
      },
      ...urlSpy(),
    });
    const pa = store.load('a');
    const pb = store.load('b');
    await Promise.resolve();
    expect(started).toEqual(['a']);
    d1.resolve(okResult());
    await Promise.all([pa, pb]);
    expect(started).toEqual(['a', 'b']);
  });

  it('dispose で未開始の読み込みは捨てる：read を呼ばず null。実行中の 1 本は終わっても URL を作らない', async () => {
    const spy = urlSpy();
    const d1 = deferred<ImageResult>();
    const started: string[] = [];
    const store = createImageStore({
      read: (path) => {
        started.push(path);
        return d1.promise;
      },
      ...spy,
    });
    const results = [store.load('a'), store.load('b'), store.load('c')];
    await Promise.resolve();
    store.dispose();
    d1.resolve(okResult());
    expect(await Promise.all(results)).toEqual([null, null, null]);
    expect(started).toEqual(['a']); // b・c は read されない
    expect(spy.created).toHaveLength(0);
  });
});
