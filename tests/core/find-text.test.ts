import { describe, expect, it } from 'vitest';
import { findText, openXlsx, type TextMatch, type XlsxSheet } from '../../src/core/index';
import { emptyLayout } from '../../src/core/internal';
import {
  buildXlsx,
  type CellSpec,
  type SheetSpec,
  type XlsxSpec,
} from '../../src/testing/make-xlsx';

/** 1 シートの xlsx を組んで開き、シートを返す */
async function sheetOf(
  rows: CellSpec[][],
  extra: Partial<SheetSpec> = {},
  book: Partial<XlsxSpec> = {},
): Promise<XlsxSheet> {
  const opened = await openXlsx(
    await buildXlsx({ sheets: [{ name: 'S', rows, ...extra }], ...book }),
  );
  return opened.openSheet(1);
}

/** 一致を「行,列:始まり-終わり」の並びにする */
const brief = (ms: readonly TextMatch[]) => ms.map((m) => `${m.r},${m.c}:${m.start}-${m.end}`);

describe('findText：照合のルール', () => {
  it('一部の一致を、行 → 列 → セルの中の位置の順にすべて返す（重ならない）', async () => {
    const sheet = await sheetOf([
      ['abcab', 'x'],
      ['zz', 'aaa ab'],
    ]);
    const { matches, truncated } = await findText(sheet, 'ab');
    expect(brief(matches)).toEqual(['1,1:0-2', '1,1:3-5', '2,2:4-6']);
    expect(truncated).toBe(false);
    expect(brief((await findText(sheet, 'aa')).matches)).toEqual(['2,2:0-2']);
  });

  it('大文字と小文字は既定で区別しない。全角と半角はいつも区別する', async () => {
    const sheet = await sheetOf([['Apple', 'ＡＰＰＬＥ', 'apple']]);
    expect(brief((await findText(sheet, 'APPLE')).matches)).toEqual(['1,1:0-5', '1,3:0-5']);
    expect(brief((await findText(sheet, 'APPLE', { ignoreCase: false })).matches)).toEqual([]);
    expect(brief((await findText(sheet, 'ａｐｐｌｅ')).matches)).toEqual(['1,2:0-5']);
  });

  it('小文字で長さの変わる文字（İ）・語末のシグマ・サロゲートペアでも元の位置を返す', async () => {
    const sheet = await sheetOf([['İx', 'ΟΔΟΣ', '𠮷野家']]);
    expect(brief((await findText(sheet, 'x')).matches)).toEqual(['1,1:1-2']);
    expect(brief((await findText(sheet, 'İ')).matches)).toEqual(['1,1:0-1']);
    expect(brief((await findText(sheet, 'σ')).matches)).toEqual(['1,2:3-4']);
    expect(brief((await findText(sheet, '野')).matches)).toEqual(['1,3:2-3']);
  });

  it('改行は既定で照合のときだけ除き、位置は元の文字で返す', async () => {
    const sheet = await sheetOf([['ab\ncd', 'ab\r\ncd', 'b\nc']]);
    expect(brief((await findText(sheet, 'bc')).matches)).toEqual(['1,1:1-4', '1,2:1-5', '1,3:0-3']);
    expect(brief((await findText(sheet, 'bc', { ignoreNewlines: false })).matches)).toEqual([]);
    expect(brief((await findText(sheet, 'b\nc', { ignoreNewlines: false })).matches)).toEqual([
      '1,1:1-4',
      '1,3:0-3',
    ]);
    // 検索語の改行も除く
    expect(brief((await findText(sheet, 'b\nc')).matches)).toHaveLength(3);
  });

  it("match: 'cell' はセルの文字全体が一致するものだけ（位置はセルの文字全体）", async () => {
    const sheet = await sheetOf([['abc', 'ABC', 'abcd', 'a\nbc']]);
    expect(brief((await findText(sheet, 'abc', { match: 'cell' })).matches)).toEqual([
      '1,1:0-3',
      '1,2:0-3',
      '1,4:0-4',
    ]);
  });

  it('空の検索語（改行だけを含む）は空の結果', async () => {
    const sheet = await sheetOf([['a']]);
    expect(await findText(sheet, '')).toEqual({ matches: [], truncated: false });
    expect(await findText(sheet, '\r\n')).toEqual({ matches: [], truncated: false });
    expect(await findText(await sheetOf([]), 'a')).toEqual({ matches: [], truncated: false });
  });

  it('表示形式を当てた後の文字で照らす', async () => {
    const sheet = await sheetOf(
      [
        [
          { v: 1234567, s: 1 },
          { v: 45607, s: 2 },
        ],
      ],
      {},
      {
        numFmts: { 164: '#,##0', 165: 'm/d(aaa)' },
        xfs: [0, 164, 165],
      },
    );
    expect(brief((await findText(sheet, '4,5')).matches)).toEqual(['1,1:4-7']);
    expect(brief((await findText(sheet, '(月)')).matches)).toEqual(['1,2:5-8']);
    expect(brief((await findText(sheet, '1234567')).matches)).toEqual([]);
  });
});

describe('findText：結合・非表示', () => {
  it('結合は左上のセルだけ（隠れるセルに残った値は探さない）', async () => {
    const sheet = await sheetOf([['key', 'key'], ['key']], { merges: ['A1:B2'] });
    expect(brief((await findText(sheet, 'key')).matches)).toEqual(['1,1:0-3']);
  });

  it('読む区切り（500 行）をまたぐ結合も 1 回だけ数える', async () => {
    const rows: CellSpec[][] = Array.from({ length: 600 }, () => [null, null]);
    rows[498] = ['X', null]; // A499:A503
    rows[499] = [null, 'X']; // B500:B502（区切りの最後の行が左上）
    rows[500] = [null, null];
    rows[504] = ['X', 'X'];
    const sheet = await sheetOf(rows, { merges: ['A499:A503', 'B500:B502'] });
    expect(brief((await findText(sheet, 'x')).matches)).toEqual([
      '499,1:0-1',
      '500,2:0-1',
      '505,1:0-1',
      '505,2:0-1',
    ]);
  });

  it('非表示の行・列は既定で探さない（skipHidden: false で探す）', async () => {
    const sheet = await sheetOf(
      [
        ['a', 'a', 'a'],
        ['a', 'a', 'a'],
      ],
      { cols: [{ min: 2, max: 2, hidden: true }], rowAttrs: { 2: { hidden: true } } },
    );
    expect(brief((await findText(sheet, 'a')).matches)).toEqual(['1,1:0-1', '1,3:0-1']);
    expect((await findText(sheet, 'a', { skipHidden: false })).matches).toHaveLength(6);
  });

  it('高さ 0 の行・幅 0 の列も非表示と同じ', async () => {
    const sheet = await sheetOf(
      [
        ['a', 'a'],
        ['a', 'a'],
      ],
      {
        cols: [{ min: 1, max: 1, width: 0 }],
        rowAttrs: { 1: { ht: 0 } },
      },
    );
    expect(brief((await findText(sheet, 'a')).matches)).toEqual(['2,2:0-1']);
  });
});

describe('findText：上限・中断・時間', () => {
  it('maxMatches で打ち切り truncated を立てる（ちょうどなら立てない）', async () => {
    const sheet = await sheetOf([['aaaa', 'aa']]);
    expect(await findText(sheet, 'a', { maxMatches: 3 })).toMatchObject({ truncated: true });
    expect((await findText(sheet, 'a', { maxMatches: 3 })).matches).toHaveLength(3);
    expect(await findText(sheet, 'a', { maxMatches: 6 })).toMatchObject({ truncated: false });
    expect(await findText(sheet, 'a', { maxMatches: 0 })).toEqual({ matches: [], truncated: true });
    // 数でない値は既定（10,000）
    expect(await findText(sheet, 'a', { maxMatches: Number.NaN })).toMatchObject({
      truncated: false,
    });
  });

  it('1 区切りが遅いときは区切りの行数を減らし、同じ行を 2 回数えない', async () => {
    // getRows が 1 行 0.05ms かかるシート（500 行で 25ms。区切りの行数が 500 → 250 → 125 と減る）
    const rowCount = 3000;
    const reads: number[] = [];
    const sheet = {
      rowCount,
      colCount: 1,
      layout: emptyLayout(),
      borders: [],
      formats: [],
      rowFormat: () => undefined,
      colFormat: () => -1,
      getRows(first: number, last: number) {
        reads.push(last - first + 1);
        const until = performance.now() + (last - first + 1) * 0.05;
        while (performance.now() < until);
        return Array.from({ length: last - first + 1 }, (_, i) => ({
          r: first + i,
          cells: [{ text: 'x', colSpan: 1, rowSpan: 1 }],
        }));
      },
    } as XlsxSheet;
    const { matches } = await findText(sheet, 'x');
    expect(matches.map((m) => m.r)).toEqual(Array.from({ length: rowCount }, (_, i) => i + 1));
    expect(reads[0]).toBe(500);
    expect(Math.min(...reads)).toBeLessThan(500);
  });

  it('中断済みの signal はすぐに AbortError で reject する', async () => {
    const sheet = await sheetOf([['a']]);
    const ac = new AbortController();
    ac.abort();
    await expect(findText(sheet, 'a', { signal: ac.signal })).rejects.toMatchObject({
      name: 'AbortError',
    });
    await expect(findText(sheet, 'a', { signal: ac.signal })).rejects.toBeInstanceOf(DOMException);
  });

  describe('5000 行 × 30 列の合成シート', () => {
    const big = (async () => {
      const rows: CellSpec[][] = Array.from({ length: 5000 }, (_, r) =>
        Array.from({ length: 30 }, (_, c) => `r${r + 1}c${c + 1} 見積 sample`),
      );
      return sheetOf(rows);
    })();

    it('件数と時間（全部で 150,000 セル）', async () => {
      const sheet = await big;
      const t0 = performance.now();
      const res = await findText(sheet, '見積', { maxMatches: 100_000 });
      const ms = performance.now() - t0;
      expect(res).toMatchObject({ truncated: true });
      expect(res.matches).toHaveLength(100_000);
      const all = await findText(sheet, 'r4999c3');
      expect(brief(all.matches)).toEqual(['4999,3:0-7', '4999,30:0-7']);
      const single = await findText(sheet, 'r1234c5 見積 sample', { match: 'cell' });
      expect(brief(single.matches)).toEqual(['1234,5:0-17']);
      expect(ms).toBeLessThan(5000);
    });

    it('検索の途中で休む（先に積んだタイマーが検索の終わりより前に動く）', async () => {
      const sheet = await big;
      const order: string[] = [];
      const p = findText(sheet, 'zzz').then(() => order.push('done'));
      setTimeout(() => order.push('timer'), 0);
      await p;
      expect(order).toEqual(['timer', 'done']);
    });

    it('検索の途中の中断で AbortError', async () => {
      const sheet = await big;
      const ac = new AbortController();
      const p = findText(sheet, 'zzz', { signal: ac.signal });
      setTimeout(() => ac.abort(), 0);
      await expect(p).rejects.toMatchObject({ name: 'AbortError' });
    });
  });
});
