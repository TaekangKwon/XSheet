import { describe, expect, it } from 'vitest';
import { openXlsx } from '../../src/core/index';
import { alignKey, readXfAlignments } from '../../src/core/cell-align';
import { buildXlsx, fontXml, type CellSpec } from '../../src/testing/make-xlsx';

/** 配置と折り返しの読み取り（M7-12・ADR-0008 §1・Moobius の依頼 R9） */

describe('readXfAlignments', () => {
  it('xf の位置ごとに horizontal・vertical・wrapText を読む', () => {
    const body = [
      '<xf numFmtId="0"/>',
      '<xf numFmtId="0"><alignment horizontal="center"/></xf>',
      '<xf numFmtId="0"><protection locked="1"/></xf>',
      '<xf numFmtId="0" applyAlignment="0"><alignment vertical="top" wrapText="1"/></xf>',
      '<xf><alignment horizontal="general" vertical="bottom" wrapText="0"/></xf>',
    ].join('');
    expect(readXfAlignments(body)).toEqual([
      undefined,
      { hAlign: 'center' },
      undefined,
      { vAlign: 'top', wrap: true },
      undefined,
    ]);
  });

  it('値の対応：distributed・fill・justify は left、centerContinuous は center、縦の justify・distributed は top', () => {
    const one = (attrs: string) => readXfAlignments(`<xf><alignment ${attrs}/></xf>`)[0];
    expect(one('horizontal="right"')).toEqual({ hAlign: 'right' });
    expect(one('horizontal="left"')).toEqual({ hAlign: 'left' });
    expect(one('horizontal="distributed"')).toEqual({ hAlign: 'left' });
    expect(one('horizontal="fill"')).toEqual({ hAlign: 'left' });
    expect(one('horizontal="justify"')).toEqual({ hAlign: 'left' });
    expect(one('horizontal="centerContinuous"')).toEqual({ hAlign: 'center' });
    expect(one('horizontal="__proto__"')).toBeUndefined();
    expect(one('vertical="center"')).toEqual({ vAlign: 'center' });
    expect(one('vertical="justify"')).toEqual({ vAlign: 'top' });
    expect(one('vertical="distributed"')).toEqual({ vAlign: 'top' });
    expect(one('wrapText="true"')).toEqual({ wrap: true });
  });

  it('名前空間の接頭辞・閉じの無い xf（次の xf で終わる）・xf の外の alignment', () => {
    expect(readXfAlignments('<x:xf><x:alignment horizontal="right"/></x:xf><x:xf/>')).toEqual([
      { hAlign: 'right' },
      undefined,
    ]);
    expect(readXfAlignments('<xf><xf><alignment horizontal="right"/></xf>')).toEqual([
      undefined,
      { hAlign: 'right' },
    ]);
    expect(readXfAlignments('<alignment horizontal="right"/><xf/>')).toEqual([undefined]);
    // xfId 等の似た名前は xf に数えない
    expect(readXfAlignments('<xfId/><xf/>')).toEqual([undefined]);
  });

  it('大量の xf でも線形', () => {
    const body = '<xf><alignment horizontal="center"/></xf>'.repeat(50_000);
    const t = performance.now();
    expect(readXfAlignments(body)).toHaveLength(50_000);
    expect(performance.now() - t).toBeLessThan(1000);
  });

  it('alignKey は配置ごとに違う', () => {
    const keys = new Set(
      [
        {},
        { hAlign: 'left' as const },
        { hAlign: 'center' as const },
        { hAlign: 'right' as const },
        { vAlign: 'top' as const },
        { vAlign: 'center' as const },
        { wrap: true as const },
      ].map(alignKey),
    );
    expect(keys.size).toBe(7);
  });
});

describe('CellFormat の配置（openXlsx）', () => {
  const text = (v: string, s: number): CellSpec => ({ v, t: 'inlineStr', s });

  it('配置だけの xf も書式になり、配置の違う書式はまとめない。配置の無い xf は書式なし', async () => {
    const book = await openXlsx(
      await buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: [[text('a', 0), text('b', 1), text('c', 2), text('d', 3), text('e', 4)]],
          },
        ],
        fonts: [fontXml('<sz val="11"/>'), fontXml('<b/>')],
        xfs: [0, 0, 0, 0, 0],
        xfFontIds: [0, 0, 0, 1, 1],
        xfAlignments: [
          undefined,
          'horizontal="center"',
          'horizontal="right" vertical="center" wrapText="1"',
          'horizontal="center"',
          undefined,
        ],
      }),
    );
    const sheet = await book.openSheet(1);
    const row = sheet.getRows(1, 1)[0]!;
    const f = (c: number) => {
      const i = row.formats?.[c];
      return i === undefined ? undefined : sheet.formats[i];
    };
    expect(f(0)).toBeUndefined();
    expect(f(1)).toEqual({ hAlign: 'center' });
    expect(f(2)).toEqual({ hAlign: 'right', vAlign: 'center', wrap: true });
    expect(f(3)).toEqual({ bold: true, hAlign: 'center' });
    expect(f(4)).toEqual({ bold: true });
    expect(new Set([1, 2, 3, 4].map((c) => row.formats?.[c])).size).toBe(4);
    book.close?.();
  });
});
