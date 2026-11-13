import { describe, expect, it } from 'vitest';
import { openXlsx, type OpenXlsxOptions } from '../../src/core/index';
import { formatNumberByCode } from '../../src/core/format';
import { prepareJaFormat, tokenizeFormat } from '../../src/core/format-ja';
import { read } from '../../src/facade/view';
import { buildXlsx, type CellSpec } from '../../src/testing/make-xlsx';

/** 和暦の字句（g・e・r）と builtinDates（M7-8・ADR-0007 §6）。Excel との照合は ref-era.test.ts */

const kinds = (code: string) =>
  tokenizeFormat(code).map((s) =>
    s.filter((t) => t.kind !== 'char').map((t) => `${t.kind}:${t.text}`),
  );

describe('和暦の字句の分解', () => {
  it('g・e・r の並び（大文字も）を拾う', () => {
    expect(kinds('[$-411]ggge"年"')).toEqual([
      ['bracket:[$-411]', 'era:ggg', 'eraYear:e', 'quoted:"年"'],
    ]);
    expect(kinds('GE rr')).toEqual([['era:G', 'eraYear:E', 'eraR:rr']]);
  });

  it('指数の E+・E-、General、引用符・角かっこの中は和暦にしない', () => {
    expect(kinds('0.00E+00;0.0e-0')).toEqual([[], []]);
    expect(kinds('General')).toEqual([['general:General']]);
    expect(kinds('"ggg"[Red]')).toEqual([['quoted:"ggg"', 'bracket:[Red]']]);
  });

  it('数値の書式は準備しない（ssf にそのまま渡す）', () => {
    for (const code of ['0.00E+00', 'General', '#,##0;[Red]-#,##0', '"Yen"0', '[$-409]#,##0'])
      expect(prepareJaFormat(code)).toBeNull();
    expect(formatNumberByCode('0.00E+00', 0.000123, false)).toBe('1.23E-04');
    expect(formatNumberByCode('General', 1234567.891, false)).toBe('1234567.891');
  });
});

describe('和暦の表示', () => {
  const show = (code: string, v: number, d1904 = false) => formatNumberByCode(code, v, d1904);

  it('依頼のファイルの書式 [$-ja-JP]ggge"年"m"月"d"日"', () => {
    expect(show('[$-ja-JP]ggge"年"m"月"d"日"', 46337)).toBe('令和8年11月11日');
  });

  it('元号の長さ（g＝記号・gg＝略称・ggg 以上＝名前）と年（e・ee）', () => {
    expect(show('[$-411]ge.m.d', 43586)).toBe('R1.5.1');
    expect(show('[$-411]gge.m.d', 32516)).toBe('平1.1.8');
    expect(show('[$-411]gggee', 9856)).toBe('昭和01');
    expect(show('[$-411]r', 46337)).toBe('08');
    expect(show('[$-411]rr', 46337)).toBe('令和08');
  });

  it('元年は [$-ja-JP-x-gannen] で 2 文字以上の元号の後だけ', () => {
    expect(show('[$-ja-JP-x-gannen]ggge"年"', 43586)).toBe('令和元年');
    expect(show('[$-ja-JP-x-gannen]gge"年"', 4595)).toBe('大元年');
    expect(show('[$-ja-JP-x-gannen]ge"年"', 43586)).toBe('R1年');
    expect(show('[$-ja-JP-x-gannen]e"年"', 43586)).toBe('1年');
  });

  it('日本語でない言語の指定は元号を空に、年は西暦。指定が無く元号の無い e は西暦のまま', () => {
    expect(show('[$-409]ggge', 46337)).toBe('2026');
    expect(show('e', 46337)).toBe('2026');
    expect(show('ggge', 46337)).toBe('令和8');
  });

  it('1904 年基準・日付の範囲の外（空）', () => {
    expect(show('[$-411]ggge"年"m"月"d"日"', 44875, true)).toBe('令和8年11月11日');
    expect(show('[$-411]ggge', -1)).toBe('');
    expect(show('[$-411]ggge', 2958466)).toBe('');
  });

  it('曜日と一緒に使える', () => {
    expect(show('[$-411]ggge"年"m"月"d"日"(aaa)', 46337)).toBe('令和8年11月11日(水)');
  });
});

describe('builtinDates', () => {
  /** 組み込みの番号 28・27・34 の 3 セル（2026-11-11） */
  const bytes = () => {
    const xfs = [0, 28, 27, 34];
    const row: CellSpec[] = [1, 2, 3].map((s) => ({ v: 46337, s }));
    return buildXlsx({ sheets: [{ name: 'S', rows: [row] }], xfs });
  };
  const texts = async (options?: OpenXlsxOptions) => {
    const book = await openXlsx(await bytes(), options);
    const sheet = await book.openSheet(1);
    const out = sheet.getRows(1, 1)[0]!.cells.map((c) => c?.text);
    book.close();
    return out;
  };

  it("既定は今までの西暦の表、'japanese-era' は日本語版 Excel の表", async () => {
    // 既定の表は今のまま（34 は時刻の表。日本語版 Excel とは違うが、既定は変えない。ADR-0007 §6）
    expect((await texts()).slice(0, 2)).toEqual(['2026/11/11', '2026/11/11']);
    expect(await texts({ builtinDates: 'japanese-era' })).toEqual([
      '令和8年11月11日',
      'R8.11.11',
      '2026年11月',
    ]);
  });

  it('read（xsheet の入口）にも渡せる', async () => {
    const book = await read(await bytes(), { builtinDates: 'japanese-era' });
    const sheet = await book.openSheet(1);
    expect(sheet.getRows(1, 1)[0]!.cells[0]?.text).toBe('令和8年11月11日');
    book.close();
  });
});
