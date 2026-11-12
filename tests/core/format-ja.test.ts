import { describe, expect, it } from 'vitest';
import { openXlsx, type XlsxSheet } from '../../src/core/index';
import {
  applyJaFormat,
  prepareJaFormat,
  tokenizeFormat,
  weekdayOfSerial,
} from '../../src/core/format-ja';
import { formatNumberByCode } from '../../src/core/format';
import { buildXlsx, type CellSpec, type XlsxSpec } from '../../src/testing/make-xlsx';

/** 字句を「種類:文字」の並びにする（見比べやすくするため） */
const kinds = (code: string) =>
  tokenizeFormat(code).map((s) =>
    s.filter((t) => t.kind !== 'char').map((t) => `${t.kind}:${t.text}`),
  );

describe('書式の字句の分解', () => {
  it('曜日の a の並び（3 つ以上）を拾い、1・2 個は文字のまま', () => {
    expect(kinds('m/d(aaa)')).toEqual([['weekday:aaa']]);
    expect(kinds('aaaa')).toEqual([['weekday:aaaa']]);
    expect(kinds('aaaaa')).toEqual([['weekday:aaaaa']]);
    expect(kinds('a aa')).toEqual([[]]);
  });

  it('大文字の AAA も曜日。並びは同じ文字（大文字小文字を区別）で数える', () => {
    expect(kinds('AAA AAAA')).toEqual([['weekday:AAA', 'weekday:AAAA']]);
    expect(kinds('aAa')).toEqual([[]]);
    expect(kinds('Aaaa')).toEqual([['weekday:aaa']]);
  });

  it('引用符・\\・_・*・[...] の中の a は曜日にしない', () => {
    expect(kinds('"aaa"aaa')).toEqual([['quoted:"aaa"', 'weekday:aaa']]);
    expect(kinds('\\aaaa')).toEqual([['escaped:\\a', 'weekday:aaa']]);
    expect(kinds('aaa_a*a')).toEqual([['weekday:aaa', 'escaped:_a', 'escaped:*a']]);
    expect(kinds('[$-ja-JP]aaaa')).toEqual([['bracket:[$-ja-JP]', 'weekday:aaaa']]);
  });

  it('AM/PM・A/P（大文字小文字を問わない）は曜日にしない', () => {
    expect(kinds('aaa h:mm AM/PM')).toEqual([['weekday:aaa', 'ampm:AM/PM']]);
    expect(kinds('aaa am/pm')).toEqual([['weekday:aaa', 'ampm:am/pm']]);
    expect(kinds('aaa a/p')).toEqual([['weekday:aaa', 'ampm:a/p']]);
    // a の並びの途中から a/p が始まるときは、手前までを曜日にする
    expect(kinds('aaaa/p')).toEqual([['weekday:aaa', 'ampm:a/p']]);
  });

  it('; で区画に分ける（引用符・角かっこの中の ; は区切りにしない）', () => {
    expect(kinds('m/d;aaa;"x;y";[<=1;]@')).toEqual([
      [],
      ['weekday:aaa'],
      ['quoted:"x;y"'],
      ['bracket:[<=1;]'],
    ]);
  });

  it('閉じていない引用符・角かっこ・末尾の \\ でも止まらない', () => {
    expect(kinds('aaa"ab')).toEqual([['weekday:aaa', 'quoted:"ab']]);
    expect(kinds('[Red aaa')).toEqual([['bracket:[Red aaa']]);
    expect(kinds('aaa\\')).toEqual([['weekday:aaa', 'escaped:\\']]);
  });

  it('字句を元の順につなぐと元の書式コードに戻る', () => {
    for (const code of ['yyyy"年"m"月"d"日"(aaaa);[Red]\\-0;_ * "-"_;@', 'aaa"ab', '[h]:mm aaa']) {
      expect(
        tokenizeFormat(code)
          .map((s) => s.map((t) => t.text).join(''))
          .join(';'),
      ).toBe(code);
    }
  });

  it('長い書式でも長さに比例した時間で終わる', () => {
    const code = '"a'.repeat(50_000) + 'a\\'.repeat(50_000) + '['.repeat(50_000);
    const t0 = performance.now();
    tokenizeFormat(code);
    expect(performance.now() - t0).toBeLessThan(500);
  });
});

describe('通し番号からの曜日（Excel の WEEKDAY と同じ）', () => {
  // 0＝日曜。1900 年基準は通し番号 1（1900-01-01）を日曜として 60（1900-02-29）も数える
  it('1900 年基準', () => {
    expect(weekdayOfSerial(1, false)).toBe(0);
    expect(weekdayOfSerial(7, false)).toBe(6);
    expect(weekdayOfSerial(59, false)).toBe(2);
    expect(weekdayOfSerial(60, false)).toBe(3);
    expect(weekdayOfSerial(61, false)).toBe(4); // 1900-03-01 は実際に木曜
    expect(weekdayOfSerial(45607, false)).toBe(1); // 2024-11-11 は月曜
    expect(weekdayOfSerial(45607.99, false)).toBe(1);
    expect(weekdayOfSerial(0, false)).toBe(6);
  });

  it('1904 年基準（通し番号 0＝1904-01-01 は金曜）', () => {
    expect(weekdayOfSerial(0, true)).toBe(5);
    expect(weekdayOfSerial(44145, true)).toBe(1); // 2024-11-11
  });
});

describe('曜日の置き換え', () => {
  it('曜日の字句が無い書式は準備しない（そのまま ssf へ）', () => {
    expect(prepareJaFormat('yyyy/m/d')).toBeNull();
    expect(prepareJaFormat('"aaa" AM/PM')).toBeNull();
  });

  it('計算した曜日を引用符で包んで置き換える', () => {
    const p = prepareJaFormat('m/d(aaa);aaaa')!;
    expect(applyJaFormat(p, 45607, false)).toBe('m/d("月");"月曜日"');
    // 1904 年基準の負の値は絶対値の日付（Excel は「-」付きで出す）
    expect(applyJaFormat(p, -44145, true)).toBe('m/d("月");"月曜日"');
    expect(applyJaFormat(p, Number.NaN, false)).toBe('m/d(aaa);aaaa');
    expect(applyJaFormat(prepareJaFormat('AAA aaaaa')!, 45607, false)).toBe('"月" "月曜日"');
  });

  it('日付の範囲の外（Excel が #### にする値）は曜日を空にする', () => {
    const p = prepareJaFormat('aaa')!;
    expect(applyJaFormat(p, -1, false)).toBe('""');
    expect(applyJaFormat(p, 2958466, false)).toBe('""');
    expect(applyJaFormat(p, 2958465, false)).toBe('"金"'); // 9999/12/31
    expect(applyJaFormat(p, 2957004, true)).toBe('""');
    expect(applyJaFormat(p, -2957003, true)).toBe('"金"');
  });

  it('表示する文字の無い言語の指定 [$-411] は外し、通貨の記号付きの [$€-2] は残す', () => {
    expect(applyJaFormat(prepareJaFormat('[$-411]aaa')!, 1, false)).toBe('"日"');
    expect(applyJaFormat(prepareJaFormat('[$-ja-JP]m/d aaaa')!, 1, false)).toBe('m/d "日曜日"');
    expect(applyJaFormat(prepareJaFormat('[$€-2]aaa')!, 1, false)).toBe('[$€-2]"日"');
    expect(formatNumberByCode('[$-411]aaa', 45607, false)).toBe('月');
  });

  it('formatNumberByCode（グラフの文字）でも曜日が出る', () => {
    expect(formatNumberByCode('m/d(aaa)', 45607, false)).toBe('11/11(月)');
    expect(formatNumberByCode('aaaa', 44145, true)).toBe('月曜日');
  });
});

/** 1 シートの xlsx を組んで開き、シートを返す */
async function sheetOf(rows: CellSpec[][], extra: Partial<XlsxSpec> = {}): Promise<XlsxSheet> {
  const book = await openXlsx(await buildXlsx({ sheets: [{ name: 'S', rows }], ...extra }));
  return book.openSheet(1);
}
const firstRow = (sheet: XlsxSheet) => sheet.getRows(1, 1)[0]!.cells.map((c) => c?.text);

describe('セルの表示（曜日の書式）', () => {
  const numFmts = {
    164: 'm/d(aaa)',
    165: 'aaaa',
    166: 'yyyy"年"m"月"d"日"(aaa) h:mm',
    167: 'aaa',
  };
  const xfs = [0, 164, 165, 166, 167];
  const x = (id: number, v: number): CellSpec => ({ v, s: xfs.indexOf(id) });

  it('数値のセル：m/d(aaa)・aaaa・日時と曜日・aaa だけ', async () => {
    const sheet = await sheetOf([[x(164, 45607), x(165, 45607), x(166, 45607.75), x(167, 45608)]], {
      xfs,
      numFmts,
    });
    expect(firstRow(sheet)).toEqual(['11/11(月)', '月曜日', '2024年11月11日(月) 18:00', '火']);
  });

  it('1904 年基準でも同じ日なら同じ曜日', async () => {
    const sheet = await sheetOf([[x(164, 44145), x(165, 44145)]], {
      xfs,
      numFmts,
      date1904: true,
    });
    expect(firstRow(sheet)).toEqual(['11/11(月)', '月曜日']);
  });

  it('t="d"（ISO 日付）のセルも曜日の書式で出す', async () => {
    const sheet = await sheetOf(
      [
        [
          { t: 'd', v: '2024-11-11', s: xfs.indexOf(164) },
          { t: 'd', v: '2024-11-12', s: xfs.indexOf(167) },
        ],
      ],
      {
        xfs,
        numFmts,
      },
    );
    expect(firstRow(sheet)).toEqual(['11/11(月)', '火']);
  });
});
