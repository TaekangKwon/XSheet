import { describe, expect, it } from 'vitest';
import { DEFAULT_LIMITS, openXlsx, XlsxError, type XlsxErrorKind } from '../../src/core/index';
import { buildXlsx, buildZip, utf8 } from '../../src/testing/make-xlsx';

/** 開こうとして出る例外の種類（出なければ 'opened'） */
async function kindOf(
  data: Uint8Array | ArrayBuffer,
  options?: Parameters<typeof openXlsx>[1],
): Promise<XlsxErrorKind | 'opened' | 'other'> {
  try {
    await openXlsx(data, options);
    return 'opened';
  } catch (e) {
    return e instanceof XlsxError ? e.kind : 'other';
  }
}

const GOOD = () => buildXlsx({ sheets: [{ name: 'S', rows: [['a']] }] });

describe('開けないファイルは決まった種類の例外（XlsxError）', () => {
  it('zip でないファイルは not-zip（空・短い・文字だけ・PDF）', async () => {
    expect(await kindOf(utf8('this is not an xlsx file at all'))).toBe('not-zip');
    expect(await kindOf(new Uint8Array(0))).toBe('not-zip');
    expect(await kindOf(new Uint8Array([0x50, 0x4b]))).toBe('not-zip');
    expect(await kindOf(utf8('%PDF-1.7\n...'))).toBe('not-zip');
  });

  it('先頭が PK だけでも、PK\\x03\\x04 でなければ not-zip', async () => {
    expect(await kindOf(utf8('PK is not enough'))).toBe('not-zip');
  });

  it('暗号化（OLE の先頭 D0 CF 11 E0）は encrypted', async () => {
    const ole = new Uint8Array(512);
    ole.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    expect(await kindOf(ole)).toBe('encrypted');
    expect(await kindOf(ole.buffer)).toBe('encrypted');
  });

  it('壊れた zip（途中で切れている）は broken', async () => {
    const good = await GOOD();
    expect(await kindOf(good.slice(0, Math.floor(good.length / 2)))).toBe('broken');
    // 先頭の印だけ合っていて、あとは壊れている
    expect(
      await kindOf(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10])),
    ).toBe('broken');
  });

  it('部品が欠けた zip（workbook が無い）は broken、Strict は unsupported', async () => {
    expect(await kindOf(await buildZip({ 'a.txt': 'x' }))).toBe('broken');
    expect(await kindOf(await buildXlsx({ strict: true, sheets: [{ name: 'S', rows: [] }] }))).toBe(
      'unsupported',
    );
  });

  it('例外のメッセージに原本の中身（シート名・セルの値）を入れない', async () => {
    const bytes = await buildXlsx({
      strict: true,
      sheets: [{ name: 'SECRET-SHEET', rows: [['SECRET-VALUE']] }],
    });
    const err = (await openXlsx(bytes).catch((e: unknown) => e)) as Error;
    expect(err).toBeInstanceOf(XlsxError);
    expect(err.message).not.toMatch(/SECRET/);
  });

  it('ArrayBuffer でも Uint8Array でも開ける', async () => {
    const good = await GOOD();
    expect(await kindOf(good)).toBe('opened');
    expect(
      await kindOf(
        good.buffer.slice(good.byteOffset, good.byteOffset + good.byteLength) as ArrayBuffer,
      ),
    ).toBe('opened');
  });
});

describe('展開の大きさの上限', () => {
  it('既定の上限は、200 万セル級（1 セル 64 バイトの見積り）の XML を読める大きさ', () => {
    expect(DEFAULT_LIMITS.maxSheetXmlBytes).toBeGreaterThanOrEqual(2_000_000 * 64);
    expect(DEFAULT_LIMITS.maxSheetXmlBytes).toBeLessThanOrEqual(512 * 1024 * 1024);
  });

  const bigSheet = () =>
    buildXlsx({
      sheets: [{ name: 'S', sheetDataXml: '<row r="1"><c r="A1"><v>1</v></c></row>'.repeat(2000) }],
    });

  it('シートの XML が上限を超えたら、シートを開くときに too-large', async () => {
    const book = await openXlsx(await bigSheet(), { limits: { maxSheetXmlBytes: 10_000 } });
    await expect(book.openSheet(1)).rejects.toMatchObject({ kind: 'too-large' });
  });

  it('上限以内なら読める', async () => {
    const book = await openXlsx(await bigSheet(), { limits: { maxSheetXmlBytes: 10_000_000 } });
    const sheet = await book.openSheet(1);
    expect(sheet.rowCount).toBe(1);
  });

  it('zip の見出しの大きさが偽りでも、展開しながら数えて止める', async () => {
    const bytes = await buildXlsx({
      sheets: [{ name: 'S', sheetDataXml: '<row r="1"><c r="A1"><v>1</v></c></row>'.repeat(2000) }],
      extraFiles: {},
    });
    // セントラルディレクトリ（PK\x01\x02）の「展開後の大きさ」（+24）を、実際より小さい 100 に書き換える
    const forged = bytes.slice();
    const view = new DataView(forged.buffer);
    // （workbook.xml など小さな部品は触らない。シートの XML の見出しだけを偽る）
    for (let i = 0; i + 46 < forged.length; i++) {
      if (
        forged[i] === 0x50 &&
        forged[i + 1] === 0x4b &&
        forged[i + 2] === 0x01 &&
        forged[i + 3] === 0x02
      ) {
        const nameLen = view.getUint16(i + 28, true);
        const name = new TextDecoder().decode(forged.subarray(i + 46, i + 46 + nameLen));
        if (name.endsWith('sheet1.xml')) view.setUint32(i + 24, 100, true);
      }
    }
    const book = await openXlsx(forged, { limits: { maxSheetXmlBytes: 2_000 } });
    await expect(book.openSheet(1)).rejects.toMatchObject({ kind: 'too-large' });
  });

  it('共有文字列が上限を超えたら too-large', async () => {
    const bytes = await buildXlsx({
      sharedStrings: Array.from({ length: 500 }, (_v, i) => `<t>string number ${i}</t>`),
      sheets: [{ name: 'S', rows: [[{ t: 's', v: 0 }]] }],
    });
    const book = await openXlsx(bytes, { limits: { maxSharedStringsBytes: 1_000 } });
    await expect(book.openSheet(1)).rejects.toMatchObject({ kind: 'too-large' });
  });

  it('小さな部品（workbook・styles）が上限を超えたら、開くとき too-large', async () => {
    const bytes = await buildXlsx({ sheets: [{ name: 'S', rows: [] }] });
    expect(await kindOf(bytes, { limits: { maxSmallPartBytes: 50 } })).toBe('too-large');
  });

  it('zip の中のファイル数・シート数・結合セル数の上限', async () => {
    const bytes = await buildXlsx({
      sheets: [
        { name: 'A', rows: [['x', 'y', 'z']], merges: ['A1:B1'] },
        { name: 'B', rows: [] },
      ],
    });
    expect(await kindOf(bytes, { limits: { maxEntries: 2 } })).toBe('too-large');
    expect(await kindOf(bytes, { limits: { maxSheets: 1 } })).toBe('too-large');
    const book = await openXlsx(bytes, { limits: { maxMerges: 0 } });
    await expect(book.openSheet(1)).rejects.toMatchObject({ kind: 'too-large' });
    expect(await kindOf(bytes)).toBe('opened');
  });
});
