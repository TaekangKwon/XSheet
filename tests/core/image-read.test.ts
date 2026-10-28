import { describe, expect, it } from 'vitest';
import { DEFAULT_LIMITS, openXlsx, type XlsxLimits } from '../../src/core/index';
import {
  makeBmp,
  makeEmf,
  makeGif,
  makeJpeg,
  makePng,
  makeSvg,
  makeTiff,
  makeWebp,
  makeWmf,
} from '../../src/testing/make-image';
import {
  buildXlsx,
  picXml,
  twoCellAnchor,
  wsDr,
  type SheetSpec,
} from '../../src/testing/make-xlsx';
import { openPackage } from '../../src/core/open';

type Files = Record<string, string | Uint8Array>;

/** 画像を media に置いた 1 シートのブックを開く（読むのは readImage だけ） */
async function bookWith(files: Files, limits?: Partial<XlsxLimits>) {
  const bytes = await buildXlsx({ sheets: [{ name: 'S', rows: [['a']] }], extraFiles: files });
  return openXlsx(bytes, { limits });
}

/** 画像 n 枚（`xl/media/i<番号>.png`）を 1 つの drawing で参照するシート */
function sheetWithPics(name: string, paths: string[]): SheetSpec {
  return {
    name,
    rows: [['a']],
    drawing: {
      xml: wsDr(
        paths.map((_p, i) =>
          twoCellAnchor(
            { col: 0, row: i },
            { col: 1, row: i + 1 },
            picXml({ name: `p${i}`, embed: `rId${i + 1}` }),
          ),
        ),
      ),
      rels: paths.map((p, i) => ({ id: `rId${i + 1}`, target: `../media/${p}` })),
    },
  };
}

/** zip の中央ディレクトリの「展開後の大きさ」を偽る（zip 爆弾の見出しの偽装） */
function lieAboutSize(zip: Uint8Array, name: string, declared: number): Uint8Array {
  const out = zip.slice();
  const dv = new DataView(out.buffer);
  const nameBytes = new TextEncoder().encode(name);
  for (let i = 0; i + 46 < out.length; i++) {
    if (dv.getUint32(i, true) !== 0x02014b50) continue;
    const nameLen = dv.getUint16(i + 28, true);
    const got = out.subarray(i + 46, i + 46 + nameLen);
    if (got.length === nameBytes.length && got.every((v, k) => v === nameBytes[k])) {
      dv.setUint32(i + 24, declared, true);
      return out;
    }
  }
  throw new Error('entry not found');
}

describe('TD-85-1 XlsxPackage.readBytes：バイト列で読み、展開しながら数えて上限で止まる', () => {
  it('部品をそのままのバイト列で返す（大きいものは複数の塊でも 1 本につなぐ）。無ければ null', async () => {
    const big = new Uint8Array(150_000).map((_v, i) => (i * 7) & 0xff);
    const zip = await buildXlsx({
      sheets: [{ name: 'S' }],
      extraFiles: { 'xl/media/a.bin': big, 'xl/media/e.bin': new Uint8Array(0) },
    });
    const pkg = await openPackage(zip, DEFAULT_LIMITS);
    expect(await pkg.readBytes('xl/media/a.bin', 1_000_000)).toEqual(big);
    expect(await pkg.readBytes('xl/media/e.bin', 10)).toEqual(new Uint8Array(0));
    expect(await pkg.readBytes('xl/media/none.bin', 10)).toBeNull();
  });

  it('上限ちょうどは読め、1 バイトでも超えたら too-large（見出しの大きさが正直な場合）', async () => {
    const zip = await buildXlsx({
      sheets: [{ name: 'S' }],
      extraFiles: { 'xl/media/a.bin': new Uint8Array(5000) },
    });
    const pkg = await openPackage(zip, DEFAULT_LIMITS);
    expect((await pkg.readBytes('xl/media/a.bin', 5000))!.length).toBe(5000);
    await expect(pkg.readBytes('xl/media/a.bin', 4999)).rejects.toMatchObject({
      kind: 'too-large',
    });
  });

  it('見出しの大きさを偽っても、展開しながら数えて止まる（zip 爆弾）', async () => {
    const zip = lieAboutSize(
      await buildXlsx({
        sheets: [{ name: 'S' }],
        extraFiles: { 'xl/media/bomb.bin': new Uint8Array(200_000) },
      }),
      'xl/media/bomb.bin',
      10,
    );
    const pkg = await openPackage(zip, DEFAULT_LIMITS);
    await expect(pkg.readBytes('xl/media/bomb.bin', 1000)).rejects.toMatchObject({
      kind: 'too-large',
    });
  });

  it('閉じた後は broken', async () => {
    const pkg = await openPackage(await buildXlsx({ sheets: [{ name: 'S' }] }), DEFAULT_LIMITS);
    pkg.dispose();
    await expect(pkg.readBytes('xl/workbook.xml', 100)).rejects.toMatchObject({ kind: 'broken' });
  });
});

describe('TD-85-10 readImage：先頭バイトで形式を判定し、縦横 px を返す', () => {
  it('PNG・JPEG・GIF・BMP・WebP の形式・MIME・縦横・バイト列', async () => {
    const book = await bookWith({
      'xl/media/a.png': makePng(40, 30),
      'xl/media/b.jpg': makeJpeg(64, 48),
      'xl/media/c.gif': makeGif(10, 20),
      'xl/media/d.bmp': makeBmp(30, 20),
      'xl/media/e.webp': makeWebp(200, 100),
    });
    const r = await book.readImage(1, 'xl/media/a.png');
    expect(r).toEqual({
      ok: true,
      format: 'png',
      mime: 'image/png',
      width: 40,
      height: 30,
      bytes: makePng(40, 30),
    });
    expect(await book.readImage(1, 'xl/media/b.jpg')).toMatchObject({
      format: 'jpeg',
      mime: 'image/jpeg',
      width: 64,
      height: 48,
    });
    expect(await book.readImage(1, 'xl/media/c.gif')).toMatchObject({
      format: 'gif',
      mime: 'image/gif',
      width: 10,
      height: 20,
    });
    expect(await book.readImage(1, 'xl/media/d.bmp')).toMatchObject({
      format: 'bmp',
      mime: 'image/bmp',
      width: 30,
      height: 20,
    });
    expect(await book.readImage(1, 'xl/media/e.webp')).toMatchObject({
      format: 'webp',
      mime: 'image/webp',
      width: 200,
      height: 100,
    });
  });

  it('拡張子を信じない：拡張子が嘘でも先頭バイトで判定する（.emf の PNG は PNG、.png の SVG・EMF は描けない）', async () => {
    const book = await bookWith({
      'xl/media/real.emf': makePng(5, 6),
      'xl/media/fake.png': makeSvg(),
      'xl/media/fake2.png': makeEmf(),
      'xl/media/noext': makeJpeg(7, 8),
    });
    expect(await book.readImage(1, 'xl/media/real.emf')).toMatchObject({
      ok: true,
      format: 'png',
      width: 5,
      height: 6,
    });
    expect(await book.readImage(1, 'xl/media/fake.png')).toEqual({
      ok: false,
      reason: 'unsupported',
      label: null,
    });
    expect(await book.readImage(1, 'xl/media/fake2.png')).toEqual({
      ok: false,
      reason: 'unsupported',
      label: 'EMF',
    });
    expect(await book.readImage(1, 'xl/media/noext')).toMatchObject({
      ok: true,
      format: 'jpeg',
      width: 7,
      height: 8,
    });
  });

  it('[Content_Types].xml の MIME を信じない：PNG と宣言されたファイルの中身が WMF なら描けない', async () => {
    const bytes = await buildXlsx({
      sheets: [{ name: 'S' }],
      extraFiles: {
        '[Content_Types].xml':
          '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>',
        'xl/media/a.png': makeWmf(),
      },
    });
    const book = await openXlsx(bytes);
    expect(await book.readImage(1, 'xl/media/a.png')).toEqual({
      ok: false,
      reason: 'unsupported',
      label: 'WMF',
    });
  });

  it('EMF・WMF・TIFF は形式名つきで「描けない」、SVG・空・乱数・途中で切れた PNG は名前なしで「描けない」。例外にしない', async () => {
    const book = await bookWith({
      'xl/media/a.emf': makeEmf(),
      'xl/media/b.wmf': makeWmf(),
      'xl/media/c.tif': makeTiff(),
      'xl/media/d.svg': makeSvg(),
      'xl/media/e.png': new Uint8Array(0),
      'xl/media/f.png': new Uint8Array(500).map((_v, i) => (i * 31 + 7) & 0xff),
      'xl/media/g.png': makePng(40, 30).slice(0, 20),
    });
    const fail = (label: string | null) => ({ ok: false, reason: 'unsupported', label });
    expect(await book.readImage(1, 'xl/media/a.emf')).toEqual(fail('EMF'));
    expect(await book.readImage(1, 'xl/media/b.wmf')).toEqual(fail('WMF'));
    expect(await book.readImage(1, 'xl/media/c.tif')).toEqual(fail('TIFF'));
    expect(await book.readImage(1, 'xl/media/d.svg')).toEqual(fail(null));
    expect(await book.readImage(1, 'xl/media/e.png')).toEqual(fail(null));
    expect(await book.readImage(1, 'xl/media/f.png')).toEqual(fail(null));
    expect(await book.readImage(1, 'xl/media/g.png')).toEqual(fail(null));
  });

  it('zip に無いパスは missing（例外にしない）。閉じた後は XlsxError(broken)', async () => {
    const book = await bookWith({});
    expect(await book.readImage(1, 'xl/media/none.png')).toEqual({
      ok: false,
      reason: 'missing',
      label: null,
    });
    book.close();
    await expect(book.readImage(1, 'xl/media/none.png')).rejects.toMatchObject({ kind: 'broken' });
  });

  it('画像のバイト・パスを console に出さない（成功・失敗とも）', async () => {
    const calls: unknown[][] = [];
    const orig = { log: console.log, warn: console.warn, error: console.error, info: console.info };
    console.log =
      console.warn =
      console.error =
      console.info =
        (...a: unknown[]) => {
          calls.push(a);
        };
    try {
      const book = await bookWith({
        'xl/media/secret.png': makePng(2, 2),
        'xl/media/bad.png': makeSvg(),
      });
      await book.readImage(1, 'xl/media/secret.png');
      await book.readImage(1, 'xl/media/bad.png');
      await book.readImage(1, 'xl/media/none.png');
    } finally {
      Object.assign(console, orig);
    }
    expect(calls).toEqual([]);
  });
});

describe('TD-85-11 readImage：上限は XlsxError にせず、理由つきで「出せない」を返す', () => {
  it('既定の上限の値（drawing 16MiB・図 1,000・画像 20MiB・合計 128MiB・1 辺 16,384px・64MP）', () => {
    expect(DEFAULT_LIMITS).toMatchObject({
      maxSmallPartBytes: 16 * 1024 * 1024,
      maxDrawings: 1000,
      maxImageBytes: 20 * 1024 * 1024,
      maxSheetImageBytes: 128 * 1024 * 1024,
      maxImageSidePx: 16384,
      maxImagePixels: 64 * 1024 * 1024,
    });
  });

  it('画像 1 枚の上限（maxImageBytes）：ちょうどは読め、超えたら too-large', async () => {
    const png = makePng(1, 1);
    const book = await bookWith({ 'xl/media/a.png': png }, { maxImageBytes: png.length });
    expect(await book.readImage(1, 'xl/media/a.png')).toMatchObject({ ok: true });
    const small = await bookWith({ 'xl/media/a.png': png }, { maxImageBytes: png.length - 1 });
    expect(await small.readImage(1, 'xl/media/a.png')).toEqual({
      ok: false,
      reason: 'too-large',
      label: null,
    });
  });

  it('既定の 20MiB を超える画像（21MiB の zip 爆弾）は too-large。表は開ける', async () => {
    const zip = await buildXlsx({
      sheets: [{ name: 'S', rows: [['cell']] }],
      extraFiles: { 'xl/media/big.png': new Uint8Array(21 * 1024 * 1024) },
    });
    const book = await openXlsx(zip);
    expect(await book.readImage(1, 'xl/media/big.png')).toEqual({
      ok: false,
      reason: 'too-large',
      label: null,
    });
    expect((await book.openSheet(1)).getRows(1, 1)[0]!.cells[0]).toMatchObject({ text: 'cell' });
  }, 30_000);

  it('見出しの大きさを偽っても、実際の大きさで上限を判定する', async () => {
    const zip = lieAboutSize(
      await buildXlsx({
        sheets: [{ name: 'S' }],
        extraFiles: { 'xl/media/bomb.png': new Uint8Array(200_000) },
      }),
      'xl/media/bomb.png',
      10,
    );
    const book = await openXlsx(zip, { limits: { maxImageBytes: 1000 } });
    expect(await book.readImage(1, 'xl/media/bomb.png')).toEqual({
      ok: false,
      reason: 'too-large',
      label: null,
    });
  });

  it('1 シートの合計（maxSheetImageBytes）：超えた分から over-budget。呼んだ順に数え、並列に呼んでも順序が保たれる', async () => {
    const png = makePng(1, 1); // 45 バイト
    const bytes = await buildXlsx({
      sheets: [sheetWithPics('S', ['a.png', 'b.png', 'c.png'])],
      extraFiles: { 'xl/media/a.png': png, 'xl/media/b.png': png, 'xl/media/c.png': png },
    });
    const book = await openXlsx(bytes, { limits: { maxSheetImageBytes: png.length * 2 } });
    const { drawings } = await book.openDrawings(1);
    const results = await Promise.all(drawings.map((d) => book.readImage(1, d.mediaPath!)));
    expect(results.map((r) => (r.ok ? 'ok' : r.reason))).toEqual(['ok', 'ok', 'over-budget']);
  });

  it('同じ画像を何度読んでも合計に 1 回だけ数える（複数のアンカーで共有する画像）', async () => {
    const png = makePng(1, 1);
    const bytes = await buildXlsx({
      sheets: [sheetWithPics('S', ['a.png'])],
      extraFiles: { 'xl/media/a.png': png },
    });
    const book = await openXlsx(bytes, { limits: { maxSheetImageBytes: png.length } });
    await book.openDrawings(1);
    for (let i = 0; i < 3; i++)
      expect(await book.readImage(1, 'xl/media/a.png')).toMatchObject({ ok: true });
  });

  it('合計はシートごとに数える（1 枚目のシートで使い切っても、2 枚目のシートの画像は読める）', async () => {
    const png = makePng(1, 1);
    const bytes = await buildXlsx({
      sheets: [sheetWithPics('S1', ['a.png', 'b.png']), sheetWithPics('S2', ['c.png'])],
      extraFiles: { 'xl/media/a.png': png, 'xl/media/b.png': png, 'xl/media/c.png': png },
    });
    const book = await openXlsx(bytes, { limits: { maxSheetImageBytes: png.length * 2 } });
    const s1 = await book.openDrawings(1);
    const s2 = await book.openDrawings(2);
    expect(await book.readImage(1, s1.drawings[0]!.mediaPath!)).toMatchObject({ ok: true });
    expect(await book.readImage(1, s1.drawings[1]!.mediaPath!)).toMatchObject({ ok: true });
    expect(await book.readImage(2, s2.drawings[0]!.mediaPath!)).toMatchObject({ ok: true });
  });

  it('1 辺の上限（16,384px）と画素数の上限（64MP）：超えたら oversize-px、ちょうどは読める', async () => {
    const book = await bookWith({
      'xl/media/side-ok.png': makePng(16384, 1),
      'xl/media/side-ng.png': makePng(16385, 1),
      'xl/media/h-ng.png': makePng(1, 16385),
      'xl/media/mp-ok.png': makePng(8192, 8192),
      'xl/media/mp-ng.png': makePng(8193, 8192),
      'xl/media/mp-ng2.png': makePng(16384, 16384),
    });
    const kind = async (p: string) => {
      const r = await book.readImage(1, p);
      return r.ok ? 'ok' : r.reason;
    };
    expect(await kind('xl/media/side-ok.png')).toBe('ok');
    expect(await kind('xl/media/side-ng.png')).toBe('oversize-px');
    expect(await kind('xl/media/h-ng.png')).toBe('oversize-px');
    expect(await kind('xl/media/mp-ok.png')).toBe('ok');
    expect(await kind('xl/media/mp-ng.png')).toBe('oversize-px');
    expect(await kind('xl/media/mp-ng2.png')).toBe('oversize-px');
  });

  it('画像の上限・合計の超過があっても、表（セルの値）の読み取りは成功する', async () => {
    const png = makePng(1, 1);
    const bytes = await buildXlsx({
      sheets: [sheetWithPics('S', ['a.png', 'b.png'])],
      extraFiles: { 'xl/media/a.png': png, 'xl/media/b.png': png },
    });
    const book = await openXlsx(bytes, { limits: { maxSheetImageBytes: 10, maxImageBytes: 10 } });
    const { drawings } = await book.openDrawings(1);
    for (const d of drawings)
      expect(await book.readImage(1, d.mediaPath!)).toMatchObject({ ok: false });
    expect((await book.openSheet(1)).getRows(1, 1)[0]!.cells[0]).toMatchObject({ text: 'a' });
  });

  it('壊れた圧縮データの画像は broken（例外にしない）', async () => {
    const zip = await buildXlsx({
      sheets: [{ name: 'S' }],
      extraFiles: { 'xl/media/a.png': new Uint8Array(50_000).map((_v, i) => (i * 13) & 0xff) },
    });
    const out = zip.slice();
    const dv = new DataView(out.buffer);
    const nameBytes = new TextEncoder().encode('xl/media/a.png');
    for (let i = 0; i + 30 < out.length; i++) {
      if (dv.getUint32(i, true) !== 0x04034b50) continue;
      const nameLen = dv.getUint16(i + 26, true);
      const extraLen = dv.getUint16(i + 28, true);
      const got = out.subarray(i + 30, i + 30 + nameLen);
      if (got.length === nameBytes.length && got.every((v, k) => v === nameBytes[k])) {
        out.fill(0xff, i + 30 + nameLen + extraLen, i + 30 + nameLen + extraLen + 8);
        break;
      }
    }
    const book = await openXlsx(out);
    expect(await book.readImage(1, 'xl/media/a.png')).toEqual({
      ok: false,
      reason: 'broken',
      label: null,
    });
  });

  it('GIF：論理画面 1×1・フレーム 65535×65535 は oversize-px（フレームの大きさで上限を判定する）', async () => {
    const book = await bookWith({
      'xl/media/a.gif': makeGif(1, 1, { frame: { w: 65535, h: 65535 } }),
      'xl/media/b.gif': makeGif(1, 1, { frame: { w: 640, h: 480 } }),
    });
    expect(await book.readImage(1, 'xl/media/a.gif')).toEqual({
      ok: false,
      reason: 'oversize-px',
      label: null,
    });
    expect(await book.readImage(1, 'xl/media/b.gif')).toMatchObject({
      ok: true,
      width: 640,
      height: 480,
    });
  });
});

describe('TD-85-15 readImage(index, path)：合計はシート番号で数え、理由を取り違えない', () => {
  it('シート 1・2 が同じ画像を参照しても、openDrawings(1)→(2)→readImage(1, 画像) は、シート 1 の合計に数える', async () => {
    const a = makePng(4, 3);
    const b = makePng(5, 5);
    const bytes = await buildXlsx({
      sheets: [sheetWithPics('S1', ['a.png', 'b.png']), sheetWithPics('S2', ['a.png'])],
      extraFiles: { 'xl/media/a.png': a, 'xl/media/b.png': b },
    });
    // シートの合計は a 1 枚分。シート 1 の a を数えたなら、b は入らない
    const book = await openXlsx(bytes, { limits: { maxSheetImageBytes: a.length } });
    const s1 = await book.openDrawings(1);
    const s2 = await book.openDrawings(2);
    expect(s1.drawings.map((d) => d.mediaPath)).toEqual(['xl/media/a.png', 'xl/media/b.png']);
    expect(s2.drawings.map((d) => d.mediaPath)).toEqual(['xl/media/a.png']);
    expect(await book.readImage(1, 'xl/media/a.png')).toMatchObject({ ok: true });
    expect(await book.readImage(1, 'xl/media/b.png')).toMatchObject({
      ok: false,
      reason: 'over-budget',
    });
    // シート 2 は自分の合計で数える（a は入る）
    expect(await book.readImage(2, 'xl/media/a.png')).toMatchObject({ ok: true });
  });

  it('シートの番号が範囲外なら RangeError', async () => {
    const book = await bookWith({});
    await expect(book.readImage(0, 'xl/media/a.png')).rejects.toBeInstanceOf(RangeError);
    await expect(book.readImage(2, 'xl/media/a.png')).rejects.toBeInstanceOf(RangeError);
  });

  it('シートの残りが 1 枚の上限より少なくても、1 枚の上限を超える画像は too-large（over-budget ではない）', async () => {
    const png = makePng(1, 1);
    const big = new Uint8Array(3000);
    big.set(png);
    const book = await bookWith(
      { 'xl/media/a.png': png, 'xl/media/big.png': big },
      { maxImageBytes: 2000, maxSheetImageBytes: 2000 },
    );
    expect(await book.readImage(1, 'xl/media/a.png')).toMatchObject({ ok: true });
    // 残りは 1955 バイト（1 枚の上限 2000 より少ない）。big は 1 枚の上限を超える
    expect(await book.readImage(1, 'xl/media/big.png')).toEqual({
      ok: false,
      reason: 'too-large',
      label: null,
    });
  });

  it('使い切り（over-budget）の後は、残りに収まる大きさの画像でも over-budget。too-large は使い切りの印にならない。数え済みの画像は読める', async () => {
    const png = makePng(1, 1); // 45 バイト
    const mid = new Uint8Array(80); // 80 バイト（PNG の先頭つき）
    mid.set(png);
    const huge = new Uint8Array(500); // 1 枚の上限（200）超え
    huge.set(png);
    const book = await bookWith(
      {
        'xl/media/a.png': png,
        'xl/media/huge.png': huge,
        'xl/media/c.png': makePng(2, 2),
        'xl/media/mid.png': mid,
        'xl/media/tiny.gif': makeGif(1, 1), // 14 バイト
      },
      { maxImageBytes: 200, maxSheetImageBytes: 100 },
    );
    const kind = async (p: string) => {
      const r = await book.readImage(1, p);
      return r.ok ? 'ok' : r.reason;
    };
    expect(await kind('xl/media/a.png')).toBe('ok'); // 45
    expect(await kind('xl/media/huge.png')).toBe('too-large'); // 使い切りの印にならない
    expect(await kind('xl/media/c.png')).toBe('ok'); // 45 ＋ 45 ＝ 90
    expect(await kind('xl/media/mid.png')).toBe('over-budget'); // 90 ＋ 80 ＞ 100 → 使い切り
    // 90 ＋ 14 ＝ 104 ＞ 100 だが、使い切りの理由は残りの大小ではなく印による
    expect(await kind('xl/media/tiny.gif')).toBe('over-budget');
    expect(await kind('xl/media/a.png')).toBe('ok');
  });

  it('使い切りの後は、残りに収まる画像（45 ＋ 14 ＜ 100）でも over-budget', async () => {
    const png = makePng(1, 1); // 45 バイト
    const mid = new Uint8Array(80);
    mid.set(png);
    const book = await bookWith(
      {
        'xl/media/a.png': png,
        'xl/media/mid.png': mid,
        'xl/media/tiny.gif': makeGif(1, 1), // 14 バイト
      },
      { maxSheetImageBytes: 100 },
    );
    expect(await book.readImage(1, 'xl/media/a.png')).toMatchObject({ ok: true });
    expect(await book.readImage(1, 'xl/media/mid.png')).toMatchObject({ reason: 'over-budget' });
    expect(await book.readImage(1, 'xl/media/tiny.gif')).toMatchObject({
      ok: false,
      reason: 'over-budget',
    });
  });
});
