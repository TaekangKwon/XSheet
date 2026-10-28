import { describe, expect, it } from 'vitest';
import { sniffImage } from '../../src/core/image';
import {
  makeBmp,
  makeBmpCore,
  makeEmf,
  makeGif,
  makeJpeg,
  makePng,
  makeSvg,
  makeTiff,
  makeWebp,
  makeWmf,
} from '../../src/testing/make-image';

describe('TD-85-2 sniffImage：先頭バイトで形式を判定し、ヘッダの縦横を読む', () => {
  it('PNG：IHDR の幅・高さ（ビッグエンディアン）', () => {
    expect(sniffImage(makePng(40, 30))).toEqual({ ok: true, format: 'png', width: 40, height: 30 });
    expect(sniffImage(makePng(16384, 1))).toMatchObject({ ok: true, width: 16384, height: 1 });
  });

  it('JPEG：SOF0 の高さ・幅。progressive（SOF2）も読み、DHT（C4）・DAC（CC）・JPG（C8）は SOF と取り違えない', () => {
    expect(sniffImage(makeJpeg(64, 48))).toEqual({
      ok: true,
      format: 'jpeg',
      width: 64,
      height: 48,
    });
    expect(sniffImage(makeJpeg(640, 480, 0xc2))).toMatchObject({ width: 640, height: 480 });
    // DHT（0xC4）だけで SOF が無い画像は縦横が読めない＝描けない
    expect(sniffImage(makeJpeg(64, 48, 0xc4))).toEqual({ ok: false, label: null });
    expect(sniffImage(makeJpeg(64, 48, 0xcc))).toEqual({ ok: false, label: null });
    expect(sniffImage(makeJpeg(64, 48, 0xc8))).toEqual({ ok: false, label: null });
  });

  it('JPEG：SOF の前に長いセグメント（コメント・EXIF）があっても辿る', () => {
    const head = makeJpeg(100, 50);
    // SOI の直後へ COM セグメント（長さ 1000）を差し込む
    const com = new Uint8Array(2 + 1000);
    com.set([0xff, 0xfe, 1000 >> 8, 1000 & 0xff], 0);
    const bytes = new Uint8Array(head.length + com.length);
    bytes.set(head.subarray(0, 2), 0);
    bytes.set(com, 2);
    bytes.set(head.subarray(2), 2 + com.length);
    expect(sniffImage(bytes)).toMatchObject({ ok: true, format: 'jpeg', width: 100, height: 50 });
  });

  it('GIF：論理画面の幅・高さ（リトルエンディアン）', () => {
    expect(sniffImage(makeGif(320, 200))).toEqual({
      ok: true,
      format: 'gif',
      width: 320,
      height: 200,
    });
  });

  it('BMP：BITMAPINFOHEADER の幅・高さ（負の高さ＝上から下の並びも絶対値）と旧形式 BITMAPCOREHEADER', () => {
    expect(sniffImage(makeBmp(30, 20))).toEqual({ ok: true, format: 'bmp', width: 30, height: 20 });
    expect(sniffImage(makeBmp(30, -20))).toMatchObject({ width: 30, height: 20 });
    expect(sniffImage(makeBmpCore(31, 21))).toMatchObject({ format: 'bmp', width: 31, height: 21 });
  });

  it('WebP：lossy（VP8）・lossless（VP8L）・extended（VP8X）', () => {
    expect(sniffImage(makeWebp(200, 100, 'lossy'))).toEqual({
      ok: true,
      format: 'webp',
      width: 200,
      height: 100,
    });
    expect(sniffImage(makeWebp(201, 101, 'lossless'))).toMatchObject({ width: 201, height: 101 });
    expect(sniffImage(makeWebp(16383, 5000, 'extended'))).toMatchObject({
      width: 16383,
      height: 5000,
    });
  });

  it('EMF・WMF・TIFF は描けない形式として名前を返す', () => {
    expect(sniffImage(makeEmf())).toEqual({ ok: false, label: 'EMF' });
    expect(sniffImage(makeWmf())).toEqual({ ok: false, label: 'WMF' });
    expect(sniffImage(makeTiff())).toEqual({ ok: false, label: 'TIFF' });
  });

  it('SVG・空・短い・乱数・形式の印だけで本体が無いバイト列は、例外を出さず「描けない」', () => {
    expect(sniffImage(makeSvg())).toEqual({ ok: false, label: null });
    expect(sniffImage(new Uint8Array(0))).toEqual({ ok: false, label: null });
    expect(sniffImage(new Uint8Array([0x89, 0x50]))).toEqual({ ok: false, label: null });
    expect(sniffImage(new Uint8Array(100).fill(0xff))).toEqual({ ok: false, label: null });
    expect(sniffImage(new Uint8Array(100))).toEqual({ ok: false, label: null });
    expect(sniffImage(new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]))).toEqual({
      ok: false,
      label: null,
    });
  });

  it('途中で切れたバイト列は、どの長さで切っても例外を出さない（形式ごとに全長を試す）', () => {
    const samples = [
      makePng(40, 30),
      makeJpeg(64, 48),
      makeGif(10, 10),
      makeGif(1, 1, { frame: { w: 300, h: 200 }, gct: true, ext: true }),
      makeBmp(10, 10),
      makeBmpCore(10, 10),
      makeWebp(10, 10, 'lossy'),
      makeWebp(10, 10, 'lossless'),
      makeWebp(10, 10, 'extended'),
      makeEmf(),
      makeWmf(),
      makeTiff(),
    ];
    for (const full of samples) {
      for (let n = 0; n <= full.length; n++) {
        expect(() => sniffImage(full.subarray(0, n))).not.toThrow();
      }
    }
  });

  it('縦横が 0 の画像は描けない（0 除算・空の枠を避ける）', () => {
    expect(sniffImage(makePng(0, 30))).toEqual({ ok: false, label: null });
    expect(sniffImage(makeGif(10, 0))).toEqual({ ok: false, label: null });
    expect(sniffImage(makeJpeg(0, 0))).toEqual({ ok: false, label: null });
  });

  it('不正な値（JPEG のセグメント長 0・1）で止まらず、無限ループしない', () => {
    // SOI＋長さ 0 の APP0 を延々と並べる
    const bytes = new Uint8Array(2000);
    bytes.set([0xff, 0xd8]);
    for (let i = 2; i < bytes.length - 3; i += 4) bytes.set([0xff, 0xe0, 0x00, 0x00], i);
    expect(sniffImage(bytes)).toEqual({ ok: false, label: null });
  });
});

describe('TD-85-14 sniffImage：GIF は最初のフレーム（Image Descriptor）の大きさも読み、論理画面と大きい方を返す', () => {
  it('論理画面 1×1 でフレームが大きい GIF は、フレームの大きさ（大きい方）を返す', () => {
    expect(sniffImage(makeGif(1, 1, { frame: { w: 65535, h: 65535 } }))).toEqual({
      ok: true,
      format: 'gif',
      width: 65535,
      height: 65535,
    });
    // 幅と高さは別々に大きい方
    expect(sniffImage(makeGif(500, 10, { frame: { w: 20, h: 300 } }))).toMatchObject({
      width: 500,
      height: 300,
    });
  });

  it('グローバルカラーテーブルと拡張ブロック（Graphic Control Extension）を読み飛ばして、フレームに届く', () => {
    expect(
      sniffImage(makeGif(1, 1, { frame: { w: 640, h: 480 }, gct: true, ext: true })),
    ).toMatchObject({ width: 640, height: 480 });
    expect(sniffImage(makeGif(1, 1, { frame: { w: 640, h: 480 }, ext: true }))).toMatchObject({
      width: 640,
      height: 480,
    });
  });

  it('フレームが論理画面より小さいときは論理画面のまま。フレーム無し・トレーラーだけの GIF も従来どおり', () => {
    expect(sniffImage(makeGif(100, 80, { frame: { w: 10, h: 10 } }))).toMatchObject({
      width: 100,
      height: 80,
    });
    expect(sniffImage(makeGif(320, 200))).toMatchObject({ width: 320, height: 200 });
  });

  it('切れた・不正な GIF（拡張ブロックのサブブロックが範囲外・ブロック種別が不正）で例外を出さず、無限ループしない', () => {
    const full = makeGif(1, 1, { frame: { w: 640, h: 480 }, gct: true, ext: true });
    for (let n = 0; n <= full.length; n++)
      expect(() => sniffImage(full.subarray(0, n))).not.toThrow();
    // サブブロックの長さが残りより大きい
    const bad = makeGif(7, 9, { ext: true });
    bad[bad.length - 4] = 255;
    expect(sniffImage(bad)).toMatchObject({ ok: true, width: 7, height: 9 });
    // 長さ 0 のサブブロックだけが続く拡張ブロック
    const loop = new Uint8Array(5000);
    loop.set(makeGif(3, 4).subarray(0, 13));
    loop[13] = 0x21;
    loop[14] = 0xf9;
    expect(sniffImage(loop)).toMatchObject({ ok: true, width: 3, height: 4 });
    // 不明なブロック種別
    const unknown = makeGif(3, 4, { frame: { w: 999, h: 999 } });
    unknown[13] = 0x55;
    expect(sniffImage(unknown)).toMatchObject({ width: 3, height: 4 });
  });
});
