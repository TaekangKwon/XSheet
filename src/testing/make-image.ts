/**
 * 試験用の画像のバイト列を組む部品（ヘッダだけの最小の画像。ブラウザで描ける完全な画像ではない）。
 * `readImage` が先頭バイトから形式と縦横を読むことの試験で、実在の画像を使わずに値を作り分けるために使う。
 */

/** 数値を n バイトのリトルエンディアンにする */
const le = (v: number, n: number): number[] =>
  Array.from({ length: n }, (_v, i) => Math.floor(v / 2 ** (8 * i)) & 0xff);

/** 数値を n バイトのビッグエンディアンにする */
const be = (v: number, n: number): number[] => le(v, n).reverse();

/** 文字を ASCII のバイトにする */
const ascii = (s: string): number[] => Array.from(s, (c) => c.charCodeAt(0));

/** PNG（シグネチャ＋IHDR＋IEND。CRC は 0 のまま） */
export function makePng(w: number, h: number): Uint8Array {
  return new Uint8Array([
    0x89,
    0x50,
    0x4e,
    0x47,
    0x0d,
    0x0a,
    0x1a,
    0x0a,
    ...be(13, 4),
    ...ascii('IHDR'),
    ...be(w, 4),
    ...be(h, 4),
    8,
    2,
    0,
    0,
    0,
    ...be(0, 4),
    ...be(0, 4),
    ...ascii('IEND'),
    ...be(0, 4),
  ]);
}

/** JPEG（SOI＋APP0＋SOF0＋EOI）。`sof` でフレーム見出しの種類（既定 0xC0＝baseline。0xC2＝progressive）を替える */
export function makeJpeg(w: number, h: number, sof = 0xc0): Uint8Array {
  return new Uint8Array([
    0xff,
    0xd8,
    0xff,
    0xe0,
    ...be(16, 2),
    ...ascii('JFIF'),
    0,
    1,
    1,
    0,
    ...be(96, 2),
    ...be(96, 2),
    0,
    0,
    0xff,
    sof,
    ...be(17, 2),
    8,
    ...be(h, 2),
    ...be(w, 2),
    3,
    1,
    0x22,
    0,
    2,
    0x11,
    1,
    3,
    0x11,
    1,
    0xff,
    0xd9,
  ]);
}

/** makeGif の追加の指定 */
export interface GifOptions {
  /** 最初の Image Descriptor の幅・高さ（論理画面と別の値にできる） */
  frame?: { w: number; h: number };
  /** グローバルカラーテーブル（2 色）を付ける */
  gct?: boolean;
  /** Image Descriptor の前に Graphic Control Extension（サブブロック 1 つ）を付ける */
  ext?: boolean;
}

/** GIF89a（論理画面の幅・高さ。`frame` を渡すと最初のフレームの Image Descriptor も付く） */
export function makeGif(w: number, h: number, o: GifOptions = {}): Uint8Array {
  return new Uint8Array([
    ...ascii('GIF89a'),
    ...le(w, 2),
    ...le(h, 2),
    o.gct ? 0x80 : 0,
    0,
    0,
    ...(o.gct ? [0, 0, 0, 255, 255, 255] : []),
    // 拡張ブロック：導入子 0x21・ラベル 0xF9・サブブロック（長さ 4）・終端 0
    ...(o.ext ? [0x21, 0xf9, 4, 0, 0, 0, 0, 0] : []),
    // Image Descriptor：0x2C・左 2・上 2・幅 2・高さ 2・フラグ 1
    ...(o.frame ? [0x2c, 0, 0, 0, 0, ...le(o.frame.w, 2), ...le(o.frame.h, 2), 0] : []),
    0x3b,
  ]);
}

/** BMP（BITMAPINFOHEADER・40 バイト）。高さが負なら上から下の並び */
export function makeBmp(w: number, h: number): Uint8Array {
  return new Uint8Array([
    ...ascii('BM'),
    ...le(54, 4),
    ...le(0, 4),
    ...le(54, 4),
    ...le(40, 4),
    ...le(w >>> 0, 4),
    ...le(h < 0 ? 2 ** 32 + h : h, 4),
    ...le(1, 2),
    ...le(24, 2),
    ...le(0, 24),
  ]);
}

/** BMP（旧形式 BITMAPCOREHEADER・12 バイト。幅・高さは 16 ビット） */
export function makeBmpCore(w: number, h: number): Uint8Array {
  return new Uint8Array([
    ...ascii('BM'),
    ...le(26, 4),
    ...le(0, 4),
    ...le(26, 4),
    ...le(12, 4),
    ...le(w, 2),
    ...le(h, 2),
    ...le(1, 2),
    ...le(24, 2),
  ]);
}

/** WebP の種類（lossy＝`VP8 `・lossless＝`VP8L`・extended＝`VP8X`） */
export type WebpKind = 'lossy' | 'lossless' | 'extended';

/** WebP（RIFF＋最初のチャンクの見出し） */
export function makeWebp(w: number, h: number, kind: WebpKind = 'lossy'): Uint8Array {
  let chunk: number[];
  if (kind === 'lossy') {
    // フレームタグ 3 バイト＋開始コード 9D 01 2A＋幅・高さ（14 ビット。上位 2 ビットは拡大率）
    chunk = [
      ...ascii('VP8 '),
      ...le(10, 4),
      0x10,
      0,
      0,
      0x9d,
      0x01,
      0x2a,
      ...le(w, 2),
      ...le(h, 2),
    ];
  } else if (kind === 'lossless') {
    // シグネチャ 0x2F＋14 ビットの（幅-1）＋14 ビットの（高さ-1）
    const bits = (w - 1) | ((h - 1) << 14);
    chunk = [...ascii('VP8L'), ...le(5, 4), 0x2f, ...le(bits, 4)];
  } else {
    // フラグ 4 バイト＋24 ビットの（幅-1）・（高さ-1）
    chunk = [...ascii('VP8X'), ...le(10, 4), 0, 0, 0, 0, ...le(w - 1, 3), ...le(h - 1, 3)];
  }
  return new Uint8Array([...ascii('RIFF'), ...le(4 + chunk.length, 4), ...ascii('WEBP'), ...chunk]);
}

/** EMF（種類 1 ＋ 40 バイト目の ` EMF`） */
export function makeEmf(): Uint8Array {
  const b = new Uint8Array(88);
  b.set(le(1, 4), 0);
  b.set(le(88, 4), 4);
  b.set(ascii(' EMF'), 40);
  return b;
}

/** WMF（Aldus の配置情報つき） */
export function makeWmf(): Uint8Array {
  const b = new Uint8Array(40);
  b.set([0xd7, 0xcd, 0xc6, 0x9a], 0);
  return b;
}

/** TIFF（リトルエンディアン） */
export function makeTiff(): Uint8Array {
  return new Uint8Array([0x49, 0x49, 0x2a, 0x00, 8, 0, 0, 0, 0, 0]);
}

/** SVG（文字。画像としては扱わない） */
export function makeSvg(): Uint8Array {
  return new TextEncoder().encode(
    '<svg xmlns="http://www.w3.org/2000/svg" width="50" height="50"><circle r="20"/></svg>',
  );
}
