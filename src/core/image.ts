// 正: InterHub docs/plan/adr/0026-xlsx-image-overlay.md
/**
 * 画像の読み取り：先頭バイトで形式（PNG・JPEG・GIF・BMP・WebP）を判定し、ヘッダから縦横の px を読む。
 * 拡張子と `[Content_Types].xml` の MIME は信じない（画像は細工できる非信頼のバイナリ）。
 * 切れたバイト列・不正な値で例外を出さない（読む前に必ず長さを確かめる）。SVG は先頭バイトで判定できないので扱わない。
 * 画像のバイト・パスはログに出さない（console を使わない）。
 */
import type { XlsxLimits } from './limits';
import type { XlsxPackage } from './open';
import { XlsxError, type ImageFormat, type ImageResult } from './types';

/** `sniffImage` の結果。描けない形式は、分かれば形式名（EMF など）を label に入れる */
export type ImageSniff =
  | { ok: true; format: ImageFormat; width: number; height: number }
  | { ok: false; label: string | null };

/** 描ける形式 → MIME */
const MIME: Readonly<Record<ImageFormat, string>> = Object.freeze({
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
});

/** 判定できなかった（描けない）結果 */
const UNKNOWN: ImageSniff = Object.freeze({ ok: false, label: null });

/** 先頭から ascii の文字列が並んでいるか（範囲外は false） */
function startsWith(b: Uint8Array, at: number, s: string): boolean {
  if (at + s.length > b.length) return false;
  for (let i = 0; i < s.length; i++) if (b[at + i] !== s.charCodeAt(i)) return false;
  return true;
}

/** 縦横が 1 以上の整数なら結果にする（0・負・非整数は不正＝描けない） */
function dims(format: ImageFormat, width: number, height: number): ImageSniff {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1)
    return UNKNOWN;
  return { ok: true, format, width, height };
}

/** JPEG のフレーム見出し（SOF0〜SOF15。DHT=C4・JPG=C8・DAC=CC は別物） */
const isSof = (m: number): boolean =>
  m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc;

/** JPEG：マーカーのセグメントを長さで飛ばして SOF を探す（SOS・EOI まで。ループごとに必ず進む） */
function sniffJpeg(b: Uint8Array): ImageSniff {
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  let i = 2;
  while (i + 4 <= b.length) {
    if (b[i] !== 0xff) return UNKNOWN;
    let m = b[i + 1]!;
    // 0xFF の詰め物を読み飛ばす
    while (m === 0xff && i + 2 < b.length) {
      i++;
      m = b[i + 1]!;
    }
    // 長さを持たないマーカー（SOI・TEM・RSTn）
    if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) {
      i += 2;
      continue;
    }
    // 縦横を読む前に画像データ（SOS）か終わり（EOI）に来た
    if (m === 0xd9 || m === 0xda) return UNKNOWN;
    if (m === 0xff || i + 4 > b.length) return UNKNOWN;
    const segLen = dv.getUint16(i + 2);
    if (segLen < 2) return UNKNOWN;
    if (isSof(m)) {
      // 長さ 2・精度 1・高さ 2・幅 2
      if (i + 9 > b.length) return UNKNOWN;
      return dims('jpeg', dv.getUint16(i + 7), dv.getUint16(i + 5));
    }
    i += 2 + segLen;
  }
  return UNKNOWN;
}

/** BMP の DIB 見出しの大きさ（旧形式 12・INFO 40・V2 52・V3 56・OS/2 64・V4 108・V5 124） */
const BMP_DIB_SIZES: ReadonlySet<number> = new Set([12, 40, 52, 56, 64, 108, 124]);

/**
 * GIF：論理画面の幅・高さ（リトルエンディアン）と、最初の Image Descriptor（0x2C）の幅・高さを読み、それぞれ大きい方を返す。
 * 論理画面を小さく偽ってフレームだけ巨大にした GIF を、上限判定で見逃さないため。
 * 拡張ブロック（0x21）はサブブロックの長さで読み飛ばす。どの読み出しも長さを確かめ、ループごとに必ず進む。
 */
function sniffGif(b: Uint8Array): ImageSniff {
  if (b.length < 10) return UNKNOWN;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  let w = dv.getUint16(6, true);
  let h = dv.getUint16(8, true);
  // 論理画面記述子（7 バイト）が揃っているときだけブロックを辿る
  if (b.length >= 13) {
    let i = 13;
    // グローバルカラーテーブル（フラグの最上位ビット。大きさは 3 × 2^(n+1) バイト）
    if ((b[10]! & 0x80) !== 0) i += 3 * 2 ** ((b[10]! & 7) + 1);
    while (i < b.length) {
      const block = b[i]!;
      if (block === 0x21) {
        // 拡張ブロック：導入子・ラベルのあと、長さ 0 のサブブロックまで
        i += 2;
        while (i < b.length) {
          const n = b[i]!;
          i += 1 + n;
          if (n === 0) break;
        }
      } else {
        // Image Descriptor：左 2・上 2・幅 2・高さ 2。トレーラー（0x3B）・不正なブロックならここで終わり
        if (block === 0x2c && i + 9 <= b.length) {
          w = Math.max(w, dv.getUint16(i + 5, true));
          h = Math.max(h, dv.getUint16(i + 7, true));
        }
        break;
      }
    }
  }
  return dims('gif', w, h);
}

/** WebP：RIFF の最初のチャンク（VP8・VP8L・VP8X）の見出しから読む */
function sniffWebp(b: Uint8Array): ImageSniff {
  if (!startsWith(b, 0, 'RIFF') || !startsWith(b, 8, 'WEBP')) return UNKNOWN;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  if (startsWith(b, 12, 'VP8 ')) {
    // フレームタグ 3 バイト・開始コード 9D 01 2A・幅と高さ（14 ビット。上位 2 ビットは拡大率）
    if (b.length < 30 || b[23] !== 0x9d || b[24] !== 0x01 || b[25] !== 0x2a) return UNKNOWN;
    return dims('webp', dv.getUint16(26, true) & 0x3fff, dv.getUint16(28, true) & 0x3fff);
  }
  if (startsWith(b, 12, 'VP8L')) {
    // シグネチャ 0x2F・14 ビットの（幅-1）・14 ビットの（高さ-1）
    if (b.length < 25 || b[20] !== 0x2f) return UNKNOWN;
    const bits = dv.getUint32(21, true);
    return dims('webp', (bits & 0x3fff) + 1, ((bits >>> 14) & 0x3fff) + 1);
  }
  if (startsWith(b, 12, 'VP8X')) {
    // フラグ 4 バイト・24 ビットの（幅-1）・（高さ-1）
    if (b.length < 30) return UNKNOWN;
    const u24 = (at: number): number => b[at]! | (b[at + 1]! << 8) | (b[at + 2]! << 16);
    return dims('webp', u24(24) + 1, u24(27) + 1);
  }
  return UNKNOWN;
}

/**
 * 先頭バイトで画像の形式を判定し、ヘッダから縦横（px）を読む。
 * 描けない形式（EMF・WMF・TIFF）は形式名つきで、判定できないもの（SVG・不正・切れた・空）は名前なしで `ok: false` を返す。
 * どんなバイト列でも例外を出さない。
 */
export function sniffImage(b: Uint8Array): ImageSniff {
  // PNG：シグネチャ 8 バイトのあと IHDR（長さ 4・種類 4・幅 4・高さ 4）
  if (b.length >= 8 && b[0] === 0x89 && startsWith(b, 1, 'PNG') && b[4] === 0x0d) {
    if (b.length < 24 || !startsWith(b, 12, 'IHDR')) return UNKNOWN;
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    const w = dv.getUint32(16);
    const h = dv.getUint32(20);
    // PNG の仕様上の上限は 2^31-1
    return w > 0x7fffffff || h > 0x7fffffff ? UNKNOWN : dims('png', w, h);
  }
  if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return sniffJpeg(b);
  if (startsWith(b, 0, 'GIF87a') || startsWith(b, 0, 'GIF89a')) return sniffGif(b);
  // BMP：DIB 見出しの大きさで、旧形式（16 ビットの幅・高さ）と 32 ビットの形式を分ける
  if (startsWith(b, 0, 'BM') && b.length >= 22) {
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    const dibSize = dv.getUint32(14, true);
    if (!BMP_DIB_SIZES.has(dibSize)) return UNKNOWN;
    if (dibSize === 12) return dims('bmp', dv.getUint16(18, true), dv.getUint16(20, true));
    if (b.length < 26) return UNKNOWN;
    // 高さが負なら上から下の並び（絶対値が高さ）
    return dims('bmp', dv.getInt32(18, true), Math.abs(dv.getInt32(22, true)));
  }
  if (startsWith(b, 0, 'RIFF')) return sniffWebp(b);
  // ここから先は描けない形式（形式名だけ分かれば枠の文字になる）
  if (
    b.length >= 44 &&
    b[0] === 1 &&
    b[1] === 0 &&
    b[2] === 0 &&
    b[3] === 0 &&
    startsWith(b, 40, ' EMF')
  )
    return { ok: false, label: 'EMF' };
  if (
    (b.length >= 4 && b[0] === 0xd7 && b[1] === 0xcd && b[2] === 0xc6 && b[3] === 0x9a) ||
    (b.length >= 4 && (b[0] === 1 || b[0] === 2) && b[1] === 0 && b[2] === 9 && b[3] === 0)
  )
    return { ok: false, label: 'WMF' };
  if (
    b.length >= 4 &&
    ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a && b[3] === 0) ||
      (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0 && b[3] === 0x2a))
  )
    return { ok: false, label: 'TIFF' };
  return UNKNOWN;
}

/** 1 シートの画像の合計の数え方（readImage の呼び出しをまたいで持つ） */
export interface ImageBudget {
  /** これまでに読んだ画像の展開後の合計（バイト） */
  used: number;
  /** 合計の上限を一度でも超えた（以後、新しい画像はすべて over-budget。ADR-0026 §決定 7 の「以後の画像」） */
  exhausted: boolean;
  /** すでに数えた画像のパス（同じ画像を何度読んでも 1 回だけ数える） */
  counted: Set<string>;
}

/** 空の ImageBudget */
export function createImageBudget(): ImageBudget {
  return { used: 0, exhausted: false, counted: new Set() };
}

/**
 * zip の中の画像を読み、形式と縦横を返す。出せないときも例外にせず、理由つきの `ok: false` を返す。
 * - 1 枚の上限（maxImageBytes）で展開しながら止める（超えたら too-large。これを先に判定するので、シートの残りが少なくても理由は取り違えない）。
 * - 1 枚に収まる画像でも、シートの合計の残りを超えるなら over-budget。一度 over-budget になったら、以後そのシートの新しい画像はすべて over-budget（読まずに断る）。
 *   すでに数えた画像（同じ画像の再読み込み）は、合計を使い切った後も読める。
 * - 形式を判定したあと、1 辺と画素数の上限を確かめる。
 */
export async function readImageFrom(
  pkg: XlsxPackage,
  path: string,
  limits: XlsxLimits,
  budget: ImageBudget,
): Promise<ImageResult> {
  if (!pkg.has(path)) return { ok: false, reason: 'missing', label: null };
  const charged = !budget.counted.has(path);
  if (charged && budget.exhausted) return { ok: false, reason: 'over-budget', label: null };
  let bytes: Uint8Array | null;
  try {
    bytes = await pkg.readBytes(path, limits.maxImageBytes);
  } catch (e) {
    if (e instanceof XlsxError && e.kind === 'too-large')
      return { ok: false, reason: 'too-large', label: null };
    return { ok: false, reason: 'broken', label: null };
  }
  if (bytes === null) return { ok: false, reason: 'missing', label: null };
  if (charged) {
    if (budget.used + bytes.length > limits.maxSheetImageBytes) {
      budget.exhausted = true;
      return { ok: false, reason: 'over-budget', label: null };
    }
    budget.used += bytes.length;
    budget.counted.add(path);
  }
  const sniff = sniffImage(bytes);
  if (!sniff.ok) return { ok: false, reason: 'unsupported', label: sniff.label };
  const { format, width, height } = sniff;
  if (
    width > limits.maxImageSidePx ||
    height > limits.maxImageSidePx ||
    width * height > limits.maxImagePixels
  )
    return { ok: false, reason: 'oversize-px', label: null };
  return { ok: true, format, mime: MIME[format], width, height, bytes };
}
