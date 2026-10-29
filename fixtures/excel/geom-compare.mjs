// ISSUE-93：Excel の図形の画像と XSheet の描画の画像を、輪郭の距離で比べる純ロジック（依存は Node 標準の zlib だけ・DOM 非依存）。
// 使い手は compare-ref-geoms.mjs（Playwright で XSheet の描画を写して比べる）と、tests/core/geom-compare.test.ts（合成画像での単体テスト）。
// 比べ方（ADR-0028 §決定 10 単位 d・U3）：
//   1. 画像から「インク」（白い背景との差が大きい画素）のマスクを取る。色によらないよう、画像の最大の濃さの半分を閾値にする（薄い色の枠も取れる）。
//   2. 両方のマスクをインクの外接矩形の左上でそろえる（Excel の画像は枠線の縁の扱いで 1px ほど位置がずれうるため）。
//   3. 輪郭（インクの画素のうち、4 近傍に非インクを持つもの）どうしの距離を、双方向に測る。95 パーセンタイルを 100% 換算の px で返す。
// scale = 画像の 1 画素あたりの 100%（96dpi）の px の逆数＝画像の画素 ÷ 100% の px。Excel（画面 125%＝120dpi）と deviceScaleFactor 1.25 の写しは 1.25。
import zlib from 'node:zlib';
import { Buffer } from 'node:buffer';

/** 輪郭の距離（95 パーセンタイル）の許容（100% 換算の px。ISSUE-93・ADR-0028 U3） */
export const TOLERANCE_PX = 1.5;
/** Excel（画面 125%）と XSheet の写し（deviceScaleFactor 1.25）の、100% の 1px あたりの画素数 */
export const SCALE_125 = 1.25;

// ---------- PNG ----------

const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10];
let crcTable = null;
/** CRC-32（PNG のチャンク用） */
function crc32(buf) {
  if (crcTable === null) {
    crcTable = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c;
    }
  }
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = crcTable[(c ^ buf[i]) & 255] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

/** PNG のチャンク 1 つ（長さ・種類・本体・CRC） */
function chunk(type, data) {
  const out = Buffer.alloc(12 + data.length);
  out.writeUInt32BE(data.length, 0);
  out.write(type, 4, 'ascii');
  Buffer.from(data).copy(out, 8);
  out.writeUInt32BE(crc32(out.subarray(4, 8 + data.length)), 8 + data.length);
  return out;
}

/**
 * フィルタ済みの走査線（各行の先頭にフィルタ番号 1 バイト）から PNG を組み立てる（テスト・検査用。画素の並びは呼び出し側が決める）。
 * @param {{ width: number, height: number, colorType: number, bitDepth: number, raw: Uint8Array, plte?: Uint8Array, trns?: Uint8Array }} o
 */
export function encodePngRaw(o) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(o.width, 0);
  ihdr.writeUInt32BE(o.height, 4);
  ihdr[8] = o.bitDepth;
  ihdr[9] = o.colorType;
  const parts = [Buffer.from(PNG_SIGNATURE), chunk('IHDR', ihdr)];
  if (o.plte) parts.push(chunk('PLTE', o.plte));
  if (o.trns) parts.push(chunk('tRNS', o.trns));
  parts.push(chunk('IDAT', zlib.deflateSync(Buffer.from(o.raw))), chunk('IEND', Buffer.alloc(0)));
  return Buffer.concat(parts);
}

/** RGBA（8bit・width×height×4）を PNG にする（フィルタなし）。デバッグ用の重ね合わせ画像の書き出しに使う */
export function encodePng(width, height, rgba) {
  const raw = Buffer.alloc(height * (width * 4 + 1));
  for (let y = 0; y < height; y++) {
    raw[y * (width * 4 + 1)] = 0;
    Buffer.from(rgba.buffer, rgba.byteOffset + y * width * 4, width * 4).copy(
      raw,
      y * (width * 4 + 1) + 1,
    );
  }
  return encodePngRaw({ width, height, colorType: 6, bitDepth: 8, raw });
}

/**
 * PNG（パレット・グレー・RGB・RGBA・グレー＋α。1〜16bit。インターレースなし）を RGBA（8bit）に読む。
 * Excel の Chart.Export はパレット形式、Chrome のスクリーンショットは RGB(A) の PNG を書く。
 * @param {Uint8Array} buf
 * @returns {{ width: number, height: number, data: Uint8Array }} data は width×height×4（R, G, B, A）
 */
export function decodePng(buf) {
  const b = Buffer.from(buf);
  for (let i = 0; i < 8; i++) if (b[i] !== PNG_SIGNATURE[i]) throw new Error('PNG ではない');
  let p = 8;
  let width = 0;
  let height = 0;
  let bd = 0;
  let ct = 0;
  let interlace = 0;
  let plte = null;
  let trns = null;
  const idat = [];
  while (p + 8 <= b.length) {
    const len = b.readUInt32BE(p);
    const type = b.toString('ascii', p + 4, p + 8);
    const d = b.subarray(p + 8, p + 8 + len);
    if (type === 'IHDR') {
      width = d.readUInt32BE(0);
      height = d.readUInt32BE(4);
      bd = d[8];
      ct = d[9];
      interlace = d[12];
    } else if (type === 'PLTE') plte = d;
    else if (type === 'tRNS') trns = d;
    else if (type === 'IDAT') idat.push(d);
    else if (type === 'IEND') break;
    p += 12 + len;
  }
  if (interlace !== 0) throw new Error('インターレースの PNG は読めない');
  const channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[ct];
  if (channels === undefined) throw new Error(`PNG の色の種類が不明: ${ct}`);
  const bitsPP = channels * bd;
  const bpp = Math.max(1, bitsPP >> 3);
  const stride = Math.ceil((width * bitsPP) / 8);
  const raw = zlib.inflateSync(Buffer.concat(idat));
  const out = Buffer.alloc(height * stride);
  for (let y = 0; y < height; y++) {
    const f = raw[y * (stride + 1)];
    const line = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1));
    for (let x = 0; x < stride; x++) {
      const a = x >= bpp ? out[y * stride + x - bpp] : 0;
      const up = y > 0 ? out[(y - 1) * stride + x] : 0;
      const c = x >= bpp && y > 0 ? out[(y - 1) * stride + x - bpp] : 0;
      let v = line[x];
      if (f === 1) v += a;
      else if (f === 2) v += up;
      else if (f === 3) v += (a + up) >> 1;
      else if (f === 4) {
        const pa = Math.abs(up - c);
        const pb = Math.abs(a - c);
        const pc = Math.abs(a + up - 2 * c);
        v += pa <= pb && pa <= pc ? a : pb <= pc ? up : c;
      }
      out[y * stride + x] = v & 255;
    }
  }
  const data = new Uint8Array(width * height * 4);
  const maxv = (1 << bd) - 1;
  /** x 番目の標本（チャンネル ch）の値を 0〜255 にする（16bit は上位 1 バイト） */
  const sample = (y, x, ch) => {
    if (bd === 8) return out[y * stride + x * channels + ch];
    if (bd === 16) return out[y * stride + (x * channels + ch) * 2];
    const bit = (x * channels + ch) * bd;
    const v = (out[y * stride + (bit >> 3)] >> (8 - bd - (bit & 7))) & maxv;
    return ct === 3 ? v : Math.round((v * 255) / maxv);
  };
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const o = (y * width + x) * 4;
      let r;
      let g;
      let bl;
      let al = 255;
      if (ct === 3) {
        const idx = sample(y, x, 0);
        if (plte === null) throw new Error('PLTE が無い');
        r = plte[idx * 3];
        g = plte[idx * 3 + 1];
        bl = plte[idx * 3 + 2];
        if (trns !== null && idx < trns.length) al = trns[idx];
      } else if (ct === 0 || ct === 4) {
        r = g = bl = sample(y, x, 0);
        if (ct === 4) al = sample(y, x, 1);
      } else {
        r = sample(y, x, 0);
        g = sample(y, x, 1);
        bl = sample(y, x, 2);
        if (ct === 6) al = sample(y, x, 3);
      }
      data[o] = r;
      data[o + 1] = g;
      data[o + 2] = bl;
      data[o + 3] = al;
    }
  }
  return { width, height, data };
}

// ---------- マスク ----------

/**
 * Excel の画像（マゼンタの背景に、白い背景の矩形つきで貼られた図）の、背景でない部分の外接矩形（x1・y1 を含む）。無ければ null。
 * @param {{ width: number, height: number, data: Uint8Array }} img
 * @param {number[]} [bg] 背景の色（既定：マゼンタ）
 * @param {number} [tol] 背景との許容差（各成分）
 */
export function pictureRect(img, bg = [255, 0, 255], tol = 6) {
  let x0 = -1;
  let y0 = -1;
  let x1 = -1;
  let y1 = -1;
  for (let y = 0; y < img.height; y++) {
    for (let x = 0; x < img.width; x++) {
      const o = (y * img.width + x) * 4;
      const isBg =
        Math.abs(img.data[o] - bg[0]) <= tol &&
        Math.abs(img.data[o + 1] - bg[1]) <= tol &&
        Math.abs(img.data[o + 2] - bg[2]) <= tol;
      if (isBg) continue;
      if (x0 < 0 || x < x0) x0 = x;
      if (x1 < 0 || x > x1) x1 = x;
      if (y0 < 0 || y < y0) y0 = y;
      if (y1 < 0 || y > y1) y1 = y;
    }
  }
  return x0 < 0 ? null : { x0, y0, x1, y1 };
}

/**
 * 画像（の矩形）から「インク」のマスクを取る。濃さ＝255 − 最小の成分（白に重ねた見た目。α は白との混色）。
 * 閾値＝最大の濃さ × fraction（ただし floor 以上）。最大の濃さが floor に満たなければ何も無い（空のマスク）。
 * @param {{ width: number, height: number, data: Uint8Array }} img
 * @param {{ x0: number, y0: number, x1: number, y1: number } | null} [rect] 取り出す矩形（省略：全体）
 * @param {{ fraction?: number, floor?: number }} [opts]
 * @returns {{ w: number, h: number, data: Uint8Array }}
 */
export function inkMask(img, rect = null, opts = {}) {
  const fraction = opts.fraction ?? 0.5;
  const floor = opts.floor ?? 32;
  const r = rect ?? { x0: 0, y0: 0, x1: img.width - 1, y1: img.height - 1 };
  const w = r.x1 - r.x0 + 1;
  const h = r.y1 - r.y0 + 1;
  const dark = new Uint8Array(w * h);
  let max = 0;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const o = ((y + r.y0) * img.width + x + r.x0) * 4;
      const a = img.data[o + 3] / 255;
      const mn = Math.min(img.data[o], img.data[o + 1], img.data[o + 2]);
      const d = Math.round(255 - (mn * a + 255 * (1 - a)));
      dark[y * w + x] = d;
      if (d > max) max = d;
    }
  }
  const data = new Uint8Array(w * h);
  if (max >= floor) {
    const thr = Math.max(floor, Math.round(max * fraction));
    for (let i = 0; i < dark.length; i++) data[i] = dark[i] >= thr ? 1 : 0;
  }
  return { w, h, data };
}

/** マスクのインクの外接矩形（x1・y1 を含む）。インクが無ければ null */
export function maskBox(mask) {
  let x0 = -1;
  let y0 = -1;
  let x1 = -1;
  let y1 = -1;
  for (let y = 0; y < mask.h; y++) {
    for (let x = 0; x < mask.w; x++) {
      if (!mask.data[y * mask.w + x]) continue;
      if (x0 < 0 || x < x0) x0 = x;
      if (x1 < 0 || x > x1) x1 = x;
      if (y0 < 0 || y < y0) y0 = y;
      if (y1 < 0 || y > y1) y1 = y;
    }
  }
  return x0 < 0 ? null : { x0, y0, x1, y1 };
}

/** マスクを矩形で切り出す（矩形の左上が新しい原点） */
export function cropMask(mask, box) {
  const w = box.x1 - box.x0 + 1;
  const h = box.y1 - box.y0 + 1;
  const data = new Uint8Array(w * h);
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++) data[y * w + x] = mask.data[(y + box.y0) * mask.w + x + box.x0];
  return { w, h, data };
}

/** マスクをインクの外接矩形で切り出す（インクが無ければ空の 0×0） */
export function cropToInk(mask) {
  const box = maskBox(mask);
  return box === null ? { w: 0, h: 0, data: new Uint8Array(0) } : cropMask(mask, box);
}

/**
 * マスクを最近傍で w×h に拡縮する（XSheet は行の高さを整数 px に丸めるため、twoCellAnchor の図が Excel の xfrm より数 px 大きくなる。
 * 形の誤りと配置の丸めを切り分けるため、比較の前に寸法をそろえる版を使う。寸法の差は別に報告する）
 */
export function resizeMask(mask, w, h) {
  const data = new Uint8Array(w * h);
  if (mask.w === 0 || mask.h === 0) return { w, h, data };
  for (let y = 0; y < h; y++) {
    const sy = Math.min(mask.h - 1, Math.floor(((y + 0.5) * mask.h) / h));
    for (let x = 0; x < w; x++)
      data[y * w + x] =
        mask.data[sy * mask.w + Math.min(mask.w - 1, Math.floor(((x + 0.5) * mask.w) / w))];
  }
  return { w, h, data };
}

/** 輪郭の画素（インクのうち、4 近傍に非インク・外側を持つもの）を `[x0, y0, x1, y1, …]` で返す */
export function contourPoints(mask) {
  const pts = [];
  const at = (x, y) =>
    x < 0 || y < 0 || x >= mask.w || y >= mask.h ? 0 : mask.data[y * mask.w + x];
  for (let y = 0; y < mask.h; y++) {
    for (let x = 0; x < mask.w; x++) {
      if (!mask.data[y * mask.w + x]) continue;
      if (!at(x - 1, y) || !at(x + 1, y) || !at(x, y - 1) || !at(x, y + 1)) pts.push(x, y);
    }
  }
  return pts;
}

/** マスクを Chebyshev 距離 r だけ膨らませる（四角い構造要素。文字と枠線の分離に使う） */
export function dilate(mask, r) {
  const tmp = new Uint8Array(mask.w * mask.h);
  for (let y = 0; y < mask.h; y++) {
    for (let x = 0; x < mask.w; x++) {
      if (!mask.data[y * mask.w + x]) continue;
      for (let dx = -r; dx <= r; dx++) {
        const xx = x + dx;
        if (xx >= 0 && xx < mask.w) tmp[y * mask.w + xx] = 1;
      }
    }
  }
  const out = new Uint8Array(mask.w * mask.h);
  for (let y = 0; y < mask.h; y++) {
    for (let x = 0; x < mask.w; x++) {
      if (!tmp[y * mask.w + x]) continue;
      for (let dy = -r; dy <= r; dy++) {
        const yy = y + dy;
        if (yy >= 0 && yy < mask.h) out[yy * mask.w + x] = 1;
      }
    }
  }
  return { w: mask.w, h: mask.h, data: out };
}

// ---------- 距離 ----------

/** Felzenszwalb の 1 次元の距離変換（f は各点の二乗距離の初期値。大きな有限値＝点なし） */
function dt1d(f, n, d, v, z) {
  let k = 0;
  v[0] = 0;
  z[0] = -1e30;
  z[1] = 1e30;
  for (let q = 1; q < n; q++) {
    let s = (f[q] + q * q - (f[v[k]] + v[k] * v[k])) / (2 * q - 2 * v[k]);
    while (s <= z[k]) {
      k--;
      s = (f[q] + q * q - (f[v[k]] + v[k] * v[k])) / (2 * q - 2 * v[k]);
    }
    k++;
    v[k] = q;
    z[k] = s;
    z[k + 1] = 1e30;
  }
  k = 0;
  for (let q = 0; q < n; q++) {
    while (z[k + 1] < q) k++;
    d[q] = (q - v[k]) * (q - v[k]) + f[v[k]];
  }
}

/**
 * w×h の格子で、点（`[x0, y0, x1, y1, …]`）までのユークリッド距離を全画素について返す（行優先）。点が 1 つも無ければ全部 Infinity。
 */
export function distanceField(w, h, pts) {
  const out = new Float64Array(w * h);
  if (pts.length === 0) return out.fill(Infinity);
  const BIG = 1e20;
  const g = new Float64Array(w * h).fill(BIG);
  for (let i = 0; i < pts.length; i += 2) g[pts[i + 1] * w + pts[i]] = 0;
  const n = Math.max(w, h);
  const f = new Float64Array(n);
  const d = new Float64Array(n);
  const v = new Int32Array(n);
  const z = new Float64Array(n + 1);
  for (let x = 0; x < w; x++) {
    for (let y = 0; y < h; y++) f[y] = g[y * w + x];
    dt1d(f, h, d, v, z);
    for (let y = 0; y < h; y++) g[y * w + x] = d[y];
  }
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) f[x] = g[y * w + x];
    dt1d(f, w, d, v, z);
    for (let x = 0; x < w; x++) out[y * w + x] = Math.sqrt(d[x]);
  }
  return out;
}

/** 昇順の配列の p 分位（最近順位法。p は 0〜1） */
export function percentile(sorted, p) {
  if (sorted.length === 0) return 0;
  return sorted[Math.max(0, Math.min(sorted.length - 1, Math.ceil(p * sorted.length) - 1))];
}

/**
 * 2 つの点の集合の、双方向の最近点距離（画素）。a の各点から b への距離と、b の各点から a への距離をまとめる。
 * @returns {number[]} 昇順
 */
export function pointDistances(w, h, ptsA, ptsB) {
  const toB = distanceField(w, h, ptsB);
  const toA = distanceField(w, h, ptsA);
  const out = [];
  for (let i = 0; i < ptsA.length; i += 2) out.push(toB[ptsA[i + 1] * w + ptsA[i]]);
  for (let i = 0; i < ptsB.length; i += 2) out.push(toA[ptsB[i + 1] * w + ptsB[i]]);
  return out.sort((x, y) => x - y);
}

/**
 * 輪郭の比較（a＝Excel・b＝XSheet。どちらもインクの外接矩形で切り出して左上をそろえたマスク）。
 * 距離・寸法の差は 100% 換算の px（画素 ÷ scale）。`dw`・`dh` は b − a（XSheet の方が大きければ正）。
 * どちらかが空なら p95 を 999 にして `empty` に空の側（'excel'・'interhub'＝XSheet の側（名前は InterHub から写したまま）・'both'）を入れる。両方空なら p95 は 0（どちらも描かない）。
 */
export function compareContours(a, b, scale = SCALE_125) {
  const empty =
    a.w === 0 && b.w === 0 ? 'both' : a.w === 0 ? 'excel' : b.w === 0 ? 'interhub' : null;
  if (empty !== null) {
    return {
      p95: empty === 'both' ? 0 : 999,
      max: empty === 'both' ? 0 : 999,
      mean: 0,
      nA: a.w === 0 ? 0 : 1,
      nB: b.w === 0 ? 0 : 1,
      dw: 0,
      dh: 0,
      empty,
    };
  }
  const W = Math.max(a.w, b.w);
  const H = Math.max(a.h, b.h);
  const ptsA = contourPoints(a);
  const ptsB = contourPoints(b);
  const d = pointDistances(W, H, ptsA, ptsB).map((v) => v / scale);
  const mean = d.reduce((s, v) => s + v, 0) / d.length;
  return {
    p95: percentile(d, 0.95),
    max: d[d.length - 1],
    mean,
    nA: ptsA.length / 2,
    nB: ptsB.length / 2,
    dw: (b.w - a.w) / scale,
    dh: (b.h - a.h) / scale,
    empty: null,
  };
}

/**
 * Zhang-Suen の細線化（8 近傍・輪郭の線を 1 画素幅の中心線にする）。枠線だけの図形の比較用：
 * 枠線の太さの描き方の違い（Excel は細く、XSheet のアンチエイリアスは 2 画素ほど）が、外接矩形・輪郭の位置の差に化けるのを避ける。
 */
export function thin(mask) {
  const w = mask.w;
  const h = mask.h;
  const d = Uint8Array.from(mask.data);
  const at = (x, y) => (x < 0 || y < 0 || x >= w || y >= h ? 0 : d[y * w + x]);
  let changed = true;
  while (changed) {
    changed = false;
    for (let step = 0; step < 2; step++) {
      const del = [];
      for (let y = 0; y < h; y++) {
        for (let x = 0; x < w; x++) {
          if (!d[y * w + x]) continue;
          const p = [
            at(x, y - 1),
            at(x + 1, y - 1),
            at(x + 1, y),
            at(x + 1, y + 1),
            at(x, y + 1),
            at(x - 1, y + 1),
            at(x - 1, y),
            at(x - 1, y - 1),
          ];
          const b = p.reduce((s, v) => s + v, 0);
          if (b < 2 || b > 6) continue;
          let a = 0;
          for (let i = 0; i < 8; i++) if (p[i] === 0 && p[(i + 1) % 8] === 1) a++;
          if (a !== 1) continue;
          if (
            step === 0
              ? p[0] * p[2] * p[4] !== 0 || p[2] * p[4] * p[6] !== 0
              : p[0] * p[2] * p[6] !== 0 || p[0] * p[4] * p[6] !== 0
          )
            continue;
          del.push(y * w + x);
        }
      }
      for (const i of del) d[i] = 0;
      if (del.length > 0) changed = true;
    }
  }
  return { w, h, data: d };
}

/**
 * 中心線（細線化）どうしの比較。a＝Excel・b＝XSheet のインクのマスク（切り出し前でよい）。
 * それぞれ細線化して中心線の外接矩形で切り出す。\`normalize\` なら b の点を a の外接矩形の大きさへ線形に拡縮してから測る（寸法の差を除いた「形だけ」の比較）。
 * \`searchPx\`（画素）> 0 なら、b を ±searchPx の整数の平行移動で動かし、双方向の平均距離が最小になる位置で測る（位置合わせ）。
 * 鋭角の頂点（矢じり・星・太陽の光線）は、角の描き方（miter の張り出し）の差で外接矩形が数画素ずれ、左上そろえでは全体が動いて見えるため。
 * \`searchPx\` が 0 なら外接矩形の左上でそろえるだけ。戻り値は compareContours と同じ＋ \`shift\`（XSheet を動かした量。100% 換算の px）。
 * dw・dh は中心線の外接矩形の寸法の差（b − a。100% 換算の px。拡縮前の値）。
 */
export function compareSkeletons(a, b, scale = SCALE_125, normalize = true, searchPx = 0) {
  const sa = cropToInk(thin(a));
  const sb = cropToInk(thin(b));
  const empty =
    sa.w === 0 && sb.w === 0 ? 'both' : sa.w === 0 ? 'excel' : sb.w === 0 ? 'interhub' : null;
  if (empty !== null) return { ...compareContours(sa, sb, scale), shift: { x: 0, y: 0 } };
  const ptsA = [];
  for (let y = 0; y < sa.h; y++)
    for (let x = 0; x < sa.w; x++) if (sa.data[y * sa.w + x]) ptsA.push(x, y);
  const ptsB = [];
  const kx = normalize && sb.w > 1 ? (sa.w - 1) / (sb.w - 1) : 1;
  const ky = normalize && sb.h > 1 ? (sa.h - 1) / (sb.h - 1) : 1;
  for (let y = 0; y < sb.h; y++)
    for (let x = 0; x < sb.w; x++)
      if (sb.data[y * sb.w + x]) ptsB.push(Math.round(x * kx), Math.round(y * ky));
  const P = Math.max(0, Math.round(searchPx));
  const W = Math.max(sa.w, Math.round((sb.w - 1) * kx) + 1) + 2 * P;
  const H = Math.max(sa.h, Math.round((sb.h - 1) * ky) + 1) + 2 * P;
  const shifted = (pts) => pts.map((v) => v + P);
  const dtA = distanceField(W, H, shifted(ptsA));
  const dtB = distanceField(W, H, shifted(ptsB));
  /** B を (tx, ty) 動かしたときの双方向の距離（画素） */
  const measure = (tx, ty) => {
    const d = [];
    for (let i = 0; i < ptsA.length; i += 2) {
      const x = ptsA[i] + P - tx;
      const y = ptsA[i + 1] + P - ty;
      d.push(x >= 0 && y >= 0 && x < W && y < H ? dtB[y * W + x] : 1e6);
    }
    for (let i = 0; i < ptsB.length; i += 2) {
      const x = ptsB[i] + P + tx;
      const y = ptsB[i + 1] + P + ty;
      d.push(x >= 0 && y >= 0 && x < W && y < H ? dtA[y * W + x] : 1e6);
    }
    return d;
  };
  let best = null;
  for (let ty = -P; ty <= P; ty++) {
    for (let tx = -P; tx <= P; tx++) {
      const d = measure(tx, ty);
      const mean = d.reduce((s, v) => s + v, 0) / d.length;
      if (
        best === null ||
        mean < best.mean - 1e-12 ||
        (Math.abs(mean - best.mean) <= 1e-12 && Math.hypot(tx, ty) < Math.hypot(best.tx, best.ty))
      )
        best = { tx, ty, mean, d };
    }
  }
  const dist = best.d.map((v) => v / scale).sort((x, y) => x - y);
  return {
    p95: percentile(dist, 0.95),
    max: dist[dist.length - 1],
    mean: best.mean / scale,
    nA: ptsA.length / 2,
    nB: ptsB.length / 2,
    dw: (sb.w - sa.w) / scale,
    dh: (sb.h - sa.h) / scale,
    shift: { x: best.tx / scale, y: best.ty / scale },
    empty: null,
  };
}

/**
 * 重ね合わせ画像（RGBA）：a のみ＝赤・b のみ＝青・両方＝黒・どちらも無い＝白。左上をそろえる（デバッグ用）。
 * @returns {{ width: number, height: number, data: Uint8Array }}
 */
export function overlayImage(a, b) {
  const width = Math.max(a.w, b.w, 1);
  const height = Math.max(a.h, b.h, 1);
  const data = new Uint8Array(width * height * 4).fill(255);
  const has = (m, x, y) => x < m.w && y < m.h && m.data[y * m.w + x] === 1;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const ia = has(a, x, y);
      const ib = has(b, x, y);
      if (!ia && !ib) continue;
      const o = (y * width + x) * 4;
      const c = ia && ib ? [0, 0, 0] : ia ? [230, 40, 40] : [40, 40, 230];
      data[o] = c[0];
      data[o + 1] = c[1];
      data[o + 2] = c[2];
    }
  }
  return { width, height, data };
}

// ---------- U4：文字 ----------

/** a から b を引く（a かつ not b。左上をそろえ、b の範囲の外は引かない） */
function subtractMask(a, b) {
  const data = new Uint8Array(a.data.length);
  for (let y = 0; y < a.h; y++) {
    for (let x = 0; x < a.w; x++) {
      const inB = x < b.w && y < b.h && b.data[y * b.w + x] === 1;
      data[y * a.w + x] = a.data[y * a.w + x] && !inB ? 1 : 0;
    }
  }
  return { w: a.w, h: a.h, data };
}

/** 文字のインクの画素（枠線の外接矩形の左上を原点にした座標）と、枠線の外接矩形の大きさ。枠線が無ければ null */
function textInkPoints(text, outline, sepPx) {
  const box = maskBox(outline);
  if (box === null) return null;
  const ink = subtractMask(text, dilate(outline, sepPx));
  const pts = [];
  for (let y = 0; y < ink.h; y++)
    for (let x = 0; x < ink.w; x++) if (ink.data[y * ink.w + x]) pts.push(x - box.x0, y - box.y0);
  return { pts, w: box.x1 - box.x0 + 1, h: box.y1 - box.y0 + 1, ink, box };
}

/** 点の集合の重心と外接矩形（100% 換算の px。点が無ければ null） */
function pointStats(pts, scale) {
  if (pts.length === 0) return null;
  let sx = 0;
  let sy = 0;
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -Infinity;
  let y1 = -Infinity;
  for (let i = 0; i < pts.length; i += 2) {
    sx += pts[i];
    sy += pts[i + 1];
    x0 = Math.min(x0, pts[i]);
    x1 = Math.max(x1, pts[i]);
    y0 = Math.min(y0, pts[i + 1]);
    y1 = Math.max(y1, pts[i + 1]);
  }
  const n = pts.length / 2;
  return {
    cx: sx / n / scale,
    cy: sy / n / scale,
    x0: x0 / scale,
    y0: y0 / scale,
    x1: (x1 + 1) / scale,
    y1: (y1 + 1) / scale,
  };
}

/** 文字のインクの輪郭の点（text のマスクを枠線の左上基準の座標にしたもの）。dilate した枠線を除いたインクの輪郭 */
function textContour(info) {
  const pts = contourPoints(info.ink);
  const out = [];
  for (let i = 0; i < pts.length; i += 2) out.push(pts[i] - info.box.x0, pts[i + 1] - info.box.y0);
  return out;
}

/**
 * 反転した図形の文字の比較（U4）。枠線だけの双子（outline）と文字入り（text）の組を、Excel（e）と XSheet（i）で渡す
 * （どちらも各画像の座標のマスク。文字入りと双子は同じ範囲を切ったもの）。
 * 枠線の外接矩形の左上を原点にして、文字のインク（枠線を `sepPx` 画素だけ膨らませた範囲の外）の重心・外接矩形・輪郭の距離を比べる。
 * `p95Upright`＝文字のインクの輪郭を各自の重心にそろえて重ねたときの距離、`p95Rot180`＝XSheet の文字を重心まわりに 180° 回してから重ねたときの距離（向きの判定。位置は `dCentroid`）。
 * `orientation`＝小さい方（'upright'＝XSheet と同じ向き、'rot180'＝Excel は 180° 回している、'unknown'＝文字が取れない）。
 * @param {{ eText: object, eOutline: object, iText: object, iOutline: object, scale?: number, sepPx?: number }} o
 */
export function compareText(o) {
  const scale = o.scale ?? SCALE_125;
  const sep = o.sepPx ?? 3;
  const e = textInkPoints(o.eText, o.eOutline, sep);
  const i = textInkPoints(o.iText, o.iOutline, sep);
  if (e === null || i === null) return { orientation: 'unknown', reason: '枠線が取れない' };
  const sa = pointStats(e.pts, scale);
  const sb = pointStats(i.pts, scale);
  if (sa === null || sb === null || e.pts.length < 20 || i.pts.length < 20) {
    return { orientation: 'unknown', reason: '文字が取れない', excel: sa, interhub: sb };
  }
  // 向きは、文字のインクを各自の重心にそろえて比べる（位置の差・文字領域の非対称に左右されない）。180° 回した版は重心まわりの点対称
  const ca = textContour(e);
  const cb = textContour(i);
  const ea = { x: Math.round(sa.cx * scale), y: Math.round(sa.cy * scale) };
  const eb = { x: Math.round(sb.cx * scale), y: Math.round(sb.cy * scale) };
  const R = 400;
  const place = (pts, c, sign) => {
    const r = [];
    for (let k = 0; k < pts.length; k += 2) {
      const x = R + sign * (pts[k] - c.x);
      const y = R + sign * (pts[k + 1] - c.y);
      if (x >= 0 && y >= 0 && x < 2 * R && y < 2 * R) r.push(x, y);
    }
    return r;
  };
  const pa = place(ca, ea, 1);
  const p95 = (x, y) =>
    x.length === 0 || y.length === 0
      ? 999
      : percentile(
          pointDistances(2 * R, 2 * R, x, y).map((v) => v / scale),
          0.95,
        );
  const p95Upright = p95(pa, place(cb, eb, 1));
  const p95Rot180 = p95(pa, place(cb, eb, -1));
  return {
    orientation: p95Upright <= p95Rot180 ? 'upright' : 'rot180',
    p95Upright,
    p95Rot180,
    excel: sa,
    interhub: sb,
    dCentroid: { x: sb.cx - sa.cx, y: sb.cy - sa.cy },
    nA: e.pts.length / 2,
    nB: i.pts.length / 2,
  };
}

// ---------- 名前・集計 ----------

/**
 * ref-geoms.xlsx の図形名 → 比較の群。'preset'（AutoShapeType ごとの既定の図形）・'adjusted'（調整値違い）・'xform'（回転・反転）・
 * 'custom'（自由図形）・'default'（既定の塗り・枠線の図形）・'skip'（文字入り・枠線の太さ・線種の図形など、輪郭の比較の対象外）。
 */
export function groupOfShape(name) {
  if (/^p\d+$/.test(name)) return 'preset';
  if (/^adj/.test(name)) return 'adjusted';
  if (/^(rot\d+|flip)/.test(name)) return 'xform';
  if (/^(ff|curve|poly)/.test(name)) return 'custom';
  if (name === 'emptyOval' || name === 'defaultRect') return 'default';
  return 'skip';
}

/**
 * 比較の結果の集計（`entries`＝`{ name, group, p95 }` の列）。許容以内の割合・群ごとの数・超えた図形（p95 の大きい順）。
 * @param {{ name: string, group: string, p95: number }[]} entries
 */
export function summarize(entries, tolerance = TOLERANCE_PX) {
  const byGroup = {};
  const ng = [];
  let ok = 0;
  for (const e of entries) {
    const g = (byGroup[e.group] ??= { total: 0, ok: 0 });
    g.total++;
    if (e.p95 <= tolerance) {
      g.ok++;
      ok++;
    } else ng.push(e);
  }
  ng.sort((x, y) => y.p95 - x.p95);
  return {
    total: entries.length,
    ok,
    ng: ng.length,
    okRatio: entries.length === 0 ? 1 : ok / entries.length,
    byGroup,
    ngNames: ng.map((e) => e.name),
  };
}
