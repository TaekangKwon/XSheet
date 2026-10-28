// ISSUE-95：make-ref-colors.ps1 の出力（com.json と png\*.png）から、太い線（3pt 以上）の筋の画素を取り出して com.json に足し、ref-colors.expected.json を作る
// 使い方：node make-ref-colors.extract.mjs <out dir>   <out dir>/com.json と <out dir>/png/ が入力。結果は <out dir>/expected.json（PNG はリポジトリに入れない）
// 事情：CopyPicture のビットマップは白い背景の矩形を含むので、ps1 の run（最初の筋）は線の内側の白まで走って使えない。ここで PNG を読み直し、
//       線の外縁から 3〜6px（太さ 6pt≒10px の中ほど）の画素の平均を stroke として記録する（矩形の枠線は左辺の中段、水平な線は中央の列）。
// 透過の基準は背景のマゼンタでなく白（貼られた絵の中の混色は白との混色）。PNG は 4/8bit のパレット形式で、このデコーダは標準の zlib だけで読む。
/* global process */
import fs from 'node:fs';
import zlib from 'node:zlib';
import { Buffer } from 'node:buffer';
const [, , dir] = process.argv;
const com = JSON.parse(fs.readFileSync(`${dir}/com.json`, 'utf8'));

/** PNG（パレット・RGB・RGBA、1〜8bit）を読み、(x, y) → [r, g, b] を返す */
function decode(path) {
  const b = fs.readFileSync(path);
  let p = 8,
    w,
    h,
    ct,
    bd,
    plte = null;
  const idat = [];
  while (p < b.length) {
    const len = b.readUInt32BE(p),
      t = b.toString('ascii', p + 4, p + 8),
      d = b.subarray(p + 8, p + 8 + len);
    if (t === 'IHDR') {
      w = d.readUInt32BE(0);
      h = d.readUInt32BE(4);
      bd = d[8];
      ct = d[9];
    }
    if (t === 'PLTE') plte = d;
    if (t === 'IDAT') idat.push(d);
    p += 12 + len;
  }
  const bitsPP = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[ct] * bd,
    bpp = Math.max(1, bitsPP >> 3);
  const raw = zlib.inflateSync(Buffer.concat(idat)),
    stride = Math.ceil((w * bitsPP) / 8),
    out = Buffer.alloc(h * stride);
  for (let y = 0; y < h; y++) {
    const f = raw[y * (stride + 1)],
      line = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1));
    for (let x = 0; x < stride; x++) {
      const a = x >= bpp ? out[y * stride + x - bpp] : 0,
        up = y > 0 ? out[(y - 1) * stride + x] : 0,
        c = x >= bpp && y > 0 ? out[(y - 1) * stride + x - bpp] : 0;
      let v = line[x];
      if (f === 1) v += a;
      else if (f === 2) v += up;
      else if (f === 3) v += (a + up) >> 1;
      else if (f === 4) {
        const pa = Math.abs(up - c),
          pb = Math.abs(a - c),
          pc = Math.abs(a + up - 2 * c);
        v += pa <= pb && pa <= pc ? a : pb <= pc ? up : c;
      }
      out[y * stride + x] = v & 255;
    }
  }
  return (x, y) => {
    if (ct === 3) {
      const bit = x * bd,
        idx = (out[y * stride + (bit >> 3)] >> (8 - bd - (bit & 7))) & ((1 << bd) - 1);
      return [plte[idx * 3], plte[idx * 3 + 1], plte[idx * 3 + 2]];
    }
    const i = y * stride + x * bpp;
    return bpp >= 3 ? [out[i], out[i + 1], out[i + 2]] : [out[i], out[i], out[i]];
  };
}
const hex = (c) => c.map((v) => v.toString(16).toUpperCase().padStart(2, '0')).join('');
const mean = (px, xs, ys) => {
  const s = [0, 0, 0];
  let n = 0;
  for (const y of ys)
    for (const x of xs) {
      px(x, y).forEach((v, i) => (s[i] += v));
      n++;
    }
  return s.map((v) => Math.round(v / n));
};
const range = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);

let stroked = 0;
for (const s of com.shapes) {
  delete s.run;
  if (!(['outline', 'default'].includes(s.group) && s.line.Visible === -1 && s.line.Weight >= 3))
    continue;
  const [, , x0, y0, x1, y1] = s.box;
  const px = decode(`${dir}/png/${s.name}.png`);
  const horizontalLine = /^(lnh_|defLine|defConnector)/.test(s.name);
  const rgb = horizontalLine
    ? mean(
        px,
        range(Math.floor((x0 + x1) / 2) - 1, Math.floor((x0 + x1) / 2) + 1),
        range(y0 + 3, y0 + 6),
      )
    : mean(
        px,
        range(x0 + 3, x0 + 6),
        range(Math.floor((y0 + y1) / 2) - 1, Math.floor((y0 + y1) / 2) + 1),
      );
  s.stroke = { axis: horizontalLine ? 'v' : 'h', rgb: hex(rgb) };
  stroked++;
}
com.baseNote = '貼られた絵は白い背景の矩形を含む。透過の混色の基準は白（bg のマゼンタではない）';
fs.writeFileSync(`${dir}/expected.json`, JSON.stringify(com, null, 1));
console.log('shapes', com.shapes.length, 'stroke', stroked);
