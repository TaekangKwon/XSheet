/**
 * 図形の輪郭の比較器（`fixtures/excel/geom-compare.mjs`・ISSUE-93）の単体テスト。
 * Excel の PNG・Playwright の写しが無くても動くよう、合成画像（輪郭の輪・楕円・矩形・F 字）で次を固定する：
 * PNG の読み書き（フィルタ 0〜4・パレット・1bit・16bit）・背景（マゼンタ）つきの画像からの矩形とインクの取り出し・距離変換・
 * 輪郭の距離の 95 パーセンタイル（位置ずれは左上そろえで消える・太さの差は許容内・半径の差は許容超え）・文字の向きの判定（U4）・図形名の群分け・集計。
 */
import { describe, expect, it } from 'vitest';
import {
  SCALE_125,
  TOLERANCE_PX,
  compareContours,
  compareSkeletons,
  compareText,
  contourPoints,
  cropToInk,
  decodePng,
  dilate,
  distanceField,
  encodePng,
  encodePngRaw,
  groupOfShape,
  inkMask,
  maskBox,
  overlayImage,
  percentile,
  resizeMask,
  pictureRect,
  pointDistances,
  summarize,
  thin,
  type Mask,
  type RgbaImage,
} from '../../fixtures/excel/geom-compare.mjs';

/** 空のマスク */
const blank = (w: number, h: number): Mask => ({ w, h, data: new Uint8Array(w * h) });
/** 条件を満たす画素をインクにしたマスク */
function draw(w: number, h: number, on: (x: number, y: number) => boolean): Mask {
  const m = blank(w, h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) m.data[y * w + x] = on(x, y) ? 1 : 0;
  return m;
}
/** 楕円の輪（半径 rx・ry・太さ t。中心 cx・cy） */
const ring = (
  w: number,
  h: number,
  cx: number,
  cy: number,
  rx: number,
  ry: number,
  t: number,
): Mask =>
  draw(w, h, (x, y) => {
    const d = Math.hypot((x + 0.5 - cx) / rx, (y + 0.5 - cy) / ry);
    return Math.abs(d - 1) * Math.min(rx, ry) <= t / 2;
  });
/** 矩形の枠（外形 w×h・太さ t） */
const frame = (w: number, h: number, t: number): Mask =>
  draw(w, h, (x, y) => x < t || y < t || x >= w - t || y >= h - t);
/** マスクをもとに RGBA 画像（インクは fg・地は bg） */
function image(
  m: Mask,
  fg: number[],
  bg: number[],
  padX = 0,
  padY = 0,
  pad: number[] = [255, 0, 255],
): RgbaImage {
  const width = m.w + 2 * padX;
  const height = m.h + 2 * padY;
  const data = new Uint8Array(width * height * 4);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const inside = x >= padX && x < padX + m.w && y >= padY && y < padY + m.h;
      const c = !inside ? pad : m.data[(y - padY) * m.w + (x - padX)] ? fg : bg;
      data.set([...c, 255], (y * width + x) * 4);
    }
  }
  return { width, height, data };
}

describe('TD-93-2 PNG の読み書き', () => {
  it('RGBA の書き出し → 読み込みで元に戻る', () => {
    const rgba = new Uint8Array(3 * 2 * 4).map((_, i) => (i * 37) & 255);
    const img = decodePng(encodePng(3, 2, rgba));
    expect([img.width, img.height]).toEqual([3, 2]);
    expect(Array.from(img.data)).toEqual(Array.from(rgba));
  });

  it('フィルタ 0〜4（なし・Sub・Up・Average・Paeth）の RGB 8bit を正しく読む', () => {
    const w = 7;
    const h = 5;
    const px = new Uint8Array(w * h * 3).map((_, i) => (i * 53 + (i >> 3) * 11) & 255);
    const bpp = 3;
    const stride = w * bpp;
    const raw = new Uint8Array(h * (stride + 1));
    for (let y = 0; y < h; y++) {
      const f = y % 5;
      raw[y * (stride + 1)] = f;
      for (let x = 0; x < stride; x++) {
        const cur = px[y * stride + x]!;
        const a = x >= bpp ? px[y * stride + x - bpp]! : 0;
        const up = y > 0 ? px[(y - 1) * stride + x]! : 0;
        const c = x >= bpp && y > 0 ? px[(y - 1) * stride + x - bpp]! : 0;
        let pred = 0;
        if (f === 1) pred = a;
        else if (f === 2) pred = up;
        else if (f === 3) pred = (a + up) >> 1;
        else if (f === 4) {
          const pa = Math.abs(up - c);
          const pb = Math.abs(a - c);
          const pc = Math.abs(a + up - 2 * c);
          pred = pa <= pb && pa <= pc ? a : pb <= pc ? up : c;
        }
        raw[y * (stride + 1) + 1 + x] = (cur - pred) & 255;
      }
    }
    const img = decodePng(encodePngRaw({ width: w, height: h, colorType: 2, bitDepth: 8, raw }));
    for (let i = 0; i < w * h; i++) {
      expect(Array.from(img.data.subarray(i * 4, i * 4 + 4))).toEqual([
        px[i * 3],
        px[i * 3 + 1],
        px[i * 3 + 2],
        255,
      ]);
    }
  });

  it('パレット 4bit（Excel の Chart.Export の形式）と tRNS の α・グレー 1bit・RGB 16bit を読む', () => {
    // 4bit パレット：幅 3（1 行 2 バイト＝0x01 0x20 → 0, 1, 2）
    const plte = new Uint8Array([255, 0, 255, 255, 255, 255, 10, 20, 30]);
    const idx = encodePngRaw({
      width: 3,
      height: 1,
      colorType: 3,
      bitDepth: 4,
      raw: new Uint8Array([0, 0x01, 0x20]),
      plte,
      trns: new Uint8Array([255, 255, 128]),
    });
    const a = decodePng(idx);
    expect(Array.from(a.data)).toEqual([255, 0, 255, 255, 255, 255, 255, 255, 10, 20, 30, 128]);
    // グレー 1bit：幅 3 → 0b101 → 255, 0, 255
    const g = decodePng(
      encodePngRaw({
        width: 3,
        height: 1,
        colorType: 0,
        bitDepth: 1,
        raw: new Uint8Array([0, 0b10100000]),
      }),
    );
    expect(Array.from(g.data)).toEqual([255, 255, 255, 255, 0, 0, 0, 255, 255, 255, 255, 255]);
    // RGB 16bit：上位 1 バイトを使う
    const rgb16 = decodePng(
      encodePngRaw({
        width: 1,
        height: 1,
        colorType: 2,
        bitDepth: 16,
        raw: new Uint8Array([0, 0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc]),
      }),
    );
    expect(Array.from(rgb16.data)).toEqual([0x12, 0x56, 0x9a, 255]);
  });

  it('PNG でないものは例外', () => {
    expect(() => decodePng(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9]))).toThrow();
  });
});

describe('TD-93-3 Excel の画像（マゼンタの背景＋白い矩形）からの矩形とインクの取り出し', () => {
  const O = ring(50, 40, 25, 20, 22, 17, 2);
  // 薄い青（Excel の既定の枠線色）の輪が、白い 50×40 の矩形に入り、周りはマゼンタ
  const excel = image(O, [0x41, 0x71, 0x9c], [255, 255, 255], 20, 10);

  it('pictureRect：白い矩形（背景のマゼンタを除く）の外接矩形', () => {
    expect(pictureRect(excel)).toEqual({ x0: 20, y0: 10, x1: 69, y1: 49 });
    expect(pictureRect(image(blank(4, 4), [0, 0, 0], [255, 0, 255], 0, 0))).toBeNull();
  });

  it('inkMask：薄い色の枠もインクとして取れる（閾値は最大の濃さの半分）。地の白はインクでない', () => {
    const m = inkMask(excel, pictureRect(excel));
    expect([m.w, m.h]).toEqual([50, 40]);
    expect(Array.from(m.data)).toEqual(Array.from(O.data));
  });

  it('inkMask：何も描かれていない画像は空（ノイズ程度の薄さは無視する）', () => {
    const m = inkMask(image(blank(10, 10), [250, 250, 250], [255, 255, 255]));
    expect(maskBox(m)).toBeNull();
    expect(cropToInk(m)).toEqual({ w: 0, h: 0, data: new Uint8Array(0) });
  });

  it('inkMask：α は白との混色（半透明の黒はグレー）。矩形の指定で切り出せる', () => {
    const img: RgbaImage = {
      width: 2,
      height: 1,
      data: new Uint8Array([0, 0, 0, 255, 0, 0, 0, 40]),
    };
    expect(Array.from(inkMask(img).data)).toEqual([1, 0]);
    expect(inkMask(img, { x0: 1, y0: 0, x1: 1, y1: 0 }).w).toBe(1);
  });

  it('maskBox・cropToInk：インクの外接矩形で切り出す', () => {
    const m = draw(10, 10, (x, y) => x >= 3 && x <= 5 && y >= 2 && y <= 6);
    expect(maskBox(m)).toEqual({ x0: 3, y0: 2, x1: 5, y1: 6 });
    const c = cropToInk(m);
    expect([c.w, c.h]).toEqual([3, 5]);
    expect(c.data.every((v) => v === 1)).toBe(true);
  });

  it('contourPoints：塗りつぶしの矩形の輪郭は外周の画素だけ。dilate は 1 画素膨らませる', () => {
    const m = draw(7, 7, (x, y) => x >= 1 && x <= 5 && y >= 1 && y <= 5);
    expect(contourPoints(m).length / 2).toBe(16);
    expect(maskBox(dilate(m, 1))).toEqual({ x0: 0, y0: 0, x1: 6, y1: 6 });
  });
});

describe('TD-93-4 距離変換・パーセンタイル', () => {
  it('distanceField：総当たりの最近点距離に一致する（乱数の点 12 個）', () => {
    const w = 23;
    const h = 17;
    let seed = 7;
    const rnd = (n: number) => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff), seed % n);
    const pts: number[] = [];
    for (let i = 0; i < 12; i++) pts.push(rnd(w), rnd(h));
    const f = distanceField(w, h, pts);
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        let best = Infinity;
        for (let i = 0; i < pts.length; i += 2)
          best = Math.min(best, Math.hypot(x - pts[i]!, y - pts[i + 1]!));
        expect(f[y * w + x]).toBeCloseTo(best, 9);
      }
    }
  });

  it('点が無ければ Infinity。pointDistances は双方向を昇順で返す', () => {
    expect(distanceField(3, 3, [])[4]).toBe(Infinity);
    expect(pointDistances(10, 10, [0, 0], [3, 4])).toEqual([5, 5]);
  });

  it('percentile：最近順位法（20 個の 95% は 19 番目）。空は 0', () => {
    const v = Array.from({ length: 20 }, (_, i) => i + 1);
    expect(percentile(v, 0.95)).toBe(19);
    expect(percentile(v, 1)).toBe(20);
    expect(percentile(v, 0)).toBe(1);
    expect(percentile([], 0.95)).toBe(0);
  });
});

describe('TD-93-5 輪郭の距離（p95・100% 換算の px）', () => {
  const A = cropToInk(ring(90, 70, 45, 35, 40, 30, 2));

  it('同じ形は 0。置かれた位置が違っても（外接矩形の左上そろえで）0', () => {
    expect(compareContours(A, A).p95).toBe(0);
    const moved = cropToInk(ring(120, 100, 70, 55, 40, 30, 2));
    const r = compareContours(A, moved);
    expect(r.p95).toBe(0);
    expect(r.dw).toBe(0);
    expect(r.dh).toBe(0);
  });

  it('枠線の太さが 2 と 3 画素の違いは許容（1.5px）以内。0.8 倍に縮小した形は許容を超える', () => {
    const thick = cropToInk(ring(90, 70, 45, 35, 40, 30, 3));
    expect(compareContours(A, thick).p95).toBeLessThanOrEqual(TOLERANCE_PX);
    const small = cropToInk(ring(90, 70, 45, 35, 32, 24, 2));
    expect(compareContours(A, small).p95).toBeGreaterThan(TOLERANCE_PX);
  });

  it('半径が 3 画素大きい円は許容を超える（寸法の差 dw・dh が出る）。楕円と矩形は大きく違う', () => {
    const c1 = cropToInk(ring(100, 100, 50, 50, 40, 40, 2));
    const c2 = cropToInk(ring(100, 100, 50, 50, 43, 43, 2));
    const r = compareContours(c1, c2);
    expect(r.p95).toBeGreaterThan(TOLERANCE_PX);
    expect(r.dw).toBeCloseTo(6 / SCALE_125, 5);
    expect(r.dh).toBeCloseTo(6 / SCALE_125, 5);
    expect(compareContours(c1, cropToInk(frame(80, 80, 2))).p95).toBeGreaterThan(TOLERANCE_PX);
  });

  it('scale を変えると距離は反比例する（scale 2 は 1.25 の 0.625 倍）', () => {
    const c1 = cropToInk(ring(100, 100, 50, 50, 40, 40, 2));
    const c2 = cropToInk(ring(100, 100, 50, 50, 43, 43, 2));
    const a = compareContours(c1, c2, 1.25).p95;
    const b = compareContours(c1, c2, 2).p95;
    expect(b / a).toBeCloseTo(1.25 / 2, 9);
  });

  it('片方が空：p95 は 999 と empty（どちらも空なら 0）', () => {
    const none = { w: 0, h: 0, data: new Uint8Array(0) };
    expect(compareContours(none, A)).toMatchObject({ p95: 999, empty: 'excel' });
    expect(compareContours(A, none)).toMatchObject({ p95: 999, empty: 'interhub' });
    expect(compareContours(none, none)).toMatchObject({ p95: 0, empty: 'both' });
  });

  it('overlayImage：a のみ＝赤・b のみ＝青・両方＝黒・どちらも無い＝白', () => {
    const a = draw(3, 1, (x) => x <= 1);
    const b = draw(3, 1, (x) => x >= 1);
    const o = overlayImage(a, b);
    expect(Array.from(o.data.subarray(0, 4))).toEqual([230, 40, 40, 255]);
    expect(Array.from(o.data.subarray(4, 8))).toEqual([0, 0, 0, 255]);
    expect(Array.from(o.data.subarray(8, 12))).toEqual([40, 40, 230, 255]);
    // PNG として書ける
    expect(decodePng(encodePng(o.width, o.height, o.data)).width).toBe(3);
  });
});

describe('TD-93-6 反転した図形の文字（U4）：向きと位置', () => {
  const W = 120;
  const H = 80;
  const O = frame(W, H, 2);
  /** F 字（縦棒 4×20・上の横棒 14×4・中の横棒 10×4）を (ox, oy) に置く。180° 回すと (W−1−x, H−1−y) */
  const fGlyph = (ox: number, oy: number, rot180: boolean) =>
    draw(W, H, (x, y) => {
      const px = rot180 ? W - 1 - x : x;
      const py = rot180 ? H - 1 - y : y;
      const lx = px - ox;
      const ly = py - oy;
      return (
        (lx >= 0 && lx < 4 && ly >= 0 && ly < 20) ||
        (lx >= 0 && lx < 14 && ly >= 0 && ly < 4) ||
        (lx >= 0 && lx < 10 && ly >= 8 && ly < 12)
      );
    });
  const union = (a: Mask, b: Mask): Mask =>
    draw(W, H, (x, y) => a.data[y * W + x] === 1 || b.data[y * W + x] === 1);
  const upright = union(O, fGlyph(10, 10, false));
  const flipped = union(O, fGlyph(10, 10, true));

  it('同じ向き・同じ位置なら upright。距離 0・重心の差 0', () => {
    const r = compareText({ eText: upright, eOutline: O, iText: upright, iOutline: O });
    expect(r.orientation).toBe('upright');
    expect(r.p95Upright).toBe(0);
    expect(r.dCentroid).toEqual({ x: 0, y: 0 });
    expect(r.p95Rot180!).toBeGreaterThan(TOLERANCE_PX);
  });

  it('Excel の文字が 180° 回っている（InterHub は回していない）と rot180。重心は中心の反対側へ動く', () => {
    const r = compareText({ eText: flipped, eOutline: O, iText: upright, iOutline: O });
    expect(r.orientation).toBe('rot180');
    expect(r.p95Rot180!).toBeLessThanOrEqual(1);
    expect(r.p95Upright!).toBeGreaterThan(TOLERANCE_PX);
    expect(r.dCentroid!.x).toBeLessThan(0);
    expect(r.dCentroid!.y).toBeLessThan(0);
  });

  it('文字の位置だけが違う（上下に 30 画素）と upright のまま重心の差に出る', () => {
    const low = union(O, fGlyph(10, 40, false));
    const r = compareText({ eText: low, eOutline: O, iText: upright, iOutline: O });
    expect(r.orientation).toBe('upright');
    expect(r.dCentroid!.y).toBeCloseTo(-30 / SCALE_125, 5);
    expect(r.dCentroid!.x).toBeCloseTo(0, 5);
  });

  it('文字が無い・枠線が無いときは unknown', () => {
    expect(compareText({ eText: O, eOutline: O, iText: upright, iOutline: O }).orientation).toBe(
      'unknown',
    );
    expect(
      compareText({ eText: upright, eOutline: blank(W, H), iText: upright, iOutline: O })
        .orientation,
    ).toBe('unknown');
  });
});

describe('TD-93-7 図形名の群・集計', () => {
  it('groupOfShape：ref-geoms.xlsx の図形名を群に分ける', () => {
    expect(groupOfShape('p9')).toBe('preset');
    expect(groupOfShape('adjCan')).toBe('adjusted');
    expect(['rot30', 'flipH', 'flipHVRot270'].map(groupOfShape)).toEqual([
      'xform',
      'xform',
      'xform',
    ]);
    expect(['ffRot', 'curve7', 'polyOpen'].map(groupOfShape)).toEqual([
      'custom',
      'custom',
      'custom',
    ]);
    expect(['emptyOval', 'defaultRect'].map(groupOfShape)).toEqual(['default', 'default']);
    expect(
      ['textOval', 'w_0', 'd_dash', 'lineNone', 'emptyDiamond', 'tbBordered'].map(groupOfShape),
    ).toEqual(Array(6).fill('skip'));
  });

  it('summarize：許容以内の数・群ごとの数・超えた図形は p95 の大きい順', () => {
    const s = summarize([
      { name: 'a', group: 'preset', p95: 0.5 },
      { name: 'b', group: 'preset', p95: 2.5 },
      { name: 'c', group: 'adjusted', p95: 1.5 },
      { name: 'd', group: 'adjusted', p95: 9 },
    ]);
    expect(s).toEqual({
      total: 4,
      ok: 2,
      ng: 2,
      okRatio: 0.5,
      byGroup: { preset: { total: 2, ok: 1 }, adjusted: { total: 2, ok: 1 } },
      ngNames: ['d', 'b'],
    });
    expect(summarize([]).okRatio).toBe(1);
  });
});

describe('TD-93-8 細線化・中心線の比較（枠線の太さ・鋭角の張り出し・寸法の丸めに左右されない）', () => {
  const ringA = ring(100, 80, 50, 40, 44, 34, 2);

  it('thin：太い輪（4 画素）は 1 画素幅の輪になり、半径は保たれる。空は空', () => {
    const thick = ring(100, 80, 50, 40, 44, 34, 4);
    const t = thin(thick);
    const before = thick.data.reduce((s, v) => s + v, 0);
    const after = t.data.reduce((s, v) => s + v, 0);
    expect(after).toBeLessThan(before / 2);
    // 細線化後の画素はすべて、元の輪の画素の上にある
    for (let i = 0; i < t.data.length; i++) if (t.data[i]) expect(thick.data[i]).toBe(1);
    // 中心線の外接矩形は、元の輪の外接矩形より内側（太さの半分ほど）
    const bt = maskBox(t)!;
    const bo = maskBox(thick)!;
    expect(bt.x0 - bo.x0).toBeGreaterThanOrEqual(1);
    expect(bt.x0 - bo.x0).toBeLessThanOrEqual(4);
    expect(bo.x1 - bt.x1).toBeGreaterThanOrEqual(1);
    expect(bo.y1 - bt.y1).toBeLessThanOrEqual(4);
    expect(thin(blank(5, 5)).data.every((v) => v === 0)).toBe(true);
  });

  it('枠線の太さが違っても中心線は一致する（2 画素と 5 画素の輪は p95 が許容内）', () => {
    const thick = ring(100, 80, 50, 40, 44, 34, 5);
    expect(compareSkeletons(ringA, thick, SCALE_125, false).p95).toBeLessThanOrEqual(TOLERANCE_PX);
  });

  it('searchPx：外接矩形が鋭角の張り出しでずれても、位置合わせで全体の距離が小さくなる。動かした量を shift に返す', () => {
    // b＝a に左へ 7 画素の突起（miter の張り出しの代わり）。外接矩形の左上でそろえると輪全体が 7 画素ずれる
    const spike = draw(100, 80, (x, y) => ringA.data[y * 100 + x] === 1 || (y === 40 && x < 7));
    const aligned = compareSkeletons(ringA, spike, SCALE_125, false, 0);
    expect(aligned.p95).toBeGreaterThan(TOLERANCE_PX);
    const searched = compareSkeletons(ringA, spike, SCALE_125, false, 8);
    expect(searched.p95).toBeLessThanOrEqual(TOLERANCE_PX);
    // 突起は輪の左端より 6 画素ほど外へ出ている（細線化で ±1）
    expect(Math.abs(searched.shift.x) * SCALE_125).toBeGreaterThanOrEqual(5);
    expect(Math.abs(searched.shift.x) * SCALE_125).toBeLessThanOrEqual(7);
    expect(searched.shift.y).toBe(0);
  });

  it('normalize：b を a の寸法へ拡縮して測る。寸法の差 dw・dh は拡縮前の値を返す', () => {
    const big = ring(110, 88, 55, 44, 48.4, 37.4, 2);
    const keep = compareSkeletons(ringA, big, SCALE_125, false, 0);
    const fit = compareSkeletons(ringA, big, SCALE_125, true, 0);
    expect(keep.p95).toBeGreaterThan(TOLERANCE_PX);
    expect(fit.p95).toBeLessThanOrEqual(TOLERANCE_PX);
    expect(fit.dw).toBeCloseTo(keep.dw, 9);
    expect(fit.dw).toBeGreaterThan(0);
  });

  it('片方が空：p95 は 999 と empty（shift は 0）', () => {
    expect(compareSkeletons(blank(5, 5), ringA)).toMatchObject({
      p95: 999,
      empty: 'excel',
      shift: { x: 0, y: 0 },
    });
    expect(compareSkeletons(ringA, blank(5, 5))).toMatchObject({ p95: 999, empty: 'interhub' });
  });

  it('resizeMask：最近傍で拡縮する（2×2 → 4×4 は各画素が 2×2 になる）。空は空', () => {
    const m = resizeMask({ w: 2, h: 2, data: new Uint8Array([1, 0, 0, 1]) }, 4, 4);
    expect(Array.from(m.data)).toEqual([1, 1, 0, 0, 1, 1, 0, 0, 0, 0, 1, 1, 0, 0, 1, 1]);
    expect(resizeMask({ w: 0, h: 0, data: new Uint8Array(0) }, 3, 2).data).toEqual(
      new Uint8Array(6),
    );
  });
});
