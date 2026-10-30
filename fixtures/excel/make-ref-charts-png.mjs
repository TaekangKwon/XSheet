// ISSUE-105（C4）：Excel が Chart.Export で書いた PNG から、目盛線・軸線の色（画素の代表値）を読んで JSON にする（依存は Node 標準と geom-compare.mjs の PNG 読みだけ）。
// COM の Format.Line は c:spPr の無いグラフで「黒 0.5pt」を返し、見た目（薄い灰色）と合わない（ADR-0031 §決定 4 の注）ため、色は画素から読む。
// 使い方：node make-ref-charts-png.mjs <out dir> [--out <出力ファイル名>] [--index <png-index の名前の正規表現>]
//   （ISSUE-110：3 回目は make-ref-charts3.ps1 の出力に対し --out ref-charts3-png.json で呼ぶ。既定は 2 回目と同じ ref-charts-png.json・png-index*.json で、2 回目の結果は変わらない）
//   入力：<out dir>/png/png-index*.json（make-ref-charts2.ps1 の png 段が書く。グラフごとの PNG ファイル名・大きさ pt・COM のプロット内側・目盛の範囲・系列の色）と、その PNG。
//   出力：<out dir>/ref-charts-png.json（グラフ名 → 線の一覧と役割ごとの代表値。PNG そのものは入れない＝コミットするのはこの JSON だけ）。
// 読み方（analyzeChartPng）：
//   1. 縮尺＝PNG の幅 ÷ ChartObject の幅（pt）。プロットの内側（COM の値＋原点 4pt）の周り ±5pt を走査する。背景＝プロットの内側で最も多い色（棒が過半を覆わない前提）。
//      （プロットの塗りが面と違う暗い背景〈c:style 42〉と、面の角丸の弧に当たる左上 10pt 内側の画素を避ける）。プロットの位置が無い時（円など）は左上 10pt 内側の画素。
//   2. 「線の画素」＝背景でも系列の色でもなく、上下（左右）d 画素離れた所が背景の画素。ここで棒の端（縁の混ざった色）を線と取り違えない。
//      行（列）ごとに最頻の色を数え、プロットの幅（高さ）の 12% 以上ある行（列）を線の行とし、隣り合う行を 1 本の線にまとめる。
//   3. 線の役割：値の方向の線（縦棒・折れ線なら横線）は、0（無ければ最小）の位置＝項目軸の線、主の目盛の位置＝主の目盛線、補助の位置＝補助の目盛線。
//      値の軸の線（縦棒なら左の縦線・横棒なら下の横線）は端の位置。
//   5. ISSUE-110（index に印があるグラフだけ。印の無いグラフの結果は 2 回目と同じ）：Corner＝グラフの面の角（scanCorner）・Ticks＝数値軸の目盛の印（scanTicks）・LabelBand＝項目軸の文字の塊の数と角度（scanLabels）。
//   4. ink＝線の行ごとの（背景 − 画素）の合計（チャンネルごと・0〜255 の画素数倍）。アンチエイリアスで細い線が薄まっても、線の太さ w px・色 c の線の ink は
//      w × (背景 − c) になる（面積の保存）ので、場面の線との比較は ink で行う（照合テスト xlsx-chart-ref2.test.ts）。
/* global process */
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { decodePng } from './geom-compare.mjs';

/** COM の座標の原点のずれ（pt）：グラフの面の左上から内側（xlsx-chart-ref.test.ts の COM_ORIGIN_PT と同じ） */
export const COM_ORIGIN_PT = 4;

const hex2 = (n) => n.toString(16).padStart(2, '0');
/** 0xRRGGBB → 'RRGGBB' */
export const toHex = (p) =>
  [(p >> 16) & 255, (p >> 8) & 255, p & 255].map(hex2).join('').toUpperCase();
const parseHex = (s) => parseInt(s, 16);
/** 2 色のチャンネルごとの差の最大 */
const chDiff = (a, b) =>
  Math.max(
    Math.abs(((a >> 16) & 255) - ((b >> 16) & 255)),
    Math.abs(((a >> 8) & 255) - ((b >> 8) & 255)),
    Math.abs((a & 255) - (b & 255)),
  );
const round2 = (x) => Math.round(x * 100) / 100;
const round1 = (x) => Math.round(x * 10) / 10;
const rgbAt = (img, x, y) => {
  const i = (y * img.width + x) * 4;
  return (img.data[i] << 16) | (img.data[i + 1] << 8) | img.data[i + 2];
};
const median = (a) => {
  if (a.length === 0) return null;
  const s = [...a].sort((p, q) => p - q);
  return s[s.length >> 1];
};

/**
 * グラフの面の左上の角を読む（ISSUE-110・角丸 c:roundedCorners の半径）。面の外側（Chart.Export の PNG では透明か背景色）の画素が、左上の隅から行ごとに何画素続くかを見る。
 * 角が直角なら (0,0) は枠線の色で、上の行の中ほどと同じ色（rounded=false・radiusPx=0）。角丸なら (0,0) は外側で、radiusPx＝外側の画素がある最後の行＋1。
 * @param {{ width: number, height: number, data: Uint8Array }} img decodePng の結果（RGBA）
 * @returns {{ rounded: boolean, transparent: boolean, outside: string | null, radiusPx: number, runs: number[] }}
 */
export function scanCorner(img) {
  const { width: W, height: H, data } = img;
  const alpha = (x, y) => data[(y * W + x) * 4 + 3];
  const c0 = rgbAt(img, 0, 0);
  const transparent = alpha(0, 0) < 128;
  const rounded = transparent || chDiff(c0, rgbAt(img, W >> 1, 0)) > 2;
  if (!rounded) return { rounded: false, transparent: false, outside: null, radiusPx: 0, runs: [] };
  const isOut = (x, y) =>
    transparent ? alpha(x, y) < 128 : alpha(x, y) >= 128 && chDiff(rgbAt(img, x, y), c0) <= 2;
  const runs = [];
  for (let y = 0; y < Math.min(H, 48); y++) {
    let n = 0;
    while (n < W && isOut(n, y)) n++;
    runs.push(n);
  }
  let last = -1;
  runs.forEach((n, y) => {
    if (n > 0) last = y;
  });
  return {
    rounded: true,
    transparent,
    outside: transparent ? null : toHex(c0),
    radiusPx: last + 1,
    runs: runs.slice(0, last + 2),
  };
}

/**
 * 項目軸の文字の塊を読む（ISSUE-110・107 の回転・折り返しの見極め）。範囲 region（px・両端を含む）の中で背景と違う画素を、g 画素以内で互いに近いもの同士まとめて塊にする（文字の字間・折り返しの行間は 1 つの塊）。
 * 目盛の印のような細い塊（画素 14 個以下で幅か高さが 2 画素以下）は除く。塊ごとに外接矩形と、画素の分散から出した主軸の角度（度・0＝水平・正＝右上がり・±90＝縦）を返す。
 * @param {{ width: number, height: number, data: Uint8Array }} img
 * @param {number} bg 背景の色（0xRRGGBB）
 * @param {{ x0: number, y0: number, x1: number, y1: number }} region
 * @param {number} g 同じ塊とみなす画素の距離
 */
export function scanLabels(img, bg, region, g) {
  const x0 = Math.max(0, region.x0);
  const y0 = Math.max(0, region.y0);
  const x1 = Math.min(img.width - 1, region.x1);
  const y1 = Math.min(img.height - 1, region.y1);
  const w = x1 - x0 + 1;
  const h = y1 - y0 + 1;
  if (w <= 0 || h <= 0)
    return {
      region: { x0, y0, x1, y1 },
      count: 0,
      clusters: [],
      medianW: null,
      medianH: null,
      medianAngle: null,
    };
  const parent = new Int32Array(w * h).fill(-1);
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++)
      if (chDiff(rgbAt(img, x0 + x, y0 + y), bg) > 30) parent[y * w + x] = y * w + x;
  const find = (i) => {
    let r = i;
    while (parent[r] !== r) r = parent[r];
    while (parent[i] !== r) {
      const nx = parent[i];
      parent[i] = r;
      i = nx;
    }
    return r;
  };
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      if (parent[i] < 0) continue;
      for (let dy = 0; dy <= g; dy++) {
        for (let dx = -g; dx <= g; dx++) {
          if (dy === 0 && dx <= 0) continue;
          const nx = x + dx;
          const ny = y + dy;
          if (nx < 0 || nx >= w || ny >= h) continue;
          const j = ny * w + nx;
          if (parent[j] < 0) continue;
          const a = find(i);
          const b = find(j);
          if (a !== b) parent[b] = a;
        }
      }
    }
  }
  const acc = new Map();
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      if (parent[i] < 0) continue;
      const r = find(i);
      let c = acc.get(r);
      if (c === undefined)
        acc.set(
          r,
          (c = { n: 0, x0: x, y0: y, x1: x, y1: y, sx: 0, sy: 0, sxx: 0, syy: 0, sxy: 0 }),
        );
      c.n++;
      c.x0 = Math.min(c.x0, x);
      c.x1 = Math.max(c.x1, x);
      c.y0 = Math.min(c.y0, y);
      c.y1 = Math.max(c.y1, y);
      c.sx += x;
      c.sy += y;
      c.sxx += x * x;
      c.syy += y * y;
      c.sxy += x * y;
    }
  }
  const clusters = [];
  for (const c of acc.values()) {
    const bw = c.x1 - c.x0 + 1;
    const bh = c.y1 - c.y0 + 1;
    if (c.n <= 14 && (bw <= 2 || bh <= 2)) continue;
    const mx = c.sx / c.n;
    const my = c.sy / c.n;
    const vx = c.sxx / c.n - mx * mx;
    const vy = c.syy / c.n - my * my;
    const cov = c.sxy / c.n - mx * my;
    // 画像の座標は y が下向きなので、右上がりの文字は cov < 0。角度は反時計回りを正にする
    const angle = round1((-0.5 * Math.atan2(2 * cov, vx - vy) * 180) / Math.PI);
    clusters.push({
      x0: x0 + c.x0,
      y0: y0 + c.y0,
      x1: x0 + c.x1,
      y1: y0 + c.y1,
      w: bw,
      h: bh,
      n: c.n,
      angle,
    });
  }
  clusters.sort((a, b) => a.x0 - b.x0 || a.y0 - b.y0);
  return {
    region: { x0, y0, x1, y1 },
    count: clusters.length,
    clusters: clusters.slice(0, 80),
    medianW: median(clusters.map((c) => c.w)),
    medianH: median(clusters.map((c) => c.h)),
    medianAngle: median(clusters.map((c) => c.angle)),
  };
}

/**
 * 数値軸の目盛の印を読む（ISSUE-110・数値軸の線が none の時の補助の目盛の印の有無）。数値軸（縦の軸は左端・横の軸は下端）から外側・内側へ、
 * 背景と違う画素が何画素続くかを軸に沿った位置ごとに数え、cap 画素に届かない短い連続（2 画素以上）を印とする（目盛線・ほかの線は cap 以上で除かれる）。
 * 隣り合う位置は 1 本の印にまとめ、主の目盛・補助の目盛の位置（px）に近いものを数える（主を先に取る）。
 * @param {{ width: number, height: number, data: Uint8Array }} img
 * @param {number} bg 背景の色
 * @param {{ vert: boolean, x0: number, y0: number, x1: number, y1: number, scale: number, majorPos: number[], minorPos: number[], tol: number }} geo
 * @returns {{ cap: number, out: object, in: object }} 側ごとに { major, minor, other, majorLen, minorLen, list }
 */
export function scanTicks(img, bg, geo) {
  const { width: W, height: H } = img;
  const { vert, x0, y0, x1, y1, scale, majorPos, minorPos, tol } = geo;
  const cap = Math.ceil(7 * scale);
  const isInk = (x, y) => x >= 0 && y >= 0 && x < W && y < H && chDiff(rgbAt(img, x, y), bg) > 30;
  const axis = vert ? x0 : y1;
  const [lo, hi] = vert ? [y0, y1] : [x0, x1];
  const side = (sgn) => {
    const rows = [];
    for (let p = lo - 2; p <= hi + 2; p++) {
      const at = (s) => (vert ? isInk(axis + sgn * s, p) : isInk(p, axis + sgn * s));
      let s0 = 0;
      for (const s of [1, 2]) {
        if (at(s)) {
          s0 = s;
          break;
        }
      }
      let len = 0;
      if (s0 > 0) {
        len = s0;
        while (len < cap && at(len + 1)) len++;
      }
      rows.push({ p, len });
    }
    // cap に届いた長い連続（目盛線など）と、1 画素以下は印にしない。隣り合う位置は 1 本にまとめる
    const ticks = [];
    let cur = null;
    for (const r of rows) {
      if (r.len >= 2 && r.len < cap) {
        if (cur !== null && r.p - cur.last <= 1) {
          cur.last = r.p;
          cur.len = Math.max(cur.len, r.len);
          cur.sum += r.p;
          cur.n++;
        } else {
          cur = { first: r.p, last: r.p, len: r.len, sum: r.p, n: 1 };
          ticks.push(cur);
        }
      } else cur = null;
    }
    const res = { major: 0, minor: 0, other: 0, majorLen: 0, minorLen: 0, list: [] };
    for (const t of ticks) {
      const pos = t.sum / t.n;
      const near = (arr) => arr.some((v) => Math.abs(v - pos) <= tol);
      const kind = near(majorPos) ? 'major' : near(minorPos) ? 'minor' : 'other';
      res[kind]++;
      if (kind !== 'other') res[kind + 'Len'] = Math.max(res[kind + 'Len'], t.len);
      if (res.list.length < 40) res.list.push({ pos: round2(pos), len: t.len, kind });
    }
    return res;
  };
  // 外側＝縦の軸は左・横の軸は下。内側はその逆
  return { cap, out: side(vert ? -1 : 1), in: side(vert ? 1 : -1) };
}

/**
 * 1 枚の PNG から線を読む。
 * @param {{ width: number, height: number, data: Uint8Array }} img decodePng の結果（RGBA）
 * @param {object} entry png-index の 1 件（WidthPt・PlotInside・HasAxes・Horizontal・ValueAxis・ExcludeRGB）
 * @param {{ d?: number, minFraction?: number, marginPt?: number }} [opts] d＝上下の背景を確かめる距離 px（既定 max(3, ceil(2×縮尺))）・minFraction＝線の行とする割合（0.12）
 * @returns {object} { width, height, scale, bg, hasAxes, plotPx?, lines[], summary }
 */
export function analyzeChartPng(img, entry, opts = {}) {
  const { width: W, height: H, data } = img;
  const scale = W / entry.WidthPt;
  const at = (x, y) => {
    const i = (y * W + x) * 4;
    return (data[i] << 16) | (data[i + 1] << 8) | data[i + 2];
  };
  const pi = entry.PlotInside;
  const hasPlot =
    entry.HasAxes === true &&
    pi &&
    [pi.Left, pi.Top, pi.Width, pi.Height].every((v) => typeof v === 'number');
  // 除く色＝系列の色。COM の Format.Line が c:spPr の無い系列で返す黒（棒に枠線は無い）は除かない（黒の目盛線を落とさない）
  const excl = (entry.ExcludeRGB ?? [])
    .filter((s) => typeof s === 'string' && s.toUpperCase() !== '000000')
    .map(parseHex);
  let bg;
  if (hasPlot) {
    // プロットの内側（縁の 3 画素内側）の画素の最頻の色（2 画素おきに数える。ExcludeRGB は白の枠線も含むので除かない）
    const hist = new Map();
    const bx0 = Math.max(0, Math.round((pi.Left + COM_ORIGIN_PT) * scale) + 3);
    const bx1 = Math.min(W - 1, Math.round((pi.Left + COM_ORIGIN_PT + pi.Width) * scale) - 3);
    const by0 = Math.max(0, Math.round((pi.Top + COM_ORIGIN_PT) * scale) + 3);
    const by1 = Math.min(H - 1, Math.round((pi.Top + COM_ORIGIN_PT + pi.Height) * scale) - 3);
    for (let y = by0; y <= by1; y += 2)
      for (let x = bx0; x <= bx1; x += 2) {
        const c = at(x, y);
        hist.set(c, (hist.get(c) ?? 0) + 1);
      }
    let best = null;
    for (const [c, n] of hist) if (best === null || n > best.n) best = { c, n };
    bg =
      best === null
        ? at(Math.min(W - 1, Math.round(10 * scale)), Math.min(H - 1, Math.round(10 * scale)))
        : best.c;
  } else bg = at(Math.min(W - 1, Math.round(10 * scale)), Math.min(H - 1, Math.round(10 * scale)));
  const res = {
    width: W,
    height: H,
    scale: round2(scale),
    scaleY: round2(H / entry.HeightPt),
    bg: toHex(bg),
    hasAxes: entry.HasAxes === true,
    lines: [],
    summary: {},
  };
  if (entry.Corner === true) res.corner = scanCorner(img);
  if (!res.hasAxes || !hasPlot) return res;

  const x0 = Math.round((pi.Left + COM_ORIGIN_PT) * scale);
  const y0 = Math.round((pi.Top + COM_ORIGIN_PT) * scale);
  const x1 = Math.round((pi.Left + COM_ORIGIN_PT + pi.Width) * scale);
  const y1 = Math.round((pi.Top + COM_ORIGIN_PT + pi.Height) * scale);
  res.plotPx = { x0, y0, x1, y1 };
  const d = opts.d ?? Math.max(3, Math.ceil(2 * scale));
  const margin = Math.ceil((opts.marginPt ?? 5) * scale);
  const minFraction = opts.minFraction ?? 0.12;
  const inset = 3;
  const isBg = (c) => chDiff(c, bg) <= 2;
  const isEx = (c) => excl.some((e) => chDiff(e, c) <= 2);

  /** 向き 'h'（行）か 'v'（列）の線の行・列を探す。返すのは { p, color, count }[] */
  const scan = (dir) => {
    const horizontal = dir === 'h';
    const [pLo, pHi] = horizontal ? [y0 - margin, y1 + margin] : [x0 - margin, x1 + margin];
    const [sLo, sHi] = horizontal ? [x0 + inset, x1 - inset] : [y0 + inset, y1 - inset];
    const pMax = horizontal ? H - 1 : W - 1;
    const get = (p, s) => (horizontal ? at(s, p) : at(p, s));
    const need = minFraction * (sHi - sLo + 1);
    // 系列の色でも、プロットの幅の 9 割以上を通る行は線（c:spPr の無いグラフの目盛線は黒＝COM が返す隠れた枠線の色と同じ）
    const full = 0.9 * (sHi - sLo + 1);
    const rows = [];
    for (let p = Math.max(d, pLo); p <= Math.min(pMax - d, pHi); p++) {
      const hist = new Map();
      for (let s = Math.max(0, sLo); s <= sHi; s++) {
        const c = get(p, s);
        if (isBg(c)) continue;
        if (!isBg(get(p - d, s)) || !isBg(get(p + d, s))) continue;
        hist.set(c, (hist.get(c) ?? 0) + 1);
      }
      let best = null;
      for (const [c, n] of hist)
        if ((!isEx(c) || n >= full) && (best === null || n > best.n)) best = { c, n };
      if (best !== null && best.n >= need) rows.push({ p, color: best.c, count: best.n });
    }
    return rows;
  };

  /** 隣り合う行を 1 本の線にまとめる。ink＝（背景 − 画素）の行ごとの合計 */
  const group = (rows, dir) => {
    const out = [];
    let cur = null;
    for (const r of rows) {
      if (cur === null || r.p - cur[cur.length - 1].p > 1) {
        cur = [];
        out.push(cur);
      }
      cur.push(r);
    }
    return out.map((rs) => {
      const ink = [0, 1, 2].map((k) =>
        rs.reduce(
          (a, r) => a + (((bg >> (16 - 8 * k)) & 255) - ((r.color >> (16 - 8 * k)) & 255)),
          0,
        ),
      );
      const weight = rs.map((r) => Math.max(1, chDiff(r.color, bg)));
      const pos =
        rs.reduce((a, r, i) => a + r.p * weight[i], 0) / weight.reduce((a, b) => a + b, 0);
      const peak = rs.reduce((a, r) => (chDiff(r.color, bg) > chDiff(a.color, bg) ? r : a), rs[0]);
      return {
        dir,
        pos: round2(pos),
        rel: 0,
        role: 'other',
        rows: rs.map((r) => ({ p: r.p, color: toHex(r.color), count: r.count })),
        peak: toHex(peak.color),
        ink,
      };
    });
  };

  const lines = [...group(scan('h'), 'h'), ...group(scan('v'), 'v')];

  // 役割
  const vert = entry.Horizontal !== true;
  const va = entry.ValueAxis;
  const tol = Math.max(2, Math.round(1.5 * scale));
  const lo = vert ? y0 : x0;
  const span = vert ? y1 - y0 : x1 - x0;
  const hasScale =
    va && typeof va.Min === 'number' && typeof va.Max === 'number' && va.Max > va.Min;
  const pxOf = (v) => {
    const t = (v - va.Min) / (va.Max - va.Min);
    return vert ? y0 + (1 - t) * span : x0 + t * span;
  };
  const ticks = (unit) => {
    if (!hasScale || !(typeof unit === 'number' && unit > 0)) return [];
    const n = Math.min(500, Math.round((va.Max - va.Min) / unit));
    return Array.from({ length: n + 1 }, (_v, k) => va.Min + k * unit);
  };
  const major = ticks(va?.Major);
  const minor = ticks(va?.Minor);
  const crossPx = hasScale ? pxOf(Math.min(Math.max(0, va.Min), va.Max)) : null;
  for (const l of lines) {
    const valueDir = l.dir === (vert ? 'h' : 'v');
    l.rel = round2(
      (l.pos - (valueDir ? lo : vert ? x0 : y0)) /
        Math.max(1, valueDir ? span : vert ? x1 - x0 : y1 - y0),
    );
    if (valueDir) {
      if (hasScale && Math.abs(l.pos - crossPx) <= tol) l.role = 'catAxisLine';
      else if (major.some((v) => Math.abs(l.pos - pxOf(v)) <= tol)) l.role = 'majorGrid';
      else if (minor.some((v) => Math.abs(l.pos - pxOf(v)) <= tol)) l.role = 'minorGrid';
    } else {
      // 値の軸の線：縦棒・折れ線は左端、横棒は下端
      const edge = vert ? x0 : y1;
      l.role = Math.abs(l.pos - edge) <= tol ? 'valAxisLine' : 'catGrid';
    }
  }
  res.lines = lines;
  // 役割ごとの代表（最初の線）：色は最も背景から離れた行の色（peak）・ink は線の全体
  for (const role of ['majorGrid', 'minorGrid', 'catAxisLine', 'valAxisLine', 'catGrid']) {
    const ls = lines.filter((l) => l.role === role);
    if (ls.length > 0)
      res.summary[role] = {
        count: ls.length,
        peak: ls[0].peak,
        ink: ls[0].ink,
        rows: ls[0].rows.length,
      };
  }
  // ISSUE-110：項目軸の文字の塊（LabelBand＝COM の CategoryAxis の位置 pt）と、数値軸の目盛の印（Ticks）
  const lb = entry.LabelBand;
  if (lb && [lb.Left, lb.Top, lb.Width, lb.Height].every((v) => typeof v === 'number')) {
    const g = Math.max(2, Math.ceil(1.8 * scale));
    const region = vert
      ? {
          x0: 0,
          y0: y1 + Math.ceil(2 * scale),
          x1: W - 1,
          y1: Math.round((lb.Top + lb.Height + COM_ORIGIN_PT) * scale) + 3,
        }
      : {
          x0: Math.round((lb.Left + COM_ORIGIN_PT) * scale) - 3,
          y0: y0 - 2,
          x1: x0 - Math.ceil(2 * scale),
          y1: y1 - 1,
        };
    res.labels = scanLabels(img, bg, region, g);
  }
  if (entry.Ticks === true && hasScale) {
    res.valTicks = scanTicks(img, bg, {
      vert,
      x0,
      y0,
      x1,
      y1,
      scale,
      majorPos: major.map(pxOf),
      minorPos: minor.map(pxOf),
      tol,
    });
  }
  return res;
}

/** 文字列の JSON を読む（BOM を除く） */
function readJson(file) {
  const text = fs.readFileSync(file, 'utf8');
  return JSON.parse(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text);
}

/** <out dir>/png の png-index*.json を全部読み、PNG を解いて ref-charts-png.json（opts.out で名前を変えられる）を書く。返すのは書いた内容 */
export function run(dir, opts = {}) {
  const pngDir = path.join(dir, 'png');
  const indexRe = new RegExp(opts.index ?? '^png-index.*\\.json$');
  const indexes = fs.readdirSync(pngDir).filter((f) => indexRe.test(f));
  if (indexes.length === 0)
    throw new Error(`${pngDir} に png-index*.json が無い（make-ref-charts2.ps1 の png 段を先に）`);
  const out = {
    note: 'Excel の Chart.Export の PNG から読んだ目盛線・軸線。ink＝線の行ごとの（背景−画素）の合計（チャンネル RGB）。照合は ink で行う（xlsx-chart-ref2.test.ts）',
    excelVersion: null,
    notes: [],
    charts: {},
  };
  for (const f of indexes) {
    const idx = readJson(path.join(pngDir, f));
    out.excelVersion ??= idx.excelVersion;
    const list = Array.isArray(idx.charts) ? idx.charts : idx.charts ? [idx.charts] : [];
    out.notes.push(...(Array.isArray(idx.notes) ? idx.notes : idx.notes ? [idx.notes] : []));
    for (const e of list) {
      const file = path.join(pngDir, e.File);
      if (!fs.existsSync(file)) {
        out.notes.push(`${e.Name}: PNG が無い（${e.File}）`);
        continue;
      }
      try {
        out.charts[e.Name] = {
          source: e.Source,
          chartStyle: e.ChartStyle ?? null,
          ...analyzeChartPng(decodePng(fs.readFileSync(file)), e),
        };
      } catch (err) {
        out.notes.push(`${e.Name}: ${err.message}`);
      }
    }
  }
  fs.writeFileSync(path.join(dir, opts.out ?? 'ref-charts-png.json'), JSON.stringify(out, null, 1));
  return out;
}

// 直接実行された時だけ走る（テストから import しても走らない）
if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  const dir = process.argv[2];
  const arg = (name) => {
    const i = process.argv.indexOf(name);
    return i >= 0 ? process.argv[i + 1] : undefined;
  };
  if (!dir || dir.startsWith('--')) {
    console.error(
      'usage: node make-ref-charts-png.mjs <out dir> [--out <file name>] [--index <regex>]',
    );
    process.exit(2);
  }
  const out = run(dir, { out: arg('--out'), index: arg('--index') });
  const names = Object.keys(out.charts);
  const noLines = names.filter((n) => out.charts[n].hasAxes && out.charts[n].lines.length === 0);
  console.log(`charts ${names.length}, notes ${out.notes.length}`);
  if (noLines.length)
    console.log(
      `警告：線が 1 本も取れなかったグラフ ${noLines.length} 件：${noLines.slice(0, 5).join(', ')}…`,
    );
}
