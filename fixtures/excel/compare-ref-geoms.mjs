// ISSUE-93：Excel が描いた図形の画像（make-ref-geom-pngs.ps1 の出力）と、InterHub が描いた図形を Playwright で写した画像を、輪郭の距離で比べる。
//   U3＝preset の形（楕円の弧の形 cloud・pie・chord・arc・blockArc・moon・smileyFace など。arcTo の角度の換算）、U4＝反転した文字入りの図形の文字の位置・向き。
// 比べ方：枠線だけの図形は中心線（細線化）を ±8 画素だけ位置合わせして測る（鋭角の張り出し・線の太さの描き方の差を除く）。塗りのある既定の図形は輪郭で測る。
// 使い方：node compare-ref-geoms.mjs <dir> [--arc visual|param] [--only name1,name2] [--shots] [--chrome <chrome.exe>]
//   <dir>      make-ref-geom-pngs.ps1 の -Dir（geoms-png\ flips-png\ ref-flips.xlsx flips.com.json export.json が入っている）。結果もここに書く
//   --arc      arcTo の角度の換算：visual＝視角→媒介変数（既定・実装どおり）、param＝角度を媒介変数のまま使う（U3 の切り替え先。esbuild の変換で差し替えるだけで src は変えない）
//   --anchors  one（既定）＝twoCellAnchor を oneCellAnchor（from＋xfrm の ext）に直した写しを InterHub に読ませる。InterHub は行の高さを整数 px に丸め（14.4pt→20px）、twoCellAnchor の図が xfrm より最大 5px ほど大きくなるため、形の比較から配置の丸めを除く。keep＝原本のまま
//   --only     指定した図形名だけ比べる（調査用）。--shots は合格した図形の重ね合わせ画像も書く（既定は許容超えのみ）
// 出力（<dir>）：ref-geoms.compare.json・ref-flips.compare.json（--arc param は .arc-param.json）＝期待の要約。ih-png\（InterHub の写し）・diff\（重ね合わせ：Excel＝赤・InterHub＝青・両方＝黒）・harness\（バンドル）は置くだけで、リポジトリには入れない。
// 実行後：ref-geoms.compare.json と ref-flips.compare.json を fixtures/excel へ写す（ref-geoms.compare.test.ts が読む）。
// 必要なもの：リポジトリの node_modules の esbuild（vite に同梱）、Playwright（playwright-core。環境変数 PLAYWRIGHT_CORE に package.json のパス、無ければ D:\tools\node の @playwright/cli）、Chrome。
// 写しは deviceScaleFactor 1.25（Excel の画面 125%＝1pt が 1.667 画素と同じ）・ライトテーマ・グリッドを隠して白地。距離は 100% 換算（画素 ÷ 1.25）。
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import process from 'node:process';
import { Buffer } from 'node:buffer';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import {
  SCALE_125,
  TOLERANCE_PX,
  compareContours,
  compareSkeletons,
  thin,
  compareText,
  cropToInk,
  decodePng,
  encodePng,
  groupOfShape,
  inkMask,
  resizeMask,
  overlayImage,
  pictureRect,
  summarize,
} from './geom-compare.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
// …/src_V1/src/lib/xlsx/fixtures/excel → リポジトリのルート
const repo = path.resolve(here, '../../../../../..');
const rel = (n) => Math.round(n * 1000) / 1000;
/** pt → CSS px（96dpi） */
const PX_PER_PT = 4 / 3;
/** 写す範囲の余白（CSS px）。枠線のはみ出し・アンチエイリアスの分 */
const MARGIN = 8;
/** 文字の位置の許容（100% 換算の px。ISSUE-93 U4：枠線の外接矩形の左上からの文字の重心の差。反転しない図形でも Excel との差が 3〜4px 出る（インセット・フォントの組版）ため 5） */
const TEXT_POS_TOL_PX = 5;
/** 中心線の位置合わせの範囲（画素。100% 換算で 6.4px） */
const SEARCH_PX = 8;
/** Excel でも見えない図形（枠も塗りも文字も無い）。InterHub は要素を作らない（ref-geoms.xlsx の 225 図形のうち 2 つ） */
const HIDDEN_GEOMS = new Set(['emptyDiamond', 'lineNone']);
/** 楕円の弧を使う形（U3） */
const U3_PRSTS = ['cloud', 'pie', 'chord', 'arc', 'blockArc', 'moon', 'smileyFace'];

const args = process.argv.slice(2);
const flag = (n) => {
  const i = args.indexOf(n);
  return i < 0 ? undefined : args[i + 1];
};
/** 値を取るオプション */
const WITH_VALUE = new Set(['--arc', '--only', '--chrome', '--anchors']);
let dir;
for (let i = 0; i < args.length; i++) {
  if (WITH_VALUE.has(args[i])) i++;
  else if (!args[i].startsWith('--')) dir ??= args[i];
}
if (!dir) {
  console.error(
    '使い方: node compare-ref-geoms.mjs <dir> [--arc visual|param] [--only a,b] [--shots] [--chrome <chrome.exe>]',
  );
  process.exit(2);
}
const outDir = path.resolve(dir);
const arcMode = flag('--arc') ?? 'visual';
if (arcMode !== 'visual' && arcMode !== 'param') throw new Error('--arc は visual か param');
const only = flag('--only') ? new Set(flag('--only').split(',')) : null;
const allShots = args.includes('--shots');
const anchorMode = flag('--anchors') ?? 'one';
if (anchorMode !== 'one' && anchorMode !== 'keep') throw new Error('--anchors は one か keep');
const suffix = arcMode === 'param' ? '.arc-param' : '';

// ---------- 依存の解決 ----------

/** Playwright（playwright-core）を探す */
function loadPlaywright() {
  const cands = [
    process.env.PLAYWRIGHT_CORE,
    path.join(repo, 'package.json'),
    'D:/tools/node/node_modules/@playwright/cli/package.json',
  ];
  for (const c of cands) {
    if (!c) continue;
    try {
      return createRequire(c)('playwright-core');
    } catch {
      /* 次の候補へ */
    }
  }
  throw new Error(
    'playwright-core が見つからない（環境変数 PLAYWRIGHT_CORE に package.json のパスを指定）',
  );
}

/** Chrome の実行ファイル（--chrome・環境変数 CHROME_PATH・既定の場所の順） */
function chromePath() {
  const cands = [
    flag('--chrome'),
    process.env.CHROME_PATH,
    'C:/Program Files/Google/Chrome/Application/chrome.exe',
    'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
  ];
  return cands.find((c) => c && fs.existsSync(c));
}

// ---------- InterHub 側のハーネス（src は変えず、renderSpreadsheet を esbuild でまとめる） ----------

/** --arc param 用：arcTo の角度の換算（視角→媒介変数）を、角度をそのまま使うように差し替える esbuild のプラグイン */
const arcParamPlugin = {
  name: 'arc-param',
  setup(b) {
    b.onLoad({ filter: /xlsx-geom\.ts$/ }, async (a) => {
      const from = 'return Math.atan2(rx * Math.sin(a), ry * Math.cos(a)) + 2 * Math.PI * turns;';
      const text = await fs.promises.readFile(a.path, 'utf8');
      if (!text.includes(from))
        throw new Error(
          'xlsx-geom.ts の parametricAngle が変わっている（arc-param の差し替え先が見つからない）',
        );
      return { contents: text.replace(from, 'return a;'), loader: 'ts' };
    });
  },
};

/**
 * twoCellAnchor の図を oneCellAnchor（from＋xfrm の ext）に直した xlsx を返す（--anchors one）。図の大きさを Excel の xfrm のとおりにする。
 * xfrm を持たない（グループなど）・回転した図形・自由図形（custGeom）の anchor は触らない（回転後の外接矩形が from・to に入っているため）。
 */
async function withOneCellAnchors(file) {
  const JSZip = createRequire(path.join(repo, 'package.json'))('jszip');
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  let n = 0;
  for (const name of Object.keys(zip.files).filter((f) =>
    /^xl\/drawings\/drawing\d+\.xml$/.test(f),
  )) {
    const xml = await zip.file(name).async('string');
    const out = xml.replace(
      /<xdr:twoCellAnchor[^>]*>([\s\S]*?)<\/xdr:twoCellAnchor>/g,
      (all, inner) => {
        const from = /<xdr:from>[\s\S]*?<\/xdr:from>/.exec(inner);
        const ext = /<a:xfrm[^>]*>\s*<a:off[^>]*\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(inner);
        // 回転した図形の from・to は回転後の外接矩形（xfrm の ext は回転前の大きさ）なので、そのまま（twoCellAnchor）にする
        if (!from || !ext || /<a:xfrm[^>]*rot=/.test(inner) || inner.includes('<a:custGeom'))
          return all;
        n++;
        const rest = inner
          .replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, '')
          .replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, '');
        return `<xdr:oneCellAnchor>${from[0]}<xdr:ext cx="${ext[1]}" cy="${ext[2]}"/>${rest}</xdr:oneCellAnchor>`;
      },
    );
    zip.file(name, out);
  }
  console.log(`twoCellAnchor → oneCellAnchor: ${n} 個`);
  return Buffer.from(await zip.generateAsync({ type: 'nodebuffer' }));
}

/** ハーネスのページ（<dir>/harness/site）を作る */
async function buildHarness() {
  const esbuild = createRequire(path.join(repo, 'package.json'))('esbuild');
  const hdir = path.join(outDir, 'harness');
  const site = path.join(hdir, 'site');
  fs.rmSync(site, { recursive: true, force: true });
  fs.mkdirSync(site, { recursive: true });
  const R = repo.split(path.sep).join('/');
  const entry = path.join(hdir, 'entry.js');
  fs.writeFileSync(
    entry,
    [
      `import '${R}/src_V1/src/styles/global.css'`,
      `import { renderSpreadsheet } from '${R}/src_V1/src/lib/viewers/spreadsheet'`,
      `import { fitTableBodyToWindow } from '${R}/src_V1/src/lib/viewers/table-fit'`,
      'window.__ready = false',
      'window.__load = async (url, name) => {',
      '  const buf = await (await fetch(url)).arrayBuffer()',
      '  const view = await renderSpreadsheet(buf, name)',
      "  document.getElementById('app').replaceChildren(view)",
      "  const body = view.querySelector('.ih-xlsx-body')",
      '  if (body) fitTableBodyToWindow(body)',
      '}',
      'window.__ready = true',
      '',
    ].join('\n'),
  );
  await esbuild.build({
    entryPoints: [entry],
    bundle: true,
    format: 'esm',
    splitting: true,
    outdir: path.join(site, 'js'),
    loader: { '.css': 'css' },
    nodePaths: [path.join(repo, 'node_modules')],
    plugins: arcMode === 'param' ? [arcParamPlugin] : [],
    logLevel: 'warning',
  });
  fs.writeFileSync(
    path.join(site, 'index.html'),
    '<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="js/entry.css"></head><body><div id="app"></div><script type="module" src="js/entry.js"></script></body></html>',
  );
  return site;
}

const MIME = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' };
/** ハーネスのページと、比べる xlsx（/xlsx/<名前>）を配る簡単なサーバ */
function serve(site, xlsx) {
  return http.createServer((req, res) => {
    const url = decodeURIComponent(req.url.split('?')[0]);
    let file = null;
    if (url.startsWith('/xlsx/')) file = xlsx[url.slice(6)] ?? null;
    else file = path.join(site, url === '/' ? 'index.html' : url);
    if (!file) {
      res.writeHead(404);
      res.end();
      return;
    }
    fs.readFile(file, (e, d) => {
      if (e) {
        res.writeHead(404);
        res.end();
        return;
      }
      res.writeHead(200, {
        'content-type': MIME[path.extname(file)] ?? 'application/octet-stream',
      });
      res.end(d);
    });
  });
}

/** 写しの前に当てる CSS：セルのグリッドを隠し、シートの面を白にする（図形の輪郭だけを取る） */
const SHOT_CSS =
  '.ih-xlsx-grid{visibility:hidden!important}.ih-xlsx-body,.ih-xlsx-wrap,.ih-xlsx-layer{background:#fff!important}.ih-xlsx-notice,.ih-sheet-tabs{display:none!important}';

/**
 * xlsx（1 枚目のシート）を InterHub で描き、図形ごとに写す。
 * 図の要素（`.ih-xlsx-figure`）は図形の並び（xlsx の drawing の順）で DOM に並び、見えない図形（枠も塗りも文字も無い）には要素が無い。
 * InterHub の行の高さは整数 px に丸めるため、図の位置は Excel の COM の位置（pt）とずれる（行が下るほど増える）。位置では対応づけず、並びの順で対応づけて、
 * 大きさ（回転していない図形が COM の幅・高さより小さくならないか）で取り違えを検査する（`warn` に入れる）。
 * @param {{ name: string, hidden?: boolean, w: number, h: number, rot: number }[]} shapes xlsx の並びの図形（w・h は COM の pt）
 * @returns {Promise<{ shots: Map<string, { png: Buffer, fig: { w: number, h: number } }>, warn: string[] }>}
 */
async function shoot(browser, port, xlsxName, shapes) {
  const page = await browser.newPage({
    viewport: { width: 1500, height: 2700 },
    deviceScaleFactor: SCALE_125,
  });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.emulateMedia({ colorScheme: 'light' });
  await page.goto(`http://127.0.0.1:${port}/index.html`);
  await page.waitForFunction(() => window.__ready === true);
  await page.addStyleTag({ content: SHOT_CSS });
  await page.evaluate(([u, n]) => window.__load(u, n), [`/xlsx/${xlsxName}`, xlsxName]);
  await page.waitForSelector('.ih-xlsx-grid', { state: 'attached' });
  await page.waitForFunction(
    () => document.querySelectorAll('.ih-xlsx-layer .ih-xlsx-figure').length > 0,
  );
  await page.waitForTimeout(800);
  const figs = await page.evaluate(() =>
    [...document.querySelectorAll('.ih-xlsx-layer .ih-xlsx-figure')].map((el) => {
      const r = el.getBoundingClientRect();
      return { x: r.x, y: r.y, w: r.width, h: r.height };
    }),
  );
  const visible = shapes.filter((s) => !s.hidden);
  if (figs.length !== visible.length)
    throw new Error(
      `${xlsxName}: 図の要素 ${figs.length} 個が、見える図形 ${visible.length} 個と合わない（全部の図が DOM に出ていない？）`,
    );
  const shots = new Map();
  const warn = [];
  for (let i = 0; i < visible.length; i++) {
    const s = visible[i];
    const f = figs[i];
    // 回転のない図形は、要素の大きさが COM の大きさより小さくはならないはず（枠線・はみ出し・弧の枠の分は大きくなりうる）
    if (s.rot === 0 && (f.w < s.w * PX_PER_PT - 6 || f.h < s.h * PX_PER_PT - 6)) {
      warn.push(
        `${s.name}: 要素 ${Math.round(f.w)}x${Math.round(f.h)}px・COM ${rel(s.w * PX_PER_PT)}x${rel(s.h * PX_PER_PT)}px`,
      );
    }
    if (only && !only.has(s.name) && !s.name.startsWith('u4')) continue;
    const x = Math.max(0, Math.floor(f.x) - MARGIN);
    const y = Math.max(0, Math.floor(f.y) - MARGIN);
    const clip = {
      x,
      y,
      width: Math.min(1500 - x, Math.ceil(f.w) + 2 * MARGIN + 1),
      height: Math.min(2700 - y, Math.ceil(f.h) + 2 * MARGIN + 1),
    };
    // 枠が重なる隣の図形（吹き出し・弧など、外接矩形が図形の枠より大きい）を写さないよう、対象の要素だけを見せる
    await page.evaluate(
      (k) =>
        document
          .querySelectorAll('.ih-xlsx-layer .ih-xlsx-figure')
          .forEach((el, j) => (el.style.visibility = j === k ? 'visible' : 'hidden')),
      i,
    );
    shots.set(s.name, {
      png: Buffer.from(await page.screenshot({ clip, type: 'png' })),
      fig: { w: f.w, h: f.h },
    });
  }
  if (errors.length > 0) console.error(`ページのエラー ${errors.length} 件:`, errors.slice(0, 3));
  console.log(
    `${xlsxName}: 要素 ${figs.length} 個・写した ${shots.size} 個・大きさの警告 ${warn.length} 件`,
  );
  await page.close();
  return { shots, warn };
}

/** Excel の PNG を読み、図の矩形のマスクを返す。PNG が無い・図が貼られていなければ null */
function excelMask(file) {
  if (!fs.existsSync(file)) return null;
  const img = decodePng(fs.readFileSync(file));
  const rect = pictureRect(img);
  return rect === null ? null : { mask: inkMask(img, rect), rect };
}

/** 輪郭の比較の結果を、JSON に書く形（丸めた数値）にする */
function brief(r) {
  return {
    p95: rel(r.p95),
    max: rel(r.max),
    mean: rel(r.mean),
    dw: rel(r.dw),
    dh: rel(r.dh),
    nA: r.nA,
    nB: r.nB,
    ...(r.empty ? { empty: r.empty } : {}),
  };
}

// ---------- 主処理 ----------

const expected = JSON.parse(fs.readFileSync(path.join(here, 'ref-geoms.expected.json'), 'utf8'));
const exportInfo = fs.existsSync(path.join(outDir, 'export.json'))
  ? JSON.parse(fs.readFileSync(path.join(outDir, 'export.json'), 'utf8'))
  : { excelVersion: null, notes: [] };
const flipsCom = fs.existsSync(path.join(outDir, 'flips.com.json'))
  ? JSON.parse(fs.readFileSync(path.join(outDir, 'flips.com.json'), 'utf8'))
  : null;
fs.mkdirSync(path.join(outDir, 'ih-png'), { recursive: true });
fs.mkdirSync(path.join(outDir, 'diff'), { recursive: true });

const site = await buildHarness();
const xlsxFiles = { 'ref-geoms.xlsx': path.join(here, 'ref-geoms.xlsx') };
if (anchorMode === 'one') {
  xlsxFiles['ref-geoms.xlsx'] = path.join(outDir, 'harness', 'ref-geoms.oneCell.xlsx');
  fs.writeFileSync(
    xlsxFiles['ref-geoms.xlsx'],
    await withOneCellAnchors(path.join(here, 'ref-geoms.xlsx')),
  );
}
if (flipsCom) xlsxFiles['ref-flips.xlsx'] = path.join(outDir, 'ref-flips.xlsx');
if (anchorMode === 'one' && flipsCom) {
  xlsxFiles['ref-flips.xlsx'] = path.join(outDir, 'harness', 'ref-flips.oneCell.xlsx');
  fs.writeFileSync(
    xlsxFiles['ref-flips.xlsx'],
    await withOneCellAnchors(path.join(outDir, 'ref-flips.xlsx')),
  );
}
const server = serve(site, xlsxFiles);
await new Promise((r) => server.listen(0, r));
const port = server.address().port;
const exe = chromePath();
const browser = await loadPlaywright().chromium.launch(
  exe ? { executablePath: exe, headless: true } : { channel: 'chrome', headless: true },
);

try {
  // ----- U3：ref-geoms.xlsx の図形 -----
  // xlsx の並びのまま（見えない図形 emptyDiamond・lineNone も含む。要素が無いだけ）
  const refShapes = Object.entries(expected.shapes).map(([n, s]) => ({
    name: n,
    group: groupOfShape(n),
    hidden: HIDDEN_GEOMS.has(n),
    prst: s.xml.geom.type === 'preset' ? s.xml.geom.prst : 'custGeom',
    type: s.com.AutoShapeType,
    rot: s.com.Rotation,
    w: s.com.Width,
    h: s.com.Height,
  }));
  const { shots, warn } = await shoot(browser, port, 'ref-geoms.xlsx', refShapes);
  const entries = {};
  const notes = warn.map((w) => `大きさの警告（要素の取り違えの疑い）: ${w}`);
  for (const s of refShapes.filter((x) => x.group !== 'skip' && (!only || only.has(x.name)))) {
    const shot = shots.get(s.name);
    if (shot) fs.writeFileSync(path.join(outDir, 'ih-png', `${s.name}.png`), shot.png);
    const ex = excelMask(path.join(outDir, 'geoms-png', `${s.name}.png`));
    if (ex === null) {
      notes.push(`Excel の画像が無い・空: ${s.name}`);
      continue;
    }
    const iImg = shot ? decodePng(shot.png) : null;
    const eCrop = cropToInk(ex.mask);
    const iCrop = iImg ? cropToInk(inkMask(iImg)) : { w: 0, h: 0, data: new Uint8Array(0) };
    // 枠線だけの図形は中心線（細線化）で比べる（線の太さの描き方の差を除く）。塗りのある図形（既定の図形）は輪郭で比べる。
    // 寸法の差（配置の丸め）は dw・dh に残し、距離は寸法をそろえて測る（rawP95＝そろえない場合の p95）
    const skel = s.group !== 'default';
    const noInk = { w: 0, h: 0, data: new Uint8Array(0) };
    const iInk = iImg ? inkMask(iImg) : noInk;
    // p95＝中心線を ±SEARCH_PX 画素だけ位置合わせして測る（鋭角の頂点の張り出しで外接矩形がずれても全体が動いて見えないように）。topLeftP95＝外接矩形の左上そろえだけの参考値
    const r = skel
      ? compareSkeletons(ex.mask, iInk, SCALE_125, false, SEARCH_PX)
      : compareContours(eCrop, iCrop, SCALE_125);
    const fit = skel
      ? compareSkeletons(ex.mask, iInk, SCALE_125, false, 0)
      : compareContours(
          eCrop,
          iCrop.w > 0 && eCrop.w > 0 ? resizeMask(iCrop, eCrop.w, eCrop.h) : iCrop,
          SCALE_125,
        );
    const ok = r.p95 <= TOLERANCE_PX;
    const entry = {
      group: s.group,
      prst: s.prst,
      type: s.type,
      ...brief(r),
      ...(skel ? { shift: { x: rel(r.shift.x), y: rel(r.shift.y) } } : {}),
      topLeftP95: rel(fit.p95),
      ok,
    };
    if (!shot) entry.noFigure = true;
    if (s.rot === 0 && s.group === 'preset')
      entry.excelPxPerPt = rel((ex.rect.x1 - ex.rect.x0 + 1) / s.w);
    entries[s.name] = entry;
    if (!ok || allShots) {
      const o = skel
        ? overlayImage(cropToInk(thin(ex.mask)), cropToInk(thin(iInk)))
        : overlayImage(eCrop, iCrop);
      fs.writeFileSync(
        path.join(outDir, 'diff', `${s.name}.png`),
        encodePng(o.width, o.height, o.data),
      );
    }
  }
  const list = Object.entries(entries).map(([name, e]) => ({ name, group: e.group, p95: e.p95 }));
  const sum = summarize(list, TOLERANCE_PX);
  const scales = Object.values(entries)
    .map((e) => e.excelPxPerPt)
    .filter((v) => v !== undefined);
  const u3 = {};
  for (const p of U3_PRSTS) {
    u3[p] = Object.entries(entries)
      .filter(([, e]) => e.prst === p && e.group === 'preset')
      .map(([n, e]) => ({ name: n, p95: e.p95, ok: e.ok }));
  }
  const medianScale = scales.length
    ? [...scales].sort((a, b) => a - b)[Math.floor(scales.length / 2)]
    : null;
  if (medianScale !== null && (medianScale < 1.6 || medianScale > 1.8))
    notes.push(
      `Excel の画像の倍率が想定外（中央値 ${medianScale} 画素/pt。画面 125% なら約 1.667）。比較の scale が合わない`,
    );
  const geomsOut = {
    version: 1,
    issue: 'ISSUE-93',
    excelVersion: exportInfo.excelVersion,
    arcMode,
    anchors: anchorMode,
    thresholdPx: TOLERANCE_PX,
    scale: SCALE_125,
    summary: {
      ...sum,
      excelPxPerPt: scales.length
        ? {
            min: Math.min(...scales),
            median: [...scales].sort((a, b) => a - b)[Math.floor(scales.length / 2)],
            max: Math.max(...scales),
          }
        : null,
      u3,
    },
    notes,
    shapes: entries,
  };
  fs.writeFileSync(
    path.join(outDir, `ref-geoms.compare${suffix}.json`),
    JSON.stringify(geomsOut, null, 2) + '\n',
  );
  console.log(
    `\n[U3] arc=${arcMode}：${sum.ok}/${sum.total} が ${TOLERANCE_PX}px 以内（${rel(sum.okRatio * 100)}%）`,
  );
  for (const [g, v] of Object.entries(sum.byGroup)) console.log(`  ${g}: ${v.ok}/${v.total}`);
  for (const n of sum.ngNames.slice(0, 60))
    console.log(
      `  NG ${n} (${entries[n].prst}) p95=${entries[n].p95} max=${entries[n].max} dw=${entries[n].dw} dh=${entries[n].dh}`,
    );
  for (const [p, v] of Object.entries(u3))
    console.log(`  U3 ${p}: ${v.map((x) => `${x.name}=${x.p95}`).join(' ') || '（図形なし）'}`);
  if (scales.length)
    console.log(
      `  Excel の画素/pt：中央値 ${geomsOut.summary.excelPxPerPt.median}（最小 ${Math.min(...scales)}・最大 ${Math.max(...scales)}。期待 1.667＋枠線の分）`,
    );

  // ----- U4：反転した文字入りの図形 -----
  if (flipsCom) {
    const flipShapes = flipsCom.shapes.map((s) => ({
      ...s,
      hidden: false,
      w: s.Width,
      h: s.Height,
      rot: s.Rotation,
    }));
    const { shots: fshots, warn: fwarn } = await shoot(browser, port, 'ref-flips.xlsx', flipShapes);
    const flipNotes = fwarn.map((w) => `大きさの警告（要素の取り違えの疑い）: ${w}`);
    const flips = {};
    for (const s of flipShapes.filter((x) => !x.twin)) {
      const twin = s.name.replace(/^u4_/, 'u4o_');
      const eT = excelMask(path.join(outDir, 'flips-png', `${s.name}.png`));
      const eO = excelMask(path.join(outDir, 'flips-png', `${twin}.png`));
      const iT = fshots.get(s.name);
      const iO = fshots.get(twin);
      if (iT) fs.writeFileSync(path.join(outDir, 'ih-png', `${s.name}.png`), iT.png);
      if (iO) fs.writeFileSync(path.join(outDir, 'ih-png', `${twin}.png`), iO.png);
      if (!eT || !eO || !iT || !iO) {
        flipNotes.push(`画像が足りない: ${s.name}`);
        continue;
      }
      const iTm = inkMask(decodePng(iT.png));
      const iOm = inkMask(decodePng(iO.png));
      const outline = compareSkeletons(eO.mask, iOm, SCALE_125, false, SEARCH_PX);
      const text = compareText({
        eText: eT.mask,
        eOutline: eO.mask,
        iText: iTm,
        iOutline: iOm,
        scale: SCALE_125,
      });
      const textOk =
        text.orientation === 'upright' &&
        text.dCentroid !== undefined &&
        Math.hypot(text.dCentroid.x, text.dCentroid.y) <= TEXT_POS_TOL_PX;
      flips[s.name] = {
        kind: s.kind,
        variant: s.variant,
        outline: { ...brief(outline), ok: outline.p95 <= TOLERANCE_PX },
        text: {
          orientation: text.orientation,
          p95Upright: text.p95Upright === undefined ? null : rel(text.p95Upright),
          p95Rot180: text.p95Rot180 === undefined ? null : rel(text.p95Rot180),
          dCentroid: text.dCentroid ? { x: rel(text.dCentroid.x), y: rel(text.dCentroid.y) } : null,
          excel: text.excel ? { cx: rel(text.excel.cx), cy: rel(text.excel.cy) } : null,
          interhub: text.interhub ? { cx: rel(text.interhub.cx), cy: rel(text.interhub.cy) } : null,
          ...(text.reason ? { reason: text.reason } : {}),
        },
        textOk,
      };
      if (!textOk) {
        const o = overlayImage(eT.mask, iTm);
        fs.writeFileSync(
          path.join(outDir, 'diff', `${s.name}.png`),
          encodePng(o.width, o.height, o.data),
        );
      }
    }
    const flipOut = {
      version: 1,
      issue: 'ISSUE-93',
      excelVersion: flipsCom.excelVersion,
      thresholdPx: TOLERANCE_PX,
      textPosTolPx: TEXT_POS_TOL_PX,
      scale: SCALE_125,
      notes: flipNotes,
      summary: {
        total: Object.keys(flips).length,
        outlineOk: Object.values(flips).filter((f) => f.outline.ok).length,
        textOk: Object.values(flips).filter((f) => f.textOk).length,
        rot180: Object.entries(flips)
          .filter(([, f]) => f.text.orientation === 'rot180')
          .map(([n]) => n),
      },
      shapes: flips,
    };
    fs.writeFileSync(
      path.join(outDir, `ref-flips.compare${suffix}.json`),
      JSON.stringify(flipOut, null, 2) + '\n',
    );
    console.log(
      `\n[U4] ${flipOut.summary.textOk}/${flipOut.summary.total} で文字の向き・位置が一致（枠線 ${flipOut.summary.outlineOk}/${flipOut.summary.total}）。Excel が 180° 回した文字：${flipOut.summary.rot180.join(' ') || 'なし'}`,
    );
    for (const [n, f] of Object.entries(flips))
      console.log(
        `  ${n}: ${f.text.orientation} d=(${f.text.dCentroid?.x ?? '-'}, ${f.text.dCentroid?.y ?? '-'}) 枠 p95=${f.outline.p95}${f.textOk ? '' : '  NG'}`,
      );
  } else {
    console.log('\n[U4] flips.com.json が無いため文字の比較は省略');
  }
} finally {
  await browser.close();
  server.close();
}
