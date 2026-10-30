// 単独版が使う側の形で動くか確かめる（M6-3・`npm run check:standalone`）。先に `npm run build` で dist/ を作っておくこと。
// 1. xsheet.min.js：jsdom の窓で <script> と同じく読み込み（vm で窓の中で実行する。文字列からのコード生成は止める）、大域の XSheet.view が
//    fixtures の xlsx をタブ付きで描くこと（図形があり、形の表が中に入っていること）。グラフの xlsx のグラフを svg で描くこと（グラフの描画が中に入っていること）。
// 2. xsheet.mjs：素の Node（文字列からのコード生成を禁止）で import し、read で開けること。遅れて読む部品（xsheet-geom.mjs・xsheet-chart.mjs）が
//    同じフォルダから読めること。
// 3. xsheet.d.ts・xsheet.d.mts：大域の XSheet・`import type … from 'xsheet'`・`import … from './xsheet.mjs'` の型検査が通ること。
// 外へは通信しない。
import { execFileSync } from 'node:child_process';
import { copyFileSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import vm from 'node:vm';
import { JSDOM } from 'jsdom';
import { STANDALONE_FILES } from './parts.mjs';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const dist = path.join(repo, 'dist');
const fixture = path.join(repo, 'fixtures/excel/ref-drawings.xlsx');
const chartFixture = path.join(repo, 'fixtures/excel/ref-charts.xlsx');

/** 窓の中で次のフレームと少しの時間を待つ */
const settle = (window, ms = 50) =>
  new Promise((resolve) => window.requestAnimationFrame(() => window.setTimeout(resolve, ms)));

// 1. <script> で読む形
{
  const dom = new JSDOM('<!doctype html><html lang="en"><head></head><body></body></html>', {
    runScripts: 'outside-only',
    pretendToBeVisual: true,
  });
  const { window } = dom;
  // 文字列からのコード生成を窓の中で止める（Node の --disallow-code-generation-from-strings は vm の文脈に届かないため、
  // eval・Function・文字列の setTimeout／setInterval を、呼ばれたら投げるものに替える。N1）
  vm.runInContext(
    `(() => {
      const no = () => { throw new Error('code generation from strings'); };
      globalThis.eval = no;
      globalThis.Function = new Proxy(Function, { apply: no, construct: no });
      for (const name of ['setTimeout', 'setInterval']) {
        const orig = globalThis[name];
        globalThis[name] = (fn, ...args) => (typeof fn === 'string' ? no() : orig(fn, ...args));
      }
    })();`,
    dom.getInternalVMContext(),
  );
  vm.runInContext(
    readFileSync(path.join(dist, 'xsheet.min.js'), 'utf8'),
    dom.getInternalVMContext(),
  );
  const X = window.XSheet;
  if (typeof X?.view !== 'function' || typeof X.read !== 'function') {
    throw new Error('xsheet.min.js が大域の XSheet（view・read）を作らない');
  }
  // jsdom はレイアウトを持たない。図の層は見えている範囲だけを描くため、要素の大きさを固定する
  for (const [name, value] of [
    ['clientWidth', 1200],
    ['clientHeight', 800],
  ]) {
    Object.defineProperty(window.HTMLElement.prototype, name, {
      configurable: true,
      get: () => value,
    });
  }
  const container = window.document.createElement('div');
  window.document.body.append(container);
  const viewer = await X.view(container, new window.Uint8Array(readFileSync(fixture)));
  // 図の層は描き終わりを待つ（形の表を読み、図形を描く）
  await settle(window, 200);
  const tabs = container.querySelectorAll('.xs-sheet-tab').length;
  const cells = [...container.querySelectorAll('tbody td')].filter((td) => td.textContent).length;
  const shapes = container.querySelectorAll('svg.xs-shape-svg path, svg.xs-line').length;
  const style = window.document.head.querySelectorAll('style[data-xsheet-style]').length;
  // 形の表を中に持つ（読み込めないと図形を簡略化して通知を出す）
  const notice = container.querySelector('.xs-notice:not([hidden])')?.textContent ?? '';
  if (notice.includes(X.MESSAGES_EN.libLoadFailed))
    throw new Error('xsheet.min.js が形の表を読めない');
  if (tabs !== viewer.sheets.length || tabs === 0 || cells === 0 || shapes === 0 || style !== 1) {
    throw new Error(
      `XSheet.view の表示が足りない（タブ ${tabs}／シート ${viewer.sheets.length}・文字のあるセル ${cells}・図形の path ${shapes}・CSS ${style}）`,
    );
  }
  viewer.destroy();
  if (container.children.length !== 0) throw new Error('destroy の後に表示が残る');
  // グラフ：見えている範囲のグラフを svg で描く（グラフの描画を中に持つ）
  const chartViewer = await X.view(container, new window.Uint8Array(readFileSync(chartFixture)));
  // 1 枚目（Data）は値だけ。2 枚目からグラフがある
  await chartViewer.showSheet(2);
  await settle(window, 300);
  const charts = container.querySelectorAll('svg.xs-chart').length;
  chartViewer.destroy();
  if (charts === 0) throw new Error('xsheet.min.js がグラフを svg で描かない');
  console.log(
    `xsheet.min.js：OK（jsdom で XSheet.view・タブ ${tabs}・図形の path ${shapes}・グラフ ${charts}）`,
  );
  window.close();
}

// 2. バンドラ無しの import の形（素の Node）
{
  const script = [
    `import { readFileSync } from 'node:fs';`,
    `const x = await import(${JSON.stringify(pathToFileURL(path.join(dist, 'xsheet.mjs')).href)});`,
    `const geom = await import(${JSON.stringify(pathToFileURL(path.join(dist, 'xsheet-geom.mjs')).href)});`,
    `const chart = await import(${JSON.stringify(pathToFileURL(path.join(dist, 'xsheet-chart.mjs')).href)});`,
    `const book = await x.read(readFileSync(${JSON.stringify(fixture)}));`,
    `const sheet = await book.openSheet(book.sheets[0].index);`,
    `const out = { sheets: book.sheets.length, rows: sheet.rowCount, presets: geom.presetNames().length, chart: typeof chart.layoutChart, view: typeof x.view };`,
    `book.close();`,
    `process.stdout.write(JSON.stringify(out));`,
  ].join('\n');
  const out = JSON.parse(
    execFileSync(
      process.execPath,
      ['--disallow-code-generation-from-strings', '--input-type=module', '-e', script],
      { cwd: repo, encoding: 'utf8' },
    ),
  );
  if (
    out.sheets === 0 ||
    out.rows === 0 ||
    out.presets < 100 ||
    out.chart !== 'function' ||
    out.view !== 'function'
  ) {
    throw new Error(`xsheet.mjs の読み取りが足りない（${JSON.stringify(out)}）`);
  }
  console.log(`xsheet.mjs：OK（素の Node で read・シート ${out.sheets}・形 ${out.presets}）`);
}

// 3. 型
{
  const dir = mkdtempSync(path.join(tmpdir(), 'xsheet-standalone-'));
  try {
    for (const f of STANDALONE_FILES.filter((f) => f.endsWith('.d.ts') || f.endsWith('.d.mts'))) {
      copyFileSync(path.join(dist, f), path.join(dir, f));
    }
    writeFileSync(
      path.join(dir, 'tsconfig.json'),
      JSON.stringify({
        compilerOptions: {
          strict: true,
          target: 'ES2022',
          module: 'ESNext',
          moduleResolution: 'bundler',
          lib: ['ES2022', 'DOM', 'DOM.Iterable'],
          types: [],
          noEmit: true,
        },
        files: ['xsheet.d.ts', 'global.ts', 'module.ts', 'esm.ts'],
      }),
    );
    // <script> の使い方（大域の XSheet）
    writeFileSync(
      path.join(dir, 'global.ts'),
      `const el = document.getElementById('viewer')!;\n` +
        `XSheet.view(el, new Blob(), { lang: 'ko', theme: 'auto' }).then((v) => v.showSheet(v.sheets[0]?.name ?? 1));\n` +
        `XSheet.read(fetch('a.xlsx')).then((b) => b.close());\n`,
    );
    writeFileSync(
      path.join(dir, 'module.ts'),
      `import type { Viewer, ViewOptions, XlsxBook } from 'xsheet';\n` +
        `export const o: ViewOptions = { height: 300, css: 'none' };\n` +
        `export async function f(v: Viewer, b: XlsxBook): Promise<void> { await v.showSheet(1); b.close(); }\n`,
    );
    writeFileSync(
      path.join(dir, 'esm.ts'),
      `import { view, MESSAGES_KO } from './xsheet.mjs';\n` +
        `export const p = view('#v', new ArrayBuffer(0), { messages: { noData: MESSAGES_KO.noData } });\n`,
    );
    execFileSync(
      process.execPath,
      [path.join(repo, 'node_modules/typescript/bin/tsc'), '-p', 'tsconfig.json'],
      { cwd: dir, stdio: 'inherit' },
    );
    console.log('xsheet.d.ts・xsheet.d.mts：OK（大域・import type・./xsheet.mjs の型検査）');
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}
