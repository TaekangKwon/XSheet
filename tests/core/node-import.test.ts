import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'vite';
import { describe, expect, it } from 'vitest';
import { openXlsx, type XlsxBook } from '../../src/core/index';

// M1-3：core が DOM・Node 固有の API なしで動くこと（01_requirements.md N8）を、vitest の変換を通さずに確かめる。
// core の入口を Vite で ES モジュールに束ね、素の Node（jsdom なし・文字列からのコード生成を禁止）で import して開く。
// 束ねるとき jszip・ssf は外に出し、使う側と同じく node_modules から読む

const root = fileURLToPath(new URL('../../', import.meta.url));
// 出力は node_modules の下（束ねたものが jszip・ssf を node_modules から解決できる場所。Git の対象外）
const outDir = fileURLToPath(
  new URL('../../node_modules/.cache/xsheet-node-import/', import.meta.url),
);
const fixture = fileURLToPath(new URL('../../fixtures/excel/ref-drawings.xlsx', import.meta.url));

/** ブックの中身の要約（シートの一覧・先頭の行の表示文字列・図の種類）。束ねたものと元のソースで同じになるはず */
async function summarize(book: XlsxBook) {
  const out = [];
  for (const info of book.sheets) {
    const sheet = await book.openSheet(info.index);
    const rows = sheet.getRows(1, Math.min(sheet.rowCount, 30));
    const drawings = await book.openDrawings(info.index);
    out.push({
      ...info,
      rowCount: sheet.rowCount,
      colCount: sheet.colCount,
      mdw: sheet.layout.mdw,
      texts: rows.map((r) => r.cells.map((c) => c?.text ?? null)),
      drawings: drawings.drawings.map((d) => d.kind),
      notices: drawings.notices,
    });
  }
  book.close();
  return out;
}

describe('core を素の Node で import する', () => {
  it('束ねた core が DOM なしで xlsx を開き、元のソースと同じ結果を返す', async () => {
    const result = await build({
      configFile: false,
      root,
      logLevel: 'silent',
      build: {
        outDir,
        emptyOutDir: true,
        minify: false,
        lib: {
          entry: { index: 'src/core/index.ts', internal: 'src/core/internal.ts' },
          formats: ['es'],
        },
        rolldownOptions: { external: ['jszip', 'ssf'] },
      },
    });
    const outputs = (Array.isArray(result) ? result : [result]).flatMap((r) =>
      'output' in r ? r.output : [],
    );
    const moduleIds = outputs.flatMap((o) => (o.type === 'chunk' ? o.moduleIds : []));
    // core は geom の defs.ts だけを引き、形の表（preset-geom.json・約 171 KB）は束ねない（N6）
    expect(moduleIds.some((id) => id.endsWith('src/geom/defs.ts'))).toBe(true);
    expect(moduleIds.some((id) => id.includes('preset-geom') || id.includes('evaluate'))).toBe(
      false,
    );

    const script = [
      `import { readFileSync } from 'node:fs';`,
      `const core = await import(${JSON.stringify(pathToFileURL(`${outDir}index.js`).href)});`,
      `const internal = await import(${JSON.stringify(pathToFileURL(`${outDir}internal.js`).href)});`,
      `const summarize = ${summarize.toString()};`,
      `const book = await core.openXlsx(readFileSync(${JSON.stringify(fixture)}));`,
      `const env = { document: typeof document, window: typeof window, internal: Object.keys(internal).sort() };`,
      `process.stdout.write(JSON.stringify({ env, sheets: await summarize(book) }));`,
    ].join('\n');
    const stdout = execFileSync(
      process.execPath,
      ['--disallow-code-generation-from-strings', '--input-type=module', '-e', script],
      { cwd: root, encoding: 'utf8' },
    );
    const child = JSON.parse(stdout) as { env: unknown; sheets: unknown };

    expect(child.env).toEqual({
      document: 'undefined',
      window: 'undefined',
      internal: [
        'DEFAULT_MDW',
        'DEFAULT_THEME',
        'applyFillMode',
        'applyTransforms',
        'emptyLayout',
        'formatNumberByCode',
      ],
    });
    const expected = await summarize(await openXlsx(readFileSync(fixture)));
    expect(expected.length).toBeGreaterThan(0);
    expect(expected.some((s) => s.drawings.length > 0)).toBe(true);
    expect(child.sheets).toEqual(JSON.parse(JSON.stringify(expected)));
  }, 60_000);
});
