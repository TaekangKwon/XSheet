import { execFileSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'vite';
import { describe, expect, it } from 'vitest';
import * as core from '../../src/core/index';
import * as all from '../../src/index';
import * as render from '../../src/render/index';

// M6-2：パッケージの入口 `xsheet`（src/index.ts）は、入口の層＋core＋render の公開物を全部出す（ADR-0003 §3）。
// render・入口の層は読み込んだ時点で DOM に触れない：入口を Vite で束ね、素の Node（jsdom なし・文字列からのコード生成を禁止）で import して read で開く。

const root = fileURLToPath(new URL('../../', import.meta.url));
// 出力は node_modules の下（束ねたものが jszip・ssf を node_modules から解決できる場所。Git の対象外）
const outDir = fileURLToPath(
  new URL('../../node_modules/.cache/xsheet-entry-import/', import.meta.url),
);
const fixture = fileURLToPath(new URL('../../fixtures/excel/ref-drawings.xlsx', import.meta.url));

describe('入口 xsheet', () => {
  it('core と render の公開物をすべて残し、view・read を足す（今の公開物は壊れない）', () => {
    const expected = [...Object.keys(core), ...Object.keys(render), 'read', 'view'].sort();
    expect(Object.keys(all).sort()).toEqual(expected);
    expect(all.openXlsx).toBe(core.openXlsx);
    expect(all.renderWorkbook).toBe(render.renderWorkbook);
    expect(Object.keys(render)).toContain('MESSAGES_KO');
  });

  it('束ねた入口を素の Node で import しても DOM に触れず、read で xlsx を開ける', async () => {
    await build({
      configFile: false,
      root,
      logLevel: 'silent',
      build: {
        outDir,
        emptyOutDir: true,
        minify: false,
        lib: { entry: { index: 'src/index.ts' }, formats: ['es'] },
        rolldownOptions: { external: ['jszip', 'ssf'] },
      },
    });
    const script = [
      `import { readFileSync } from 'node:fs';`,
      `const x = await import(${JSON.stringify(pathToFileURL(`${outDir}index.js`).href)});`,
      `const book = await x.read(readFileSync(${JSON.stringify(fixture)}));`,
      `const names = book.sheets.map((s) => s.name);`,
      `book.close();`,
      `process.stdout.write(JSON.stringify({ document: typeof document, window: typeof window, names, view: typeof x.view, keys: Object.keys(x).length }));`,
    ].join('\n');
    const stdout = execFileSync(
      process.execPath,
      ['--disallow-code-generation-from-strings', '--input-type=module', '-e', script],
      { cwd: root, encoding: 'utf8' },
    );
    const child = JSON.parse(stdout) as Record<string, unknown>;
    const expected = await core.openXlsx(
      new Uint8Array((await import('node:fs')).readFileSync(fixture)),
    );
    expect(child).toEqual({
      document: 'undefined',
      window: 'undefined',
      names: expected.sheets.map((s) => s.name),
      view: 'function',
      keys: Object.keys(all).length,
    });
    expected.close();
  }, 60_000);
});
