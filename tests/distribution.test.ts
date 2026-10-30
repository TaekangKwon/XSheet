import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

/**
 * M4-2：配る物の約束（docs/02_architecture.md §8）。ビルドしなくても確かめられることだけを見る
 * （dist の中身・NOTICE の同梱・大きさは `npm run check:consumer`・`npm run size`）。
 */

function read(path: string): string {
  return readFileSync(join(process.cwd(), path), 'utf8').replace(/\r\n/g, '\n');
}

const pkg = JSON.parse(read('package.json')) as {
  files: string[];
  dependencies: Record<string, string>;
  exports: Record<string, string | Record<string, string>>;
};
const doc = read('docs/02_architecture.md');
const section = doc.slice(doc.indexOf('\n## 8.'));

/** §8 の表のうち、見出しが `heading` の表の行（セルの配列） */
function tableRows(heading: string): string[][] {
  const start = section.indexOf(`\n### ${heading}`);
  expect(start, heading).toBeGreaterThanOrEqual(0);
  const end = section.indexOf('\n### ', start + 1);
  return section
    .slice(start, end < 0 ? undefined : end)
    .split('\n')
    .filter((l) => l.startsWith('| `'))
    .map((l) => l.split('|').map((c) => c.trim()));
}

describe('依存のライセンス', () => {
  const lock = JSON.parse(read('package-lock.json')) as {
    packages: Record<string, { version?: string; license?: string; dev?: boolean }>;
  };
  const runtime = Object.entries(lock.packages)
    .filter(([key, v]) => key !== '' && v.dev !== true)
    .map(([key, v]) => ({
      name: key.replace(/^.*node_modules\//, ''),
      version: v.version!,
      license: v.license!,
    }))
    .sort((a, b) => a.name.localeCompare(b.name));

  it('実行時の依存は jszip と ssf だけ（CLAUDE.md：依存を増やさない）', () => {
    expect(Object.keys(pkg.dependencies).sort()).toEqual(['jszip', 'ssf']);
  });

  it('§8 の表は lockfile の実行時の依存（名前・版・ライセンス）と同じ', () => {
    const documented = tableRows('依存のライセンス')
      .map((cells) => ({
        name: cells[1]!.replace(/`/g, ''),
        version: cells[2]!,
        license: cells[3]!.replace(/（.*）$/, ''),
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
    expect(documented).toEqual(runtime);
  });

  it('ライセンスは使う側の配布を縛らないものだけ（GPL は MIT との選択のときだけ）', () => {
    const allowed = new Set([
      'MIT',
      'ISC',
      'Apache-2.0',
      '(MIT AND Zlib)',
      '(MIT OR GPL-3.0-or-later)',
    ]);
    expect(runtime.filter((p) => !allowed.has(p.license))).toEqual([]);
  });
});

describe('配る物の形', () => {
  it('files は dist と NOTICE だけ（src・テストを配らない）', () => {
    expect(pkg.files).toEqual(['dist', 'NOTICE']);
  });

  it('exports は dist の中を指し、JS の入口には型が付く', () => {
    for (const [sub, target] of Object.entries(pkg.exports)) {
      if (typeof target === 'string') {
        expect(target, sub).toMatch(/^\.\/dist\//);
      } else {
        expect(Object.keys(target), sub).toEqual(['types', 'import']);
        expect(target.import, sub).toMatch(/^\.\/dist\/.*\.js$/);
        expect(target.types, sub).toBe(target.import!.replace(/\.js$/, '.d.ts'));
      }
    }
  });

  it('ビルドの入口（vite.config.ts）は exports の JS と同じで、NOTICE と style.css を dist へ写す', () => {
    const config = read('vite.config.ts');
    const entries = [...config.matchAll(/'?([a-z]+(?:\/[a-z]+)?)'?: src\('[a-z/]+\.ts'\)/g)].map(
      (m) => m[1],
    );
    const fromExports = Object.values(pkg.exports)
      .filter((t): t is Record<string, string> => typeof t !== 'string')
      .map((t) => t.import!.replace(/^\.\/dist\/|\.js$/g, ''));
    expect(entries.sort()).toEqual(fromExports.sort());
    expect(config).toMatch(/'style\.css': src\('style\.css'\)/);
    expect(config).toMatch(/NOTICE: `\$\{root\}NOTICE`/);
    expect(pkg.exports['./style.css']).toBe('./dist/style.css');
  });

  it('大きさの上限（tools/size.mjs）は §8 の表と同じ', () => {
    const script = read('tools/size.mjs');
    const limits = Object.fromEntries(
      [...script.matchAll(/'([a-z]+\/index)': ([\d_]+),/g)].map((m) => [
        m[1],
        Number(m[2]!.replace(/_/g, '')),
      ]),
    );
    const documented = Object.fromEntries(
      tableRows('大きさ').map((cells) => [
        /`([a-z]+\/index)`/.exec(cells[1]!)![1],
        Number(cells[5]!.replace(/,/g, '')),
      ]),
    );
    expect(Object.keys(limits)).toHaveLength(3);
    expect(documented).toEqual(limits);
  });
});
