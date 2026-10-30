// 同期の道具（tools/sync/sync.mjs・M6-1）の試験。
// 本物の InterHub は使わない（CI には無い）。一時フォルダに小さな git の模型を作り、INTERHUB_DIR の代わりに渡す。
// 道具が InterHub に書き込まないこと（HEAD・作業ツリー・.git の中身が同じ）を、実行の前後のファイルの一覧で確かめる。
import { execFileSync, spawnSync } from 'node:child_process';
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, relative } from 'node:path';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import {
  assertReadOnlyGitArgs,
  collectStatus,
  compileConfig,
  createGitRunner,
  formatStatus,
  loadConfig,
  main,
  mapPath,
  parseRange,
  READ_ONLY_GIT,
  REPO,
  resolveImport,
  rewritePatch,
  runPatch,
  transformLine,
  type SyncStatus,
} from '../tools/sync/sync.mjs';

const cfg = loadConfig();
const HAS_GIT = spawnSync('git', ['--version']).status === 0;

describe('git は読むだけ', () => {
  it('読むだけのサブコマンドの一覧', () => {
    expect([...READ_ONLY_GIT].sort()).toEqual(
      ['cat-file', 'diff', 'log', 'ls-tree', 'rev-parse', 'show'].sort(),
    );
  });

  it.each([
    ['checkout', 'main'],
    ['fetch'],
    ['pull'],
    ['push'],
    ['status'],
    ['commit', '-m', 'x'],
    ['reset', '--hard'],
    ['apply', 'x.patch'],
    ['gc'],
    ['config', 'user.name', 'x'],
    ['-c', 'core.pager=x', 'log'],
  ])('一覧に無いサブコマンドは呼ぶ前に断る：git %s', (...args) => {
    const exec = vi.fn();
    const git = createGitRunner('/nowhere', { exec });
    expect(() => git(args)).toThrow(/使わない/);
    expect(exec).not.toHaveBeenCalled();
  });

  it('書き出し・外のプログラムを動かすオプションは断る', () => {
    for (const bad of ['--output=x.txt', '--output', '--ext-diff', '--textconv', '--filters']) {
      expect(() => assertReadOnlyGitArgs(['diff', bad])).toThrow(/使わない/);
    }
    expect(() => assertReadOnlyGitArgs([])).toThrow();
    expect(() => assertReadOnlyGitArgs('log')).toThrow();
  });

  it('許す呼び出しは -C で InterHub を指し、引数を配列のまま渡す', () => {
    const exec = vi.fn(() => Buffer.from(''));
    const git = createGitRunner('/ih', { exec });
    git(['log', '--format=%H', 'a..b']);
    expect(exec).toHaveBeenCalledTimes(1);
    const [file, args, opts] = exec.mock.calls[0] as unknown as [
      string,
      string[],
      { env: Record<string, string> },
    ];
    expect(file).toBe('git');
    expect(args.slice(0, 2)).toEqual(['-C', '/ih']);
    expect(args.slice(-3)).toEqual(['log', '--format=%H', 'a..b']);
    expect(opts.env.GIT_OPTIONAL_LOCKS).toBe('0');
  });

  it('範囲の指定にオプションや空白は通さない', () => {
    expect(parseRange('943e4f0')).toEqual(['943e4f0', 'HEAD']);
    expect(parseRange('943e4f0..')).toEqual(['943e4f0', 'HEAD']);
    expect(parseRange('943e4f0..main~2')).toEqual(['943e4f0', 'main~2']);
    for (const bad of ['--all', 'a..--output=x', 'a b', 'a...b', '']) {
      expect(() => parseRange(bad)).toThrow();
    }
  });
});

describe('対応表（interhub.json）', () => {
  const at = (p: string) => mapPath(cfg, `src_V1/${p}`);

  it('取り込み済みのコミットと InterHub の場所', () => {
    expect(cfg.baseCommit).toBe('943e4f0');
    expect(cfg.envVar).toBe('INTERHUB_DIR');
    expect(cfg.root).toBe('src_V1/');
  });

  it('写した実績の組は files から引く（03 の記録と同じ）', () => {
    expect(at('src/lib/xlsx/sheet.ts')).toEqual({ kind: 'file', dest: 'src/core/sheet.ts' });
    expect(at('src/lib/xlsx/make-xlsx.ts').dest).toBe('src/testing/make-xlsx.ts');
    expect(at('src/lib/xlsx/linear-time.test.ts').dest).toBe('tests/core/linear-time.perf.test.ts');
    expect(at('src/lib/viewers/xlsx-image-layer.ts').dest).toBe('src/render/image-layer.ts');
    expect(at('src/lib/viewers/xlsx-geom.ts').dest).toBe('src/geom/evaluate.ts');
    expect(at('src/lib/viewers/xlsx-geom.test.ts').dest).toBe('tests/geom/evaluate.test.ts');
  });

  it('files の行き先は XSheet に実在する', () => {
    const missing = [...cfg.files.values()].filter((p) => !existsSync(join(REPO, p)));
    expect(missing).toEqual([]);
  });

  it('まだ写していないファイルは patterns から行き先を提案する', () => {
    expect(at('src/lib/xlsx/chart.ts')).toMatchObject({ kind: 'rule', dest: 'src/core/chart.ts' });
    expect(at('src/lib/xlsx/chart.test.ts').dest).toBe('tests/core/chart.test.ts');
    expect(at('src/lib/viewers/xlsx-chart-layout.ts').dest).toBe('src/render/chart-layout.ts');
    expect(at('src/lib/viewers/xlsx-chart.perf.test.ts').dest).toBe(
      'tests/render/chart.perf.test.ts',
    );
    expect(at('src/lib/viewers/xlsx-chart-fixtures.ts').dest).toBe(
      'tests/render/chart-fixtures.ts',
    );
    expect(at('src/lib/xlsx/__fixtures__/excel/ref-charts.xlsx').dest).toBe(
      'fixtures/excel/ref-charts.xlsx',
    );
  });

  it('手作業・除外・対応なし', () => {
    expect(at('src/lib/viewers/spreadsheet.ts').kind).toBe('manual');
    expect(at('src/lib/viewers/spreadsheet.image.render.test.ts').kind).toBe('manual');
    expect(at('src/styles/global.css').kind).toBe('manual');
    expect(at('src/lib/viewers/pdf.ts').kind).toBe('excluded');
    expect(at('src/lib/viewers/table-fit.ts').kind).toBe('excluded');
    expect(at('src/lib/viewers/newview.ts').kind).toBe('unmapped');
    expect(mapPath(cfg, 'README.md').kind).toBe('unmapped');
  });

  it('壊れた設定は読み込みで断る', () => {
    const raw = JSON.parse(
      readFileSync(new URL('../tools/sync/interhub.json', import.meta.url), 'utf8'),
    );
    expect(() =>
      compileConfig({ ...raw, interhub: { ...raw.interhub, baseCommit: 'HEAD' } }),
    ).toThrow(/baseCommit/);
    expect(() =>
      compileConfig({ ...raw, patterns: [{ from: 'src/lib/xlsx/*.ts', to: 'src/core/x.ts' }] }),
    ).toThrow(/\*/);
  });
});

describe('変換の規則', () => {
  const render = { from: 'src_V1/src/lib/viewers/xlsx-grid.ts', dest: 'src/render/grid.ts' };

  it('接頭辞 ih- を xs- 系へ（03 §3 の表のとおり）', () => {
    const t = (s: string) => transformLine(cfg, render, s);
    expect(t("el.className = 'ih-xlsx-figure ih-xlsx-line';")).toBe(
      "el.className = 'xs-figure xs-line';",
    );
    expect(t('svg.ih-xlsx-backdrop')).toBe('svg.xs-backdrop');
    expect(t('`ih-xf-${n}`')).toBe('`xs-xf-${n}`');
    expect(t("root.setAttribute('data-ih-xlsx-fmt', '')")).toBe(
      "root.setAttribute('data-xs-fmt', '')",
    );
    expect(t('var(--ih-xlsx-cell-bg) var(--ih-fg)')).toBe('var(--xs-cell-bg) var(--xs-fg)');
    expect(t('.ih-spreadsheet td')).toBe('.xs-sheet td');
    expect(t("'ih-dv-rownum' 'ih-sheet-tabs'")).toBe("'xs-rownum' 'xs-sheet-tabs'");
    // 語の途中・ih-png のような別の語は変えない
    expect(t("path.join(outDir, 'ih-png')")).toBe("path.join(outDir, 'ih-png')");
    expect(t('which-ih-xlsx-x')).toBe('which-ih-xlsx-x');
  });

  it('出典のコメントは InterHub の文書を指す', () => {
    expect(transformLine(cfg, render, '// 正: docs/plan/adr/0031-xlsx-charts.md')).toBe(
      '// 正: InterHub docs/plan/adr/0031-xlsx-charts.md',
    );
  });

  it('相対の import は対応表で引いて、行き先どうしの相対パスに書き直す', () => {
    const w: string[] = [];
    const chartSvg = {
      from: 'src_V1/src/lib/viewers/xlsx-chart-svg.ts',
      dest: 'src/render/chart-svg.ts',
    };
    expect(transformLine(cfg, chartSvg, "import { dashArray } from './xlsx-shapes';", w)).toBe(
      "import { dashArray } from './shapes';",
    );
    expect(transformLine(cfg, chartSvg, "import type { ChartResult } from '../xlsx';", w)).toBe(
      "import type { ChartResult } from '../core/index';",
    );
    expect(transformLine(cfg, chartSvg, "import type { X } from '../xlsx/types';", w)).toBe(
      "import type { X } from '../core/index';",
    );
    expect(w).toEqual([]);
    // render は core の中の作りに直接つながない：internal へ寄せて、確かめることに足す
    expect(
      transformLine(cfg, chartSvg, "import { applyTransforms } from '../xlsx/color';", w),
    ).toBe("import { applyTransforms } from '../core/internal';");
    expect(w).toHaveLength(1);

    const coreTest = {
      from: 'src_V1/src/lib/xlsx/chart.test.ts',
      dest: 'tests/core/chart.test.ts',
    };
    expect(transformLine(cfg, coreTest, "import { parseChart } from './chart';")).toBe(
      "import { parseChart } from '../../src/core/chart';",
    );
    expect(transformLine(cfg, coreTest, "const m = await import('./make-xlsx');")).toBe(
      "const m = await import('../../src/testing/make-xlsx');",
    );
    expect(
      transformLine(
        cfg,
        {
          from: 'src_V1/src/lib/viewers/xlsx-chart-ref.test.ts',
          dest: 'tests/render/chart-ref.test.ts',
        },
        'new URL(`../xlsx/__fixtures__/excel/${n}`, import.meta.url)',
      ),
    ).toBe('new URL(`../../fixtures/excel/${n}`, import.meta.url)');
    expect(
      resolveImport(
        cfg,
        'src_V1/src/lib/xlsx/__fixtures__/excel/make-ref-charts-default.ts',
        'fixtures/excel/make-ref-charts-default.ts',
        '../../make-xlsx',
      ),
    ).toBe('../../src/testing/make-xlsx');
  });

  it('行き先の無い import は変えずに、確かめることに足す', () => {
    const w: string[] = [];
    const line = "import { t } from '../i18n';";
    expect(transformLine(cfg, render, line, w)).toBe(line);
    expect(w.join('\n')).toMatch(/\.\.\/i18n/);
  });
});

describe('パッチの書き直し（文字列）', () => {
  it('見出しのパスを行き先に替え、hunk の行に変換をかけ、行き先の無いファイルは manual に分ける', () => {
    const diff = [
      'diff --git a/src_V1/src/lib/viewers/xlsx-grid.ts b/src_V1/src/lib/viewers/xlsx-grid.ts',
      'index 1111111..2222222 100644',
      '--- a/src_V1/src/lib/viewers/xlsx-grid.ts',
      '+++ b/src_V1/src/lib/viewers/xlsx-grid.ts',
      '@@ -1,2 +1,3 @@',
      " import { a } from '../xlsx';",
      " export const cls = 'ih-xlsx-cell';",
      "+export const wrap = 'ih-xlsx-wrap ih-toggle';",
      'diff --git a/src_V1/src/lib/viewers/newview.ts b/src_V1/src/lib/viewers/newview.ts',
      'new file mode 100644',
      'index 0000000..3333333',
      '--- /dev/null',
      '+++ b/src_V1/src/lib/viewers/newview.ts',
      '@@ -0,0 +1 @@',
      "+export const v = 'ih-xlsx-x';",
      '',
    ].join('\n');
    const r = rewritePatch(cfg, diff);
    const text = r.patch.toString('utf8');
    expect(text).toContain('diff --git a/src/render/grid.ts b/src/render/grid.ts');
    expect(text).toContain('--- a/src/render/grid.ts\n+++ b/src/render/grid.ts');
    expect(text).toContain(" import { a } from '../core/index';");
    expect(text).toContain(" export const cls = 'xs-cell';");
    expect(text).toContain("+export const wrap = 'xs-wrap ih-toggle';");
    expect(text).not.toContain('newview');
    expect(r.files).toEqual([
      expect.objectContaining({ status: 'M', dest: 'src/render/grid.ts', kind: 'file' }),
    ]);
    expect(r.skipped).toEqual([
      expect.objectContaining({
        status: 'A',
        path: 'src_V1/src/lib/viewers/newview.ts',
        kind: 'unmapped',
      }),
    ]);
    expect(r.manual.toString('utf8')).toContain('+++ b/src_V1/src/lib/viewers/newview.ts');
    // 残った ih- は確かめることに出す
    expect(r.warnings.some((w) => w.startsWith('src/render/grid.ts') && /接頭辞/.test(w))).toBe(
      true,
    );
  });
});

/** フォルダの下の全ファイル（.git を含む）の大きさと更新時刻 */
function snapshot(dir: string): Record<string, string> {
  const out: Record<string, string> = {};
  const walk = (d: string) => {
    for (const e of readdirSync(d, { withFileTypes: true })) {
      const p = join(d, e.name);
      if (e.isDirectory()) walk(p);
      else {
        const s = statSync(p);
        out[relative(dir, p).replace(/\\/g, '/')] = `${s.size}:${s.mtimeMs}`;
      }
    }
  };
  walk(dir);
  return out;
}

describe.skipIf(!HAS_GIT)('InterHub の模型で status・patch（git が無い環境では飛ばす）', () => {
  let root = '';
  let ih = '';
  let base = '';
  const git = (dir: string, ...args: string[]) =>
    execFileSync(
      'git',
      [
        '-C',
        dir,
        '-c',
        'user.name=xsheet-test',
        '-c',
        'user.email=xsheet-test@example.invalid',
        '-c',
        'commit.gpgsign=false',
        '-c',
        'core.autocrlf=false',
        ...args,
      ],
      { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] },
    ).trim();
  const put = (dir: string, file: string, text: string) => {
    mkdirSync(dirname(join(dir, file)), { recursive: true });
    writeFileSync(join(dir, file), text);
  };
  const commit = (message: string, files: Record<string, string>) => {
    for (const [f, t] of Object.entries(files)) put(ih, f, t);
    git(ih, 'add', '-A');
    git(ih, 'commit', '-q', '-m', message);
  };
  const X = 'src_V1/src/lib/xlsx';
  const V = 'src_V1/src/lib/viewers';

  beforeAll(() => {
    root = mkdtempSync(join(tmpdir(), 'xsheet-sync-test-'));
    ih = join(root, 'interhub');
    mkdirSync(ih);
    git(ih, 'init', '-q');
    commit('base', {
      [`${X}/index.ts`]: 'export const a = 1;\n',
      [`${V}/xlsx-grid.ts`]: "import { a } from '../xlsx';\nexport const cls = 'ih-xlsx-cell';\n",
      [`${V}/pdf.ts`]: 'export const pdf = 1;\n',
      'README.md': 'readme\n',
    });
    base = git(ih, 'rev-parse', 'HEAD');
    commit('feat: グラフを読む', {
      [`${X}/chart.ts`]: "import type { ChartModel } from './types';\nexport const chart = 1;\n",
      [`${X}/index.ts`]: "export const a = 1;\nexport { chart } from './chart';\n",
    });
    commit('feat: グラフを描く', {
      [`${V}/xlsx-chart-svg.ts`]: [
        "import { cls } from './xlsx-grid';",
        "export const svgClass = 'ih-xlsx-chart';",
        "export const fill = 'var(--ih-xlsx-backdrop)';",
        '',
      ].join('\n'),
      [`${V}/xlsx-chart-svg.test.ts`]: [
        "import { svgClass } from './xlsx-chart-svg';",
        "import { openXlsx } from '../xlsx';",
        '',
      ].join('\n'),
      [`${V}/xlsx-grid.ts`]:
        "import { a } from '../xlsx';\nexport const cls = 'ih-xlsx-cell';\nexport const wrap = 'ih-xlsx-wrap';\n",
    });
    commit('docs: 関係ないファイルだけ', { 'README.md': 'readme 2\n' });
    commit('feat: 別の表示', {
      [`${V}/newview.ts`]: 'export const v = 1;\n',
      [`${V}/pdf.ts`]: 'export const pdf = 2;\n',
    });
    // 未コミットの変更（作業ツリー・index・追跡していないファイル）は読まない
    put(ih, `${X}/index.ts`, 'export const dirty = 1;\n');
    put(ih, `${X}/staged.ts`, 'export const staged = 1;\n');
    git(ih, 'add', `${X}/staged.ts`);
    put(ih, `${X}/untracked.ts`, 'export const untracked = 1;\n');
  }, 60_000);

  afterAll(() => {
    if (root) rmSync(root, { recursive: true, force: true });
  });

  it('status：取り込み済みより後で対応表のパスに触れたコミットと、ファイルごとの行き先', () => {
    const st: SyncStatus = collectStatus(createGitRunner(ih), cfg, { range: `${base}..HEAD` });
    expect(st.base).toBe(base);
    expect(st.allCount).toBe(4);
    expect(st.commits.map((c) => c.subject)).toEqual([
      'feat: グラフを読む',
      'feat: グラフを描く',
      'feat: 別の表示',
    ]);
    const files = (i: number) =>
      st.commits[i].files.map((f) => [
        f.status,
        f.path.slice(f.path.lastIndexOf('/') + 1),
        f.kind,
        f.dest,
      ]);
    expect(files(0)).toEqual([
      ['A', 'chart.ts', 'rule', 'src/core/chart.ts'],
      ['M', 'index.ts', 'file', 'src/core/index.ts'],
    ]);
    expect(files(1)).toEqual([
      ['A', 'xlsx-chart-svg.test.ts', 'rule', 'tests/render/chart-svg.test.ts'],
      ['A', 'xlsx-chart-svg.ts', 'rule', 'src/render/chart-svg.ts'],
      ['M', 'xlsx-grid.ts', 'file', 'src/render/grid.ts'],
    ]);
    expect(files(2)).toEqual([
      ['A', 'newview.ts', 'unmapped', undefined],
      ['M', 'pdf.ts', 'excluded', undefined],
    ]);
    const grid = st.commits[1].files.find((f) => f.path.endsWith('xlsx-grid.ts'))!;
    expect([grid.added, grid.deleted]).toEqual([1, 0]);
    // 未コミットの変更は出てこない
    const all = JSON.stringify(st);
    expect(all).not.toMatch(/dirty|staged|untracked/);
    const text = formatStatus(st);
    expect(text).toContain('src/lib/xlsx/chart.ts  +2 -0  → src/core/chart.ts（規則・新規）');
    expect(text).toContain('対応なし');
  });

  it('patch：パスを行き先に替え ih- を xs- にしたパッチを書き、XSheet の形のファイルに当たる', () => {
    const out = join(root, 'out');
    const r = runPatch(createGitRunner(ih), cfg, {
      range: `${base}..HEAD`,
      outDir: out,
      interhubDir: ih,
    });
    expect(r.file.startsWith(out)).toBe(true);
    const text = readFileSync(r.file, 'utf8');
    expect(text).toContain('diff --git a/src/core/chart.ts b/src/core/chart.ts');
    expect(text).toContain('+++ b/src/render/chart-svg.ts');
    expect(text).toContain("+import { cls } from './grid';");
    expect(text).toContain("+export const svgClass = 'xs-chart';");
    expect(text).toContain("+export const fill = 'var(--xs-backdrop)';");
    expect(text).toContain("+import { svgClass } from '../../src/render/chart-svg';");
    expect(text).toContain("+import { openXlsx } from '../../src/core/index';");
    expect(text).toContain("+export const wrap = 'xs-wrap';");
    expect(text).not.toMatch(/ih-xlsx|src_V1|dirty|staged|untracked/);
    expect(r.skipped.map((s) => [s.path.slice(s.path.lastIndexOf('/') + 1), s.kind])).toEqual([
      ['newview.ts', 'unmapped'],
      ['pdf.ts', 'excluded'],
    ]);
    expect(readFileSync(r.manualFile!, 'utf8')).toContain(
      '+++ b/src_V1/src/lib/viewers/newview.ts',
    );

    // XSheet の形（写した後の取り込み済みの状態）のリポジトリに当てて確かめる
    const xs = join(root, 'xsheet');
    mkdirSync(xs);
    git(xs, 'init', '-q');
    put(xs, 'src/core/index.ts', 'export const a = 1;\n');
    put(
      xs,
      'src/render/grid.ts',
      "import { a } from '../core/index';\nexport const cls = 'xs-cell';\n",
    );
    git(xs, 'add', '-A');
    git(xs, 'commit', '-q', '-m', 'base');
    git(xs, 'apply', r.file);
    expect(readFileSync(join(xs, 'src/render/grid.ts'), 'utf8')).toBe(
      "import { a } from '../core/index';\nexport const cls = 'xs-cell';\nexport const wrap = 'xs-wrap';\n",
    );
    expect(readFileSync(join(xs, 'src/core/index.ts'), 'utf8')).toBe(
      "export const a = 1;\nexport { chart } from './chart';\n",
    );
    expect(existsSync(join(xs, 'tests/render/chart-svg.test.ts'))).toBe(true);
  });

  it('CLI（main）：INTERHUB_DIR の模型で status・patch が動く', () => {
    const logs: string[] = [];
    const env = { INTERHUB_DIR: ih };
    expect(main(['status', `${base}..HEAD`, '--json'], { env, log: (t) => logs.push(t) })).toBe(0);
    const st = JSON.parse(logs[0]) as SyncStatus;
    expect(st.commits).toHaveLength(3);
    const out = join(root, 'cli-out');
    expect(main(['patch', `${base}..`, '--out-dir', out], { env, log: (t) => logs.push(t) })).toBe(
      0,
    );
    expect(logs[1]).toContain('git apply --3way');
    expect(readdirSync(out).some((f) => f.endsWith('.patch'))).toBe(true);
  });

  it('パッチを InterHub の下には書かない', () => {
    expect(() =>
      runPatch(createGitRunner(ih), cfg, {
        range: `${base}..HEAD`,
        outDir: join(ih, 'tmp'),
        interhubDir: ih,
      }),
    ).toThrow(/InterHub の下/);
    expect(existsSync(join(ih, 'tmp'))).toBe(false);
  });

  it('道具は InterHub の模型に書き込まない（HEAD・作業ツリーの状態・.git を含む全ファイルが同じ）', () => {
    const head = git(ih, 'rev-parse', 'HEAD');
    const porcelain = git(ih, 'status', '--porcelain');
    const before = snapshot(ih);
    const runner = createGitRunner(ih);
    collectStatus(runner, cfg, { range: `${base}..HEAD` });
    runPatch(runner, cfg, { range: `${base}..HEAD`, outDir: join(root, 'out2'), interhubDir: ih });
    main(['status', `${base}..HEAD`], { env: { INTERHUB_DIR: ih }, log: () => {} });
    expect(snapshot(ih)).toEqual(before);
    expect(git(ih, 'rev-parse', 'HEAD')).toBe(head);
    expect(git(ih, 'status', '--porcelain')).toBe(porcelain);
    expect(porcelain).toContain('staged.ts');
    expect(porcelain).toContain('untracked.ts');
  });
});
