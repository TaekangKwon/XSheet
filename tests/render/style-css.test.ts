import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { DEFAULTS, css, rule } from './style-css.helpers';

/**
 * M3-1：`style.css` と render の DOM・文書の対応（docs/02_architecture.md §6）。
 * 規則の無いクラス・定義の無い変数・使わない変数があれば落とす。
 */

/** コメントを除いた文面 */
function stripComments(text: string): string {
  return text.replace(/\/\*[\s\S]*?\*\/|\/\/.*$/gm, '');
}

const renderDir = join(process.cwd(), 'src/render');
const renderSources = readdirSync(renderDir)
  .filter((name) => name.endsWith('.ts'))
  .map((name) => stripComments(readFileSync(join(renderDir, name), 'utf8')));
const cssBody = stripComments(css);

/** 既定値のブロックで宣言している変数 */
const defined = [...rule(DEFAULTS).matchAll(/(--xs-[a-z-]+):/g)].map((m) => m[1]!);

describe('style.css と render の DOM', () => {
  it('render が要素に付ける xs- のクラスには、すべて規則がある（段落の xs-para・書式の xs-xf-N を除く）', () => {
    const used = new Set<string>();
    for (const src of renderSources)
      for (const m of src.matchAll(/['"`](xs-[a-z0-9-]+)/g)) used.add(m[1]!);
    used.delete('xs-xf-'); // 書式の規則は表示の根の <style>（format-css.ts）
    used.delete('xs-para'); // 段落は JS が style に入れる（InterHub にも規則が無い）
    expect(used.size).toBeGreaterThanOrEqual(25); // 拾い方が壊れて空にならないこと
    const missing = [...used].filter((c) => !new RegExp(`\\.${c}(?![a-z0-9-])`).test(cssBody));
    expect(missing).toEqual([]);
  });

  it('使う変数（style.css と render の JS）はすべて既定値のブロックで宣言し、宣言した変数はすべて使う', () => {
    const usedInCss = new Set([...cssBody.matchAll(/var\((--xs-[a-z-]+)/g)].map((m) => m[1]!));
    const usedInJs = new Set(
      renderSources.flatMap((src) => [...src.matchAll(/var\((--xs-[a-z-]+)/g)].map((m) => m[1]!)),
    );
    expect(usedInJs).toEqual(new Set(['--xs-cell-bg', '--xs-fg']));
    const used = new Set([...usedInCss, ...usedInJs]);
    expect([...used].filter((v) => !defined.includes(v))).toEqual([]);
    expect(defined.filter((v) => !used.has(v))).toEqual([]);
  });

  it('既定値のブロックは 1 つだけで、詳細度 0（:where）。:root に値を書かない', () => {
    expect(css.split(`\n${DEFAULTS} {`)).toHaveLength(2);
    expect(cssBody).not.toMatch(/(^|\n):root\s*\{/);
  });

  it('InterHub の接頭辞（ih-）が残っていない（出典を書いた冒頭のコメントを除く）', () => {
    expect(cssBody).not.toMatch(/\bih-|--ih-/);
  });
});

describe('スクロール枠と固定の見出し', () => {
  it('表示の根は container の高さいっぱいの縦の flex で、本文枠（.xs-body）が残りを埋めてスクロールする', () => {
    const roots = rule('.xs-sheet,\n.xs-workbook');
    expect(roots).toContain('display: flex');
    expect(roots).toContain('flex-direction: column');
    expect(roots).toContain('height: 100%');
    const body = rule('.xs-body');
    expect(body).toContain('flex: 1 1 auto');
    expect(body).toContain('min-height: 0');
    expect(body).toContain('overflow: auto');
    expect(body).toContain('overflow-anchor: none');
    expect(rule('.xs-workbook > .xs-sheet')).toContain('flex: 1 1 0');
  });

  it('見出し行（z 2）・行番号（z 1）・左上の角（z 3）は sticky で、図の層（z 0）より上', () => {
    const head = rule('.xs-body thead th');
    expect(head).toContain('position: sticky');
    expect(head).toContain('top: 0');
    expect(head).toContain('z-index: 2');
    const rownum = rule('.xs-grid .xs-rownum');
    expect(rownum).toContain('position: sticky');
    expect(rownum).toContain('left: 0');
    expect(rule('.xs-body tbody .xs-rownum')).toContain('z-index: 1');
    expect(rule('.xs-body thead th.xs-rownum')).toContain('z-index: 3');
  });

  it('見出し行の高さは GRID_HEADER_HEIGHT（24px）と同じ', async () => {
    const { GRID_HEADER_HEIGHT } = await import('../../src/render/grid');
    expect(rule('.xs-grid thead th')).toContain(`height: ${GRID_HEADER_HEIGHT}px`);
  });

  it('通知の帯は hidden で消え、◀▶ のボタンも hidden で消える（display を指定した規則が hidden に勝たない）', () => {
    expect(rule('.xs-notice[hidden]')).toContain('display: none');
    expect(rule('.xs-sheet-nav[hidden]')).toContain('display: none');
  });

  it('タブの列は横にスクロールし、offsetLeft の基準になる（position: relative）', () => {
    const tabs = rule('.xs-sheet-tabs');
    expect(tabs).toContain('position: relative');
    expect(tabs).toContain('overflow-x: auto');
    expect(tabs).toContain('min-width: 0');
    expect(rule('.xs-sheet-tab')).toContain('flex-shrink: 0');
  });

  it('タブ列を下に置くとき（xs-tabs-bottom）は、表との間の余白を上へ回す', () => {
    expect(rule('.xs-sheet-strip')).toContain('margin-bottom: 16px');
    const bottom = rule('.xs-workbook.xs-tabs-bottom > .xs-sheet-strip');
    expect(bottom).toContain('margin-top: 16px');
    expect(bottom).toContain('margin-bottom: 0');
  });
});

describe('ダーク表示（M6-2・ADR-0003 §2）', () => {
  /** テーマの属性のブロックの変数（名前 → 値）。auto は prefers-color-scheme: dark の中 */
  function themeVars(theme: 'dark' | 'auto'): Record<string, string> {
    const m = new RegExp(`\\[data-xs-theme='${theme}'\\] \\{([^}]*)\\}`).exec(cssBody);
    expect(m, theme).not.toBeNull();
    return Object.fromEntries(
      [...m![1]!.matchAll(/(--xs-[a-z-]+):\s*([^;]+);/g)].map((v) => [v[1]!, v[2]!.trim()]),
    );
  }

  it('dark と auto（prefers-color-scheme: dark の中）は、既定値のブロックの変数すべてに同じ値を持つ', () => {
    const dark = themeVars('dark');
    expect(Object.keys(dark).sort()).toEqual([...defined].sort());
    expect(themeVars('auto')).toEqual(dark);
    expect(cssBody).toMatch(/@media \(prefers-color-scheme: dark\) \{\s*\[data-xs-theme='auto'\]/);
    expect(rule("[data-xs-theme='dark']")).toContain('color-scheme: dark');
  });

  it('dark の値は docs/02_architecture.md §6 の「ダークの例」の列と同じ', () => {
    const doc = readFileSync(join(process.cwd(), 'docs/02_architecture.md'), 'utf8').replace(
      /\r\n/g,
      '\n',
    );
    const section = doc.slice(doc.indexOf('\n## 6.'), doc.indexOf('\n## 7.'));
    const dark = themeVars('dark');
    const rows = section.split('\n').filter((line) => line.startsWith('| `--xs-'));
    expect(rows.length).toBeGreaterThan(0);
    for (const line of rows) {
      const cells = line.split('|').map((c) => c.trim());
      const names = [...cells[1]!.matchAll(/`(--xs-[a-z-]+)`/g)].map((v) => v[1]!);
      const values = [...cells[4]!.matchAll(/`([^`]+)`/g)].map((v) => v[1]!);
      expect(values, cells[1]).toHaveLength(names.length);
      names.forEach((name, i) => expect(dark[name], name).toBe(values[i]));
    }
  });

  it('view の表示の根（.xs-viewer）は container の高さいっぱい', () => {
    expect(rule('.xs-viewer')).toContain('height: 100%');
  });
});

describe('文書との対応', () => {
  it('docs/02_architecture.md §6 の変数の表と、既定値のブロックの変数が同じ', () => {
    const doc = readFileSync(join(process.cwd(), 'docs/02_architecture.md'), 'utf8').replace(
      /\r\n/g,
      '\n',
    );
    const section = doc.slice(doc.indexOf('\n## 6.'), doc.indexOf('\n## 7.'));
    const rows = section.split('\n').filter((line) => line.startsWith('| `--xs-'));
    const documented = rows.flatMap((line) =>
      [...line.split('|')[1]!.matchAll(/`(--xs-[a-z-]+)`/g)].map((m) => m[1]!),
    );
    expect([...documented].sort()).toEqual([...defined].sort());
    // 明るい既定値の列が style.css の値と同じ
    for (const line of rows) {
      const cells = line.split('|').map((c) => c.trim());
      const names = [...cells[1]!.matchAll(/`(--xs-[a-z-]+)`/g)].map((m) => m[1]!);
      const values = [...cells[3]!.matchAll(/`([^`]+)`/g)].map((m) => m[1]!);
      expect(values, cells[1]).toHaveLength(names.length);
      names.forEach((name, i) => expect(rule(DEFAULTS), name).toContain(`${name}: ${values[i]};`));
    }
  });
});
