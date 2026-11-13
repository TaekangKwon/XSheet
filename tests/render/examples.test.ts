import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { DEFAULTS, rule } from './style-css.helpers';

/**
 * M3-2・M3-3：例のページ（examples/）の文面の確認。見た目は実機で目視する。
 * ダークの値が文書（docs/02_architecture.md §6）とずれないこと、CSP が N7 の要件どおりであることを確かめる。
 */

function read(path: string): string {
  return readFileSync(join(process.cwd(), path), 'utf8').replace(/\r\n/g, '\n');
}

const exampleCss = read('examples/example.css');
const indexHtml = read('examples/index.html');
const mainTs = read('examples/main.ts');

/** `selector {` から最初の `}` までの宣言（`--xs-*` だけ）を { 名前: 値 } で */
function variables(text: string, selector: string): Record<string, string> {
  const i = text.indexOf(`${selector} {`);
  expect(i, selector).toBeGreaterThanOrEqual(0);
  const body = text.slice(i, text.indexOf('}', i));
  return Object.fromEntries(
    [...body.matchAll(/(--xs-[a-z-]+):\s*([^;]+);/g)].map((m) => [m[1]!, m[2]!.trim()]),
  );
}

/** §6 の表の「ダークの例」の列を { 名前: 値 } で */
function documentedDark(): Record<string, string> {
  const doc = read('docs/02_architecture.md');
  const section = doc.slice(doc.indexOf('\n## 6.'), doc.indexOf('\n## 7.'));
  const result: Record<string, string> = {};
  for (const line of section.split('\n').filter((l) => l.startsWith('| `--xs-'))) {
    const cells = line.split('|').map((c) => c.trim());
    const names = [...cells[1]!.matchAll(/`(--xs-[a-z-]+)`/g)].map((m) => m[1]!);
    const values = [...cells[4]!.matchAll(/`([^`]+)`/g)].map((m) => m[1]!);
    expect(values, cells[1]).toHaveLength(names.length);
    names.forEach((name, k) => (result[name] = values[k]!));
  }
  return result;
}

describe('examples/example.css のダークの値', () => {
  const byAttr = variables(exampleCss, ":root[data-theme='dark']");
  const byMedia = variables(exampleCss, ":root:not([data-theme='light'])");

  it('テーマの属性と prefers-color-scheme の 2 つのブロックが同じ値', () => {
    expect(byMedia).toEqual(byAttr);
    expect(exampleCss).toContain('@media (prefers-color-scheme: dark)');
  });

  it('値は §6 の「ダークの例」の列と同じで、style.css の変数を全部上書きする', () => {
    expect(byAttr).toEqual(documentedDark());
    const defined = [...rule(DEFAULTS).matchAll(/(--xs-[a-z-]+):/g)].map((m) => m[1]!);
    expect(Object.keys(byAttr).sort()).toEqual(defined.sort());
  });
});

describe('examples/index.html', () => {
  const csp = /http-equiv="Content-Security-Policy"\s+content="([^"]+)"/.exec(indexHtml)?.[1] ?? '';
  const directives = new Map(
    csp.split(';').map((d) => {
      const [name, ...values] = d.trim().split(/\s+/);
      return [name!, values] as const;
    }),
  );

  it('CSP は N7 の要件だけを許す（script は同じ出所・画像は blob:・外の出所や eval を許さない）', () => {
    expect(directives.get('default-src')).toEqual(["'none'"]);
    expect(directives.get('script-src')).toEqual(["'self'"]);
    expect(directives.get('img-src')).toEqual(['blob:']);
    expect(csp).not.toMatch(/unsafe-eval|https?:|\*/);
  });

  it('インラインのスクリプトを持たない（スクリプトは外のファイルだけ）', () => {
    const scripts = [...indexHtml.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/g)];
    expect(scripts.length).toBeGreaterThan(0);
    for (const [, attrs, body] of scripts) {
      expect(attrs).toContain('src=');
      expect(body!.trim()).toBe('');
    }
  });

  it('style-src は nonce だけ（unsafe-inline・self を許さない。M3-3）', () => {
    expect(directives.get('style-src')).toEqual(["'nonce-xsheet-example'"]);
  });

  it('ライブラリの style.css と例の CSS を、CSP と同じ nonce を付けた <link> で読む', () => {
    const links = [...indexHtml.matchAll(/<link rel="stylesheet" ([^>]*)\/>/g)].map((m) => m[1]!);
    expect(links.map((a) => /href="([^"]+)"/.exec(a)?.[1])).toEqual([
      '/src/style.css',
      './example.css',
    ]);
    for (const attrs of links) {
      expect(attrs).toContain('nonce="xsheet-example"');
    }
  });
});

describe('examples/main.ts', () => {
  it('ライブラリは公開の入口（core・render の index）からだけ引く', () => {
    const imports = [...mainTs.matchAll(/from '([^']+)'/g)].map((m) => m[1]);
    // core は検索（findText・openXlsx。M7-10）
    expect(imports).toEqual(['../src/core/index', '../src/render/index']);
  });

  it('ページの nonce（<link> の nonce プロパティ）を renderWorkbook に渡す', () => {
    expect(mainTs).toMatch(/link\[rel="stylesheet"\]\[nonce\]'\)\?\.nonce/);
    expect(mainTs).toMatch(/renderWorkbook\([^)]*\bnonce,/);
  });

  it('サンプルは fixtures/excel にある自作のファイル', () => {
    const files = [...mainTs.matchAll(/file: '([^']+)'/g)].map((m) => m[1]!);
    expect(files.length).toBeGreaterThan(0);
    for (const f of files) {
      expect(existsSync(join(process.cwd(), 'fixtures/excel', f)), f).toBe(true);
    }
  });

  it('サンプルはページからの相対で読む（Pages の <サイト>/examples/ でも動く・ADR-0005）', () => {
    expect(mainTs).toContain('fetch(`../fixtures/excel/${file}`)');
    expect(mainTs).not.toMatch(/fetch\(`\//);
  });
});

describe('examples/standalone.html・esm.html（M6-4・単独版）', () => {
  const pages = {
    standalone: read('examples/standalone.html'),
    esm: read('examples/esm.html'),
  };

  it.each(Object.entries(pages))(
    '%s：CSP は通信・外の出所・eval・unsafe-inline を許さず、画像は blob: だけ',
    (_, html) => {
      const csp = /http-equiv="Content-Security-Policy"\s+content="([^"]+)"/.exec(html)?.[1] ?? '';
      expect(csp).toContain("default-src 'none'");
      expect(csp).toContain('img-src blob:');
      expect(csp).not.toMatch(/connect-src|unsafe-|https?:|\*/);
    },
  );

  it.each(Object.entries(pages))(
    '%s：インラインの script・style は CSP の nonce を持ち、同じ nonce を view に渡す',
    (_, html) => {
      const nonce = /'nonce-([\w-]+)'/.exec(html)?.[1];
      expect(nonce).toBeDefined();
      // 説明のコメントの中の <style> などは数えない
      const code = html.replace(/<!--[\s\S]*?-->/g, '');
      const inline = [...code.matchAll(/<(script|style)\b([^>]*)>([\s\S]*?)<\/\1>/g)].filter(
        ([, , attrs, body]) => !attrs!.includes('src=') && body!.trim() !== '',
      );
      expect(inline.length).toBeGreaterThan(0);
      for (const [, , attrs] of inline) expect(attrs).toContain(`nonce="${nonce}"`);
      expect(html).toContain(`nonce: '${nonce}'`);
    },
  );

  it('standalone は ../dist/xsheet.min.js を 1 つだけ読み（zip を作る道具がこの場所を直す）、esm は ../dist/xsheet.mjs を import する', () => {
    expect(pages.standalone.split('<script src="../dist/xsheet.min.js"></script>')).toHaveLength(2);
    expect(pages.standalone).toContain('XSheet.view(');
    expect(pages.esm).toContain("import { view } from '../dist/xsheet.mjs';");
  });
});
