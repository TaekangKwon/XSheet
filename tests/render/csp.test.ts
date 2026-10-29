// @vitest-environment jsdom
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { cleanupViews, mountWorkbook, stubClientSize, wait } from './render-sheet.helpers';

/**
 * M3-3：CSP の要件（docs/01_requirements.md N7・docs/02_architecture.md §3「CSP」）。
 * render が CSP で止まる書き方をしないこと（style 属性・HTML の文字列・data: の画像）と、
 * 書式の `<style>` が nonce を持てることを確かめる。実際に CSP の下で動くことは examples（nonce だけの style-src）で見る。
 */

afterEach(() => {
  cleanupViews();
});

/** render のソース（コメントを除く）を { ファイル名: 文面 } で。`http://…` の `//` はコメントとみなさない */
const renderDir = join(process.cwd(), 'src/render');
const sources = Object.fromEntries(
  readdirSync(renderDir)
    .filter((name) => name.endsWith('.ts'))
    .map((name) => [
      name,
      readFileSync(join(renderDir, name), 'utf8').replace(
        /\/\*[\s\S]*?\*\/|(?<![:'"])\/\/.*$/gm,
        '',
      ),
    ]),
);

/** パターンに当たるファイル名 */
function filesMatching(re: RegExp): string[] {
  return Object.keys(sources).filter((name) => re.test(sources[name]!));
}

describe('render のソース：CSP で止まる書き方をしない', () => {
  it('style 属性を setAttribute で付けない（unsafe-inline が要る。CSSOM の el.style.x は CSP で止まらない）', () => {
    expect(filesMatching(/setAttribute\(\s*['"]style['"]/)).toEqual([]);
  });

  it('HTML の文字列を DOM にしない（innerHTML・outerHTML・insertAdjacentHTML・document.write）', () => {
    expect(filesMatching(/\.innerHTML\b|\.outerHTML\b|insertAdjacentHTML|document\.write/)).toEqual(
      [],
    );
  });

  it('<style> を作るのは書式の 1 か所だけで、nonce を付けられる', () => {
    expect(filesMatching(/createElement\(\s*['"]style['"]/)).toEqual(['sheet-view.ts']);
    expect(sources['sheet-view.ts']).toMatch(/setAttribute\(\s*'nonce'/);
  });

  it('画像は blob URL だけで出す（data: の URL を作らない。img-src は blob: だけでよい）', () => {
    expect(filesMatching(/['"`]data:|toDataURL|readAsDataURL/)).toEqual([]);
    expect(filesMatching(/createObjectURL/)).toEqual(['image-store.ts']);
  });
});

describe('描いた DOM に style 属性を文字列で付けない', () => {
  // setAttribute を汎用の部品（名前を引数で受ける）から呼ぶ所もあるため、ソースの文面だけでなく実際に描いて確かめる
  it.each(['ref-drawings.xlsx', 'ref-shapes.xlsx'])('%s', async (file) => {
    const names = new Set<string>();
    const original = Element.prototype.setAttribute;
    const spy = vi.spyOn(Element.prototype, 'setAttribute').mockImplementation(function (
      this: Element,
      name: string,
      value: string,
    ) {
      names.add(name.toLowerCase());
      original.call(this, name, value);
    });
    const restore = stubClientSize(1600, 1200);
    try {
      const bytes = readFileSync(join(process.cwd(), 'fixtures/excel', file));
      const { root } = await mountWorkbook(new Uint8Array(bytes));
      await wait(50);
      expect(root.querySelectorAll('.xs-layer .xs-figure').length).toBeGreaterThan(0);
    } finally {
      restore();
      spy.mockRestore();
    }
    expect(names.size).toBeGreaterThan(0);
    expect(names.has('style')).toBe(false);
  });
});

describe('nonce の受け渡し', () => {
  const spec = { sheets: [{ name: 'S', rows: [['a']] }] };

  it('renderWorkbook の nonce は書式の <style> に付く', async () => {
    const { root } = await mountWorkbook(spec, { nonce: 'n0nce-123' });
    const styles = root.querySelectorAll('style');
    expect(styles).toHaveLength(1);
    expect(styles[0]!.getAttribute('nonce')).toBe('n0nce-123');
  });

  it('nonce を指定しなければ付けない', async () => {
    const { root } = await mountWorkbook(spec);
    expect(root.querySelector('style')!.hasAttribute('nonce')).toBe(false);
  });
});
