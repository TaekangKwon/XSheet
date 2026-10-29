// `src/style.css` の文面を読むテスト（*-css.test.ts・style-css.test.ts）の共通の部品。
// jsdom は CSS を適用しないため、規則は文面で確かめる（実際の見た目は M3-2 の examples で目視）。
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect } from 'vitest';

/** `src/style.css` の文面（改行は LF にそろえる）。jsdom でも読めるよう、リポジトリの根（vitest の root）から辿る */
export const css = readFileSync(join(process.cwd(), 'src/style.css'), 'utf8').replace(
  /\r\n/g,
  '\n',
);

/** 変数の既定値のブロックのセレクタ */
export const DEFAULTS = ':where(:root, :host)';

/** セレクタ（行頭〜`{`）で始まる最初の規則の本文 */
export function rule(selector: string): string {
  const i = css.indexOf(`\n${selector} {`);
  expect(i, `規則 ${selector}`).toBeGreaterThanOrEqual(0);
  return css.slice(i, css.indexOf('}', i));
}

/** 変数の宣言（`--name: 値;`）の値をすべて、出現順に */
export function declaredValues(name: string): string[] {
  const re = new RegExp(`${name.replace(/-/g, '\\-')}:\\s*([^;]+);`, 'g');
  return [...css.matchAll(re)].map((m) => m[1]!.trim());
}
