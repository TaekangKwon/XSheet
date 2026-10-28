// @vitest-environment jsdom
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { SheetDrawing } from '../../src/core/index';
import { createImageLayer, frameLabel } from '../../src/render/image-layer';
import {
  formatMessage,
  MESSAGES_EN,
  MESSAGES_JA,
  resolveMessages,
} from '../../src/render/messages';

// M2-2：文言（docs/02_architecture.md §7）。InterHub の t() の代わりに、表を引数で回す
describe('文言の表', () => {
  it('日本語と英語は同じ 23 キーを持ち、空の文言が無い', () => {
    const keys = Object.keys(MESSAGES_JA).sort();
    expect(keys).toHaveLength(23);
    expect(Object.keys(MESSAGES_EN).sort()).toEqual(keys);
    for (const table of [MESSAGES_JA, MESSAGES_EN])
      for (const v of Object.values(table)) expect(v.trim()).not.toBe('');
  });

  it('既定の表は凍結されていて、使う側が書き換えられない', () => {
    expect(Object.isFrozen(MESSAGES_JA)).toBe(true);
    expect(Object.isFrozen(MESSAGES_EN)).toBe(true);
  });

  it('差し込みのある文言は parseFailed（{detail}）と imageFormat（{format}）だけ', () => {
    for (const table of [MESSAGES_JA, MESSAGES_EN]) {
      const withParams = Object.entries(table)
        .filter(([, v]) => /\{\w+\}/.test(v))
        .map(([k, v]) => [k, [...v.matchAll(/\{(\w+)\}/g)].map((m) => m[1])]);
      expect(Object.fromEntries(withParams)).toEqual({
        parseFailed: ['detail'],
        imageFormat: ['format'],
      });
    }
  });
});

describe('resolveMessages', () => {
  it('指定が無ければ既定（日本語）の表そのもの', () => {
    expect(resolveMessages()).toBe(MESSAGES_JA);
  });

  it('渡したキーだけを上書きし、残りは既定のまま', () => {
    const m = resolveMessages({ chart: 'Diagramm' });
    expect(m.chart).toBe('Diagramm');
    expect(m.shape).toBe(MESSAGES_JA.shape);
    expect(Object.isFrozen(m)).toBe(true);
    expect(MESSAGES_JA.chart).toBe('グラフ');
  });

  it('英語を土台にもできる', () => {
    expect(resolveMessages({ group: 'Gruppe' }, MESSAGES_EN)).toMatchObject({
      group: 'Gruppe',
      chart: 'Chart',
    });
  });

  it('文字列でない値・表に無いキー・__proto__ は無視する', () => {
    const input = JSON.parse('{"chart": 1, "unknownKey": "x", "__proto__": {"shape": "polluted"}}');
    const m = resolveMessages(input);
    expect(m.chart).toBe(MESSAGES_JA.chart);
    expect(m.shape).toBe(MESSAGES_JA.shape);
    expect(m).not.toHaveProperty('unknownKey');
    expect(({} as Record<string, unknown>).shape).toBeUndefined();
  });
});

describe('formatMessage', () => {
  it('{name} を差し込み、値の無いものはそのまま残す', () => {
    expect(formatMessage('画像（{format}）', { format: 'EMF' })).toBe('画像（EMF）');
    expect(formatMessage('a {x} b {y}', { x: 1 })).toBe('a 1 b {y}');
  });

  it('差し込む値の中の {…} はもう一度は置き換えない', () => {
    expect(formatMessage('{a}{b}', { a: '{b}', b: 'B' })).toBe('{b}B');
  });

  it('Object の組み込みの名前（{constructor} など）は差し込まない', () => {
    expect(formatMessage('{constructor}{toString}')).toBe('{constructor}{toString}');
  });
});

describe('文言が描画に届く', () => {
  const chart: SheetDrawing = {
    anchor: {
      type: 'twoCell',
      from: { col: 0, colOff: 0, row: 0, rowOff: 0 },
      to: { col: 1, colOff: 0, row: 1, rowOff: 0 },
    },
    kind: 'chart',
    name: 'c',
    descr: '',
    mediaPath: null,
  } as SheetDrawing;

  it('frameLabel は渡した表を使い、省略すると日本語', () => {
    expect(frameLabel(chart)).toBe('グラフ');
    expect(frameLabel(chart, undefined, MESSAGES_EN)).toBe('Chart');
    const image = { ...chart, kind: 'image' } as SheetDrawing;
    expect(frameLabel(image, { ok: false, reason: 'unsupported', label: 'EMF' }, MESSAGES_EN)).toBe(
      'Image (EMF)',
    );
  });

  it('createImageLayer に渡した文言が破線枠に出る', () => {
    const layer = createImageLayer({ messages: resolveMessages({ chart: '図表' }) });
    layer.setSheet({
      drawings: [chart],
      boxes: [{ index: 0, rect: { x: 0, y: 0, w: 40, h: 30 } }],
      readImage: async () => ({ ok: false, reason: 'missing', label: null }),
    } as unknown as Parameters<typeof layer.setSheet>[0]);
    layer.update({ x: 0, y: 0, w: 1000, h: 1000 });
    expect(layer.element.querySelector('.xs-frame')?.textContent).toBe('図表');
    layer.dispose();
  });
});

describe('InterHub の t() が残っていない', () => {
  it('src/render に t( の呼び出しと spreadsheet. のキーが無い', () => {
    // jsdom では import.meta.url が file: でないため、リポジトリの根（vitest の root）から辿る
    const dir = join(process.cwd(), 'src/render');
    for (const name of readdirSync(dir)) {
      const src = readFileSync(join(dir, name), 'utf8').replace(/\/\*[\s\S]*?\*\/|\/\/.*$/gm, '');
      expect(src, name).not.toMatch(/\bt\(/);
      expect(src, name).not.toMatch(/'spreadsheet\./);
    }
  });
});
