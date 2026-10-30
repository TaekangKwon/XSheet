// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { read, view, type Viewer, type ViewOptions } from '../../src/facade/index';
import { MESSAGES_EN, MESSAGES_JA, MESSAGES_KO } from '../../src/render/messages';
import { buildXlsx } from '../../src/testing/make-xlsx';
import { stubClientSize, wait } from '../render/render-sheet.helpers';

/**
 * M6-2：入口の層 `view`・`read`（docs/adr/0003-universal-distribution.md §2）。
 * jsdom はレイアウトを持たない（高さ 0）ため、高さの既定は clientHeight で見る。見た目は examples で目視。
 */

/** 暗号化（OLE の先頭）のファイル。開けないので帯に文言が出る */
const OLE = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0, 0, 0]);

let book: Uint8Array;
const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
const viewers: Viewer[] = [];

beforeAll(async () => {
  book = await buildXlsx({
    sheets: [
      { name: 'One', rows: [['first']] },
      { name: 'Hid', state: 'hidden', rows: [['hidden']] },
      { name: 'Two', rows: [['second']] },
    ],
  });
});

afterEach(() => {
  viewers.splice(0).forEach((v) => v.destroy());
  document.body.replaceChildren();
  document.head.querySelectorAll('style').forEach((s) => s.remove());
  document.documentElement.lang = '';
});

function container(id = 'viewer'): HTMLElement {
  const el = document.createElement('div');
  el.id = id;
  document.body.append(el);
  return el;
}

async function open(
  target: HTMLElement | string,
  source: Parameters<typeof view>[1],
  options?: ViewOptions,
): Promise<Viewer> {
  const v = await view(target, source, options);
  viewers.push(v);
  return v;
}

/** 表示中のシートの A1 の文字 */
const a1 = (el: HTMLElement): string | null | undefined =>
  el.querySelector('tbody tr:not(.xs-spacer) td:not(.xs-rownum)')?.textContent;

const banner = (el: HTMLElement): string | null | undefined =>
  el.querySelector('.xs-banner')?.textContent;

describe('view：表示と Viewer', () => {
  it('セレクタの要素に div.xs-viewer > div.xs-workbook を足し、最初の表示シートを出す', async () => {
    const el = container();
    const v = await open('#viewer', book);
    const root = el.querySelector(':scope > .xs-viewer');
    expect(root).not.toBeNull();
    expect(root!.querySelector(':scope > .xs-workbook')).not.toBeNull();
    expect(a1(el)).toBe('first');
    expect(v.sheets.map((s) => [s.index, s.name, s.hidden])).toEqual([
      [1, 'One', false],
      [2, 'Hid', true],
      [3, 'Two', false],
    ]);
  });

  it('showSheet は番号でも名前でも替えられ、無い名前・番号は RangeError', async () => {
    const el = container();
    const v = await open(el, book);
    await v.showSheet('Two');
    expect(a1(el)).toBe('second');
    await v.showSheet(2);
    expect(a1(el)).toBe('hidden');
    await expect(v.showSheet('Nope')).rejects.toThrow(RangeError);
    await expect(v.showSheet(9)).rejects.toThrow(RangeError);
  });

  it('無い番号の showSheet も入口の層の RangeError（英語）', async () => {
    const v = await open(container(), book);
    await expect(v.showSheet(9)).rejects.toThrow(/^XSheet: no such sheet$/);
  });

  it('同じ container に続けて view すると、前の表示を外して差し替える', async () => {
    const el = container();
    const first = await open(el, book);
    const second = await open(el, OLE, { lang: 'en' });
    expect(el.querySelectorAll(':scope > .xs-viewer')).toHaveLength(1);
    expect(banner(el)).toBe(MESSAGES_EN.encrypted);
    first.destroy(); // 外した表示の destroy は何もしない（新しい表示を消さない）
    expect(el.querySelectorAll(':scope > .xs-viewer')).toHaveLength(1);
    second.destroy();
    expect(el.children).toHaveLength(0);
  });

  it('読み込み中に同じ container へ次の view が来たら、前の呼び出しは表示の無い Viewer を返す', async () => {
    const el = container();
    const [a, b] = await Promise.all([view(el, Promise.resolve(book)), view(el, book)]);
    viewers.push(a, b);
    expect(el.querySelectorAll(':scope > .xs-viewer')).toHaveLength(1);
    expect(a.sheets).toEqual([]);
    await expect(a.showSheet(1)).rejects.toThrow(RangeError);
    expect(b.sheets).toHaveLength(3);
    expect(a1(el)).toBe('first');
  });

  it('destroy で表示の根を外す（何度呼んでもよい）。container 自身は残る', async () => {
    const el = container();
    const v = await view(el, book);
    v.destroy();
    v.destroy();
    expect(el.children).toHaveLength(0);
    expect(el.isConnected).toBe(true);
  });

  it('開けないファイルは例外にせず帯に文言を出し、sheets は空', async () => {
    const el = container();
    const v = await open(el, OLE, { lang: 'ja' });
    expect(banner(el)).toBe(MESSAGES_JA.encrypted);
    expect(v.sheets).toEqual([]);
    await expect(v.showSheet(1)).rejects.toThrow(RangeError);
  });

  it('見つからないセレクタ・要素でない target は TypeError で、DOM に何も足さない', async () => {
    await expect(view('#missing', book)).rejects.toThrow(TypeError);
    await expect(view({} as HTMLElement, book)).rejects.toThrow(TypeError);
    expect(document.body.children).toHaveLength(0);
    expect(document.head.querySelector('style')).toBeNull();
  });
});

describe('view：source の形', () => {
  const copy = (bytes: Uint8Array): ArrayBuffer => bytes.slice().buffer;

  it('Blob・File・ArrayBuffer・型付き配列（途中から）・Response・Promise を受ける', async () => {
    const padded = new Uint8Array(book.length + 8);
    padded.set(book, 4);
    const sources = [
      new Blob([copy(book)]),
      new File([copy(book)], 'a.xlsx'),
      copy(book),
      padded.subarray(4, 4 + book.length),
      new DataView(copy(book)),
      new Response(copy(book)),
      Promise.resolve(new Response(copy(book))),
      Promise.resolve(copy(book)),
    ];
    for (const source of sources) {
      const el = container();
      await open(el, source);
      expect(a1(el), String(source)).toBe('first');
    }
  });

  it('失敗した Response・URL の文字列・ほかの形は TypeError（通信しない）', async () => {
    await expect(view(container(), new Response('x', { status: 404 }))).rejects.toThrow(/HTTP 404/);
    await expect(view(container(), 'https://example.com/a.xlsx' as never)).rejects.toThrow(
      /fetch\(url\)/,
    );
    await expect(view(container(), 42 as never)).rejects.toThrow(TypeError);
    expect(document.querySelectorAll('.xs-viewer')).toHaveLength(0);
  });
});

describe('view：言語（lang・messages）', () => {
  it('lang で表を選び、messages を上に重ねる', async () => {
    const ko = container('ko');
    await open(ko, OLE, { lang: 'ko' });
    expect(banner(ko)).toBe(MESSAGES_KO.encrypted);
    const en = container('en');
    await open(en, OLE, { lang: 'en', messages: { encrypted: 'locked' } });
    expect(banner(en)).toBe('locked');
  });

  it('auto（既定）は <html lang>、無ければブラウザの言語、合わなければ英語', async () => {
    document.documentElement.lang = 'ko-KR';
    const a = container('a');
    await open(a, OLE);
    expect(banner(a)).toBe(MESSAGES_KO.encrypted);

    document.documentElement.lang = 'ja';
    const b = container('b');
    await open(b, OLE, { lang: 'auto' });
    expect(banner(b)).toBe(MESSAGES_JA.encrypted);

    document.documentElement.lang = '';
    const languages = vi.spyOn(navigator, 'languages', 'get').mockReturnValue(['fr-FR']);
    const language = vi.spyOn(navigator, 'language', 'get').mockReturnValue('de');
    const c = container('c');
    await open(c, OLE);
    expect(banner(c)).toBe(MESSAGES_EN.encrypted);
    languages.mockReturnValue(['ja-JP']);
    const d = container('d');
    await open(d, OLE);
    expect(banner(d)).toBe(MESSAGES_JA.encrypted);
    languages.mockRestore();
    language.mockRestore();
  });

  it('<html lang> が合わない言語なら、ブラウザの言語を見ずに英語', async () => {
    document.documentElement.lang = 'zh-CN';
    const languages = vi.spyOn(navigator, 'languages', 'get').mockReturnValue(['ko-KR']);
    const el = container();
    await open(el, OLE);
    expect(banner(el)).toBe(MESSAGES_EN.encrypted);
    languages.mockRestore();
  });

  it('型を通らない lang（地域付き・大文字・知らない言語）も日本語に落ちない', async () => {
    document.documentElement.lang = 'en';
    const got = [];
    for (const lang of ['ko-KR', 'JA', 'zh', '', 'toString']) {
      const el = container();
      await open(el, OLE, { lang: lang as never });
      got.push(banner(el));
    }
    expect(got).toEqual([
      MESSAGES_KO.encrypted,
      MESSAGES_JA.encrypted,
      MESSAGES_EN.encrypted,
      MESSAGES_EN.encrypted,
      MESSAGES_EN.encrypted,
    ]);
  });

  it('韓国語の表は日本語・英語と同じ 23 キーを持ち、凍結されている', () => {
    expect(Object.keys(MESSAGES_KO).sort()).toEqual(Object.keys(MESSAGES_JA).sort());
    expect(Object.values(MESSAGES_KO).every((v) => v.trim() !== '')).toBe(true);
    expect(MESSAGES_KO.parseFailed).toContain('{detail}');
    expect(MESSAGES_KO.imageFormat).toContain('{format}');
    expect(Object.isFrozen(MESSAGES_KO)).toBe(true);
  });
});

describe('view：テーマ・CSS・高さ', () => {
  const rootOf = (el: HTMLElement): HTMLElement => el.querySelector<HTMLElement>('.xs-viewer')!;

  it('theme は表示の根の data-xs-theme（light は何も付けない）', async () => {
    const themes = ['light', 'dark', 'auto', undefined] as const;
    const got = [];
    for (const theme of themes) {
      const el = container();
      await open(el, book, { theme });
      got.push(rootOf(el).getAttribute('data-xs-theme'));
    }
    expect(got).toEqual([null, 'dark', 'auto', null]);
  });

  it('CSS を head に 1 回だけ入れ、nonce を付ける。css: none なら入れない', async () => {
    await open(container(), book, { css: 'none' });
    expect(document.head.querySelectorAll('style[data-xsheet-style]')).toHaveLength(0);
    await open(container(), book, { nonce: 'abc' });
    await open(container(), book);
    const styles = document.head.querySelectorAll<HTMLStyleElement>('style[data-xsheet-style]');
    expect(styles).toHaveLength(1);
    expect(styles[0]!.getAttribute('nonce')).toBe('abc');
    expect(styles[0]!.textContent).toContain('.xs-workbook');
    expect(styles[0]!.textContent).toContain("[data-xs-theme='dark']");
  });

  it('style.css を読み込み済み（.xs-workbook の規則が効く）なら CSS を入れない', async () => {
    const linked = document.createElement('style');
    linked.textContent = '.xs-workbook { display: flex; }';
    document.head.append(linked);
    await open(container(), book);
    expect(document.head.querySelector('style[data-xsheet-style]')).toBeNull();
  });

  it('使う側が :root に --xs-* を書いただけ（style.css は無い）なら CSS を入れる', async () => {
    const vars = document.createElement('style');
    vars.textContent = ':root { --xs-fg: #000; }';
    document.head.append(vars);
    await open(container(), book);
    expect(document.head.querySelectorAll('style[data-xsheet-style]')).toHaveLength(1);
  });

  it('Shadow DOM の中の container には、文書に CSS があってもその ShadowRoot に CSS を入れる', async () => {
    await open(container(), book);
    expect(document.head.querySelectorAll('style[data-xsheet-style]')).toHaveLength(1);
    const host = container('host');
    const shadow = host.attachShadow({ mode: 'open' });
    const el = document.createElement('div');
    shadow.append(el);
    await open(el, book);
    expect(shadow.querySelectorAll('style[data-xsheet-style]')).toHaveLength(1);
    expect(document.head.querySelectorAll('style[data-xsheet-style]')).toHaveLength(1);
  });

  it('container の高さが 0 なら 480px を付けて warn を 1 回だけ出す。高さがあれば付けない', async () => {
    warn.mockClear();
    const a = container();
    await open(a, book);
    const b = container();
    await open(b, book);
    expect(rootOf(a).style.height).toBe('480px');
    expect(rootOf(b).style.height).toBe('480px');
    // warn はモジュールで 1 回だけ（このファイルの最初の高さ 0 の表示より前に出ていれば 0 回）
    expect(warn.mock.calls.length).toBeLessThanOrEqual(1);
    const restore = stubClientSize(800, 600);
    try {
      const c = container();
      await open(c, book);
      expect(rootOf(c).style.height).toBe('');
    } finally {
      restore();
    }
  });

  it('height に数値（px）・CSS の長さを渡せる', async () => {
    const a = container();
    await open(a, book, { height: 300 });
    const b = container();
    await open(b, book, { height: '50vh' });
    expect(rootOf(a).style.height).toBe('300px');
    expect(rootOf(b).style.height).toBe('50vh');
  });

  it('表示の根に style 属性の文字列を書かない（CSSOM だけ。CSP の unsafe-inline が要らない）', async () => {
    const el = container();
    const setAttribute = vi.spyOn(Element.prototype, 'setAttribute');
    await open(el, book, { height: 200, theme: 'dark' });
    await wait();
    expect(setAttribute.mock.calls.filter(([name]) => name === 'style')).toEqual([]);
    setAttribute.mockRestore();
  });
});

describe('read', () => {
  it('view と同じ形を受けて DOM なしで開く', async () => {
    const opened = await read(Promise.resolve(new Blob([book.slice().buffer])));
    expect(opened.sheets.map((s) => s.name)).toEqual(['One', 'Hid', 'Two']);
    const sheet = await opened.openSheet(1);
    expect(sheet.getRows(1, 1)[0]?.cells[0]?.text).toBe('first');
    opened.close();
    expect(document.body.children).toHaveLength(0);
  });

  it('開けないファイルは XlsxError（kind で種類が分かる）、上限は limits で替えられる', async () => {
    await expect(read(OLE)).rejects.toMatchObject({ kind: 'encrypted' });
    await expect(read(book, { limits: { maxSheets: 1 } })).rejects.toMatchObject({
      kind: 'too-large',
    });
  });
});
