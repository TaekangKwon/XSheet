import { describe, expect, it } from 'vitest';
import type { CellFormat } from '../../src/core/index';
import {
  buildFormatCss,
  formatClassName,
  hasOverFill,
  spillInkGradient,
} from '../../src/render/format-css';

/**
 * 書式 CSS の生成（ISSUE-100・ADR-0030 §決定 1・4〜6）：純関数の文字列を確かめる。
 * 規則は 1 書式 1 行。`.xs-sheet[data-xs-fmt="<連番>"] td.xs-xf-<番号> { … }` の形。
 */

const rgb = (n: number): { rgb: number; alpha: number } => ({ rgb: n, alpha: 1 });
const WHITE = 'var(--xs-xlsx-cell-bg)';

/** 書式 1 件の規則の本文（`{ … }` の中。連番 1・書式番号 0 で作った 1 行から取り出す） */
function body(f: CellFormat): string {
  const css = buildFormatCss([f], 1);
  const m = /^\.xs-sheet\[data-xs-fmt="1"\] td\.xs-xf-0 \{ (.*) \}$/.exec(css);
  expect(m, css).not.toBeNull();
  return m![1]!;
}

/** 宣言を `プロパティ → 値` の表にする */
function decls(f: CellFormat): Record<string, string> {
  const out: Record<string, string> = {};
  for (const d of body(f).split(';')) {
    const t = d.trim();
    if (t === '') continue;
    const i = t.indexOf(':');
    out[t.slice(0, i).trim()] = t.slice(i + 1).trim();
  }
  return out;
}

describe('TD-100-1 塗りありの規則（白以外）', () => {
  it('背景・右と下のグリッド線は塗りの色。自動の文字は黒（ライト・ダークとも）', () => {
    expect(decls({ fill: rgb(0xff0000) })).toEqual({
      'background-color': '#ff0000',
      color: '#000000',
      'border-right-color': '#ff0000',
      'border-bottom-color': '#ff0000',
    });
  });

  it('明示の文字色はそのまま（塗りの上でも）', () => {
    const d = decls({ fill: rgb(0x0000ff), color: rgb(0xffffff) });
    expect(d['background-color']).toBe('#0000ff');
    expect(d['color']).toBe('#ffffff');
  });

  it('色は 6 桁の小文字 #rrggbb（先頭の 0 を埋める）', () => {
    expect(decls({ fill: rgb(0x00ff01) })['background-color']).toBe('#00ff01');
    expect(decls({ fill: rgb(0x000001) })['background-color']).toBe('#000001');
  });
});

describe('TD-100-2 塗りなしの規則', () => {
  it('自動の文字は color を書かない（画面の文字色を継承）。装飾だけの規則になる', () => {
    expect(decls({ bold: true })).toEqual({ 'font-weight': '700' });
  });

  it('明示の文字色は #rrggbb。グリッド線・背景は書かない', () => {
    expect(decls({ color: rgb(0xff0000) })).toEqual({ color: '#ff0000' });
  });
});

describe('TD-100-3 白の塗り（紙・ADR-0030 1.1）', () => {
  it('#FFFFFF は背景も右・下のグリッド線も面の変数。自動の文字は color なし', () => {
    expect(decls({ fill: rgb(0xffffff) })).toEqual({
      'background-color': WHITE,
      'border-right-color': WHITE,
      'border-bottom-color': WHITE,
    });
  });

  it('明示の文字色はそのまま', () => {
    expect(decls({ fill: rgb(0xffffff), color: rgb(0xff0000) })['color']).toBe('#ff0000');
  });

  it('#FEFEFE・#F2F2F2 は Excel の色のまま（紙にしない）', () => {
    for (const n of [0xfefefe, 0xf2f2f2]) {
      const d = decls({ fill: rgb(n) });
      const hex = `#${n.toString(16)}`;
      expect(d['background-color']).toBe(hex);
      expect(d['border-right-color']).toBe(hex);
      expect(d['color']).toBe('#000000');
    }
  });

  it('hasOverFill：白以外の塗りだけ true（白の塗り・塗りなし・書式なしは false）', () => {
    expect(hasOverFill({ fill: rgb(0xff0000) })).toBe(true);
    expect(hasOverFill({ fill: rgb(0xffffff) })).toBe(false);
    expect(hasOverFill({ bold: true })).toBe(false);
    expect(hasOverFill(undefined)).toBe(false);
  });
});

describe('TD-100-4 文字の装飾', () => {
  it('太字・斜体', () => {
    expect(decls({ bold: true, italic: true })).toEqual({
      'font-weight': '700',
      'font-style': 'italic',
    });
  });

  it('下線（一重）は underline だけ。二重は text-decoration-style: double を足す', () => {
    expect(decls({ underline: 'single' })).toEqual({ 'text-decoration-line': 'underline' });
    expect(decls({ underline: 'double' })).toEqual({
      'text-decoration-line': 'underline',
      'text-decoration-style': 'double',
    });
  });

  it('取り消し線は line-through。下線と両方なら 2 語並べる', () => {
    expect(decls({ strike: true })).toEqual({ 'text-decoration-line': 'line-through' });
    expect(decls({ strike: true, underline: 'single' })['text-decoration-line']).toBe(
      'underline line-through',
    );
    expect(decls({ strike: true, underline: 'double' })).toEqual({
      'text-decoration-line': 'underline line-through',
      'text-decoration-style': 'double',
    });
  });

  it('文字サイズの比は calc(<比> * 1em)', () => {
    expect(decls({ sizeRatio: 1.5 })).toEqual({ 'font-size': 'calc(1.5 * 1em)' });
    expect(decls({ sizeRatio: 0.82 })['font-size']).toBe('calc(0.82 * 1em)');
    expect(decls({ sizeRatio: 4 })['font-size']).toBe('calc(4 * 1em)');
  });

  it('範囲外・非数の比は書かない（0.5〜4 の外・NaN・Infinity）', () => {
    for (const r of [0.1, 5, Number.NaN, Number.POSITIVE_INFINITY, -1]) {
      expect(buildFormatCss([{ sizeRatio: r, bold: true }], 1)).not.toContain('font-size');
    }
  });
});

describe('TD-100-5 セレクタ・書式番号・空', () => {
  it('セレクタは表示の連番の属性と書式番号のクラスで絞る。規則は書式ごとに 1 行', () => {
    const css = buildFormatCss([{ bold: true }, { italic: true }, { fill: rgb(0x112233) }], 7);
    const lines = css.split('\n');
    expect(lines).toHaveLength(3);
    expect(lines[0]).toMatch(/^\.xs-sheet\[data-xs-fmt="7"\] td\.xs-xf-0 \{ /);
    expect(lines[1]).toMatch(/^\.xs-sheet\[data-xs-fmt="7"\] td\.xs-xf-1 \{ /);
    expect(lines[2]).toMatch(/^\.xs-sheet\[data-xs-fmt="7"\] td\.xs-xf-2 \{ /);
  });

  it('formatClassName は xs-xf-<番号>', () => {
    expect(formatClassName(0)).toBe('xs-xf-0');
    expect(formatClassName(4095)).toBe('xs-xf-4095');
  });

  it('書式が空なら空の文字列。宣言が 1 つも出ない書式は規則を作らない', () => {
    expect(buildFormatCss([], 1)).toBe('');
    expect(buildFormatCss([{}], 1)).toBe('');
    // 規則を作らなくても、後ろの書式の番号はずれない
    expect(buildFormatCss([{}, { bold: true }], 1)).toContain('td.xs-xf-1 {');
  });

  it('連番が整数でなければ 0 にする（属性値に変な文字を入れない）', () => {
    expect(buildFormatCss([{ bold: true }], Number.NaN)).toContain('data-xs-fmt="0"');
    expect(buildFormatCss([{ bold: true }], 1.5)).toContain('data-xs-fmt="0"');
  });
});

describe('TD-100-6 CSS にファイル由来の文字列が入らない', () => {
  it('宣言は固定の語・整数・#rrggbb・小数の比・var(--xs-xlsx-cell-bg) だけ', () => {
    const formats: CellFormat[] = [
      {
        fill: rgb(0xff0000),
        color: rgb(0x00ff00),
        bold: true,
        italic: true,
        strike: true,
        underline: 'double',
        sizeRatio: 1.25,
      },
      { fill: rgb(0xffffff), color: rgb(0x123456) },
      { fill: rgb(0xf2f2f2) },
    ];
    const css = buildFormatCss(formats, 3);
    // 値として許す語を取り除くと、記号類だけが残る
    const stripped = css
      .replace(/\.xs-sheet\[data-xs-fmt="\d+"\] td\.xs-xf-\d+/g, '')
      .replace(/var\(--xs-xlsx-cell-bg\)/g, '')
      .replace(/#[0-9a-f]{6}\b/g, '')
      .replace(/calc\(\d+(\.\d+)? \* 1em\)/g, '')
      .replace(
        /background-color|border-right-color|border-bottom-color|color|font-weight|font-style|text-decoration-line|text-decoration-style|font-size|italic|underline|line-through|double|700/g,
        '',
      );
    expect(stripped.replace(/[\s{}:;]/g, '')).toBe('');
    expect(css).not.toMatch(/[<>'\\]/); // < > ' \ を含まない
    expect(css).not.toMatch(/url\(|expression|@import/i);
  });
});

describe('spillInkGradient（はみ出した文字の色の塗り分け・ADR-0030 §決定 9）', () => {
  const YELLOW: CellFormat = { fill: rgb(0xffff00) };
  const BLUE: CellFormat = { fill: rgb(0xccddff) };
  const PAPER: CellFormat = { fill: rgb(0xffffff) };
  const BOLD: CellFormat = { bold: true };
  const RED_INK: CellFormat = { fill: rgb(0xffff00), color: rgb(0xff0000) };
  const WHITE_INK: CellFormat = { color: rgb(0xffffff) };
  /** 0 黄・1 青・2 白（紙）・3 太字だけ・4 黄＋赤の文字・5 白い文字だけ */
  const table = [YELLOW, BLUE, PAPER, BOLD, RED_INK, WHITE_INK];
  const BLACK = '#000000';
  const INK = 'var(--xs-fg)';

  it('TD-100-n: 塗りありの元（自動色＝黒）が右の塗りなしの空セルへはみ出すと、列ごとに 黒 → 画面の文字色', () => {
    expect(spillInkGradient(table, [0, undefined, undefined], 0, [64, 64, 64])).toBe(
      `linear-gradient(to right, ${BLACK} 0px 64px, ${INK} 64px 192px)`,
    );
  });

  it('TD-100-n: 塗りなしの元（自動色＝画面の文字色）が右の塗りありの空セルへはみ出すと、画面の文字色 → 黒', () => {
    expect(spillInkGradient(table, [undefined, 1, 1], 0, [64, 64, 64])).toBe(
      `linear-gradient(to right, ${INK} 0px 64px, ${BLACK} 64px 192px)`,
    );
  });

  it('TD-100-n: 白の塗り（紙）・塗りなしの装飾・書式なし・範囲外の番号は「塗りなし」の色（画面の文字色）', () => {
    expect(spillInkGradient(table, [0, 2, 3, undefined, 99], 0, [10, 20, 30, 40, 50])).toBe(
      `linear-gradient(to right, ${BLACK} 0px 10px, ${INK} 10px 150px)`,
    );
  });

  it('TD-100-n: 隣り合う同じ色は 1 区間にまとめ、離れた同じ色は別の区間（区切りは累積の px）', () => {
    expect(spillInkGradient(table, [0, undefined, undefined, 1], 0, [64, 20, 30, 64])).toBe(
      `linear-gradient(to right, ${BLACK} 0px 64px, ${INK} 64px 114px, ${BLACK} 114px 178px)`,
    );
  });

  it('TD-100-n: 全部の列が同じ色なら塗り分けない（null）。すべて塗りあり・すべて塗りなし・1 列だけ', () => {
    expect(spillInkGradient(table, [0, 1, 0], 0, [64, 64, 64])).toBeNull();
    expect(spillInkGradient(table, [undefined, 2, 3], 0, [64, 64, 64])).toBeNull();
    expect(spillInkGradient(table, undefined, 0, [64, 64])).toBeNull();
    expect(spillInkGradient(table, [0], 0, [64])).toBeNull();
    expect(spillInkGradient(table, [0, undefined], 0, [])).toBeNull();
  });

  it('TD-100-n: 元の文字が明示色（赤・白など）なら塗り分けない（単色のまま・Excel と同じ）', () => {
    expect(spillInkGradient(table, [4, undefined, undefined], 0, [64, 64, 64])).toBeNull();
    expect(spillInkGradient(table, [5, 1, 1], 0, [64, 64, 64])).toBeNull();
  });

  it('TD-100-n: c が 0 でない位置の元から数える（元の左端が 0px）。元より左の列は見ない', () => {
    expect(spillInkGradient(table, [0, undefined, 1, 1], 1, [50, 50, 50])).toBe(
      `linear-gradient(to right, ${INK} 0px 50px, ${BLACK} 50px 150px)`,
    );
  });

  it('TD-100-n: 幅 0（非表示の列）は区間に数えない。不正な幅（負・非数）も 0 扱い', () => {
    expect(spillInkGradient(table, [0, 1, undefined], 0, [64, 0, 64])).toBe(
      `linear-gradient(to right, ${BLACK} 0px 64px, ${INK} 64px 128px)`,
    );
    expect(spillInkGradient(table, [0, 1, undefined], 0, [64, Number.NaN, 64])).toBe(
      `linear-gradient(to right, ${BLACK} 0px 64px, ${INK} 64px 128px)`,
    );
    expect(spillInkGradient(table, [0, undefined, 1], 0, [64, Number.NaN, -5])).toBeNull(); // 残るのは 1 区間だけ
  });

  it('TD-100-n: 文字列は固定の語・整数の px・#rrggbb・var(--xs-fg) だけ（ファイル由来の文字列が入らない）', () => {
    const g = spillInkGradient(table, [0, undefined, 1], 0, [64.4, 64.6, 64]);
    expect(g).not.toBeNull();
    const stripped = g!
      .replace(/linear-gradient\(to right, /, '')
      .replace(/var\(--xs-fg\)/g, '')
      .replace(/#000000/g, '')
      .replace(/\d+px/g, '');
    expect(stripped.replace(/[\s,)]/g, '')).toBe('');
  });
});
