// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { findText } from '../../src/core/index';
import { applyRuns, runSlicer } from '../../src/render/rich-text';
import { estimateCellTextWidth } from '../../src/render/spill-extent';
import { fontXml, type CellSpec } from '../../src/testing/make-xlsx';
import { cleanupViews, mount, openBook, wait } from './render-sheet.helpers';

/**
 * リッチテキストの区切りの描き方（M7-17・ADR-0009 §2・§3）。
 * DOM の形と当てた書式を jsdom で確かめ、実機の見え方は依頼のサンプル（06_検索 B9）で目視する。
 */

afterEach(() => {
  cleanupViews();
});

const red = { rgb: 0xff0000, alpha: 1 };

describe('applyRuns', () => {
  it('区切りを span.xs-runs 1 つで包み、書式の無い区切りはテキスト、ある区切りは span に書式', () => {
    const box = document.createElement('div');
    box.textContent = 'abcdef';
    const slice = applyRuns(
      box,
      'abcdef',
      'abcdef',
      [
        { text: 'ab' },
        { text: 'cd', font: { bold: true, color: red, underline: 'double' } },
        { text: 'ef', font: {} },
      ],
      undefined,
    );
    expect(slice).toBeDefined();
    const wrap = box.firstElementChild as HTMLElement;
    expect(box.childNodes).toHaveLength(1);
    expect(wrap.className).toBe('xs-runs');
    expect(wrap.textContent).toBe('abcdef');
    const [a, b, c] = [...wrap.childNodes];
    expect(a!.nodeType).toBe(Node.TEXT_NODE);
    const sb = (b as HTMLElement).style;
    expect(sb.color).toBe('rgb(255, 0, 0)');
    expect(sb.fontWeight).toBe('700');
    expect(sb.fontStyle).toBe('normal');
    expect(sb.textDecorationLine).toBe('underline');
    expect(sb.textDecorationStyle).toBe('double');
    // 書式のある区切りはセルの書式を継がない（太字なし）。セルに色が無ければ色は当てない（自動の色を継ぐ）
    const sc = (c as HTMLElement).style;
    expect(sc.fontWeight).toBe('400');
    expect(sc.color).toBe('');
  });

  it('はみ出しの span.xs-text の中を置き換える。区切りが無い・### なら何もしない', () => {
    const box = document.createElement('div');
    const label = document.createElement('span');
    label.className = 'xs-text';
    label.textContent = 'xy';
    box.append(label);
    applyRuns(box, 'xy', 'xy', [{ text: 'x', font: { italic: true } }, { text: 'y' }], undefined);
    expect(label.querySelector('.xs-runs')?.textContent).toBe('xy');

    const plain = document.createElement('div');
    plain.textContent = 'xy';
    expect(applyRuns(plain, 'xy', 'xy', undefined, undefined)).toBeUndefined();
    expect(applyRuns(plain, 'xy', '###', [{ text: 'xy', font: {} }], undefined)).toBeUndefined();
    expect(plain.innerHTML).toBe('xy');
  });

  it('色の無い区切りは、セルに色があれば自動の色（塗りの上は黒）。大きさはセルの比に対する em', () => {
    const [onPaper] = runSlicer('a', [{ text: 'a', font: {} }], { color: red })(
      0,
      1,
    ) as HTMLElement[];
    expect(onPaper!.style.color).toBe('var(--xs-fg)');
    const [onFill] = runSlicer('a', [{ text: 'a', font: { sizeRatio: 2 } }], {
      color: red,
      fill: { rgb: 0x00ff00, alpha: 1 },
      sizeRatio: 1.25,
    })(0, 1) as HTMLElement[];
    expect(onFill!.style.color).toBe('rgb(0, 0, 0)');
    expect(onFill!.style.fontSize).toBe('1.6em');
  });

  it('作り手は区切りをまたぐ範囲を区切りごとに分ける', () => {
    const slice = runSlicer(
      '検索テスト後',
      [{ text: '検索' }, { text: 'テスト', font: { bold: true } }, { text: '後' }],
      undefined,
    );
    const nodes = slice(1, 4);
    expect(nodes.map((n) => n.textContent)).toEqual(['索', 'テス']);
    expect(nodes[1]).toBeInstanceOf(HTMLSpanElement);
    expect(slice(0, 0)).toEqual([]);
  });
});

describe('リッチテキストのセル（renderSheet）', () => {
  const rich: CellSpec = {
    t: 'inlineStr',
    is: '<r><t>検索</t></r><r><rPr><b val="1"/><color rgb="FFFF0000"/></rPr><t>テスト</t></r><r><t>（書式の境目）</t></r>',
  };

  async function richSheet() {
    const book = await openBook({
      sheets: [{ name: 'S', rows: [['label', rich]] }],
      fonts: [fontXml('<sz val="11"/>')],
      xfs: [0],
    });
    const m = await mount(book);
    await wait();
    return m;
  }

  const cellOf = (root: HTMLElement) =>
    [...root.querySelectorAll<HTMLElement>('.xs-cell')].find((e) =>
      e.textContent?.startsWith('検索'),
    )!;

  it('区切りごとに描き、はみ出しの span の中も区切る', async () => {
    const { root } = await richSheet();
    const cell = cellOf(root);
    expect(cell.textContent).toBe('検索テスト（書式の境目）');
    const runs = cell.querySelector('.xs-runs')!;
    const styled = runs.querySelector('span')!;
    expect(styled.textContent).toBe('テスト');
    expect(styled.style.color).toBe('rgb(255, 0, 0)');
    expect(styled.style.fontWeight).toBe('700');
  });

  it('区切りをまたぐ一致は 1 つの片のまま、片の中を区切りで分ける', async () => {
    const { root, view, book } = await richSheet();
    view.setHighlights((await findText(await book.openSheet(1), '検索テスト')).matches);
    await wait();
    const cell = cellOf(root);
    const pieces = cell.querySelectorAll('.xs-hit');
    expect(pieces).toHaveLength(1);
    expect(pieces[0]!.textContent).toBe('検索テスト');
    const inner = pieces[0]!.querySelector('span')!;
    expect(inner.textContent).toBe('テスト');
    expect(inner.style.color).toBe('rgb(255, 0, 0)');
    expect(cell.textContent).toBe('検索テスト（書式の境目）');
    expect(cell.querySelectorAll('.xs-hits')).toHaveLength(1);
  });
});

describe('はみ出しの幅の見積もり（区切り）', () => {
  it('区切りの大きさ・太字で数え、書式の無い区切りはセルの書式', () => {
    const plain = estimateCellTextWidth('abcd');
    const big = estimateCellTextWidth('abcd', undefined, [
      { text: 'ab' },
      { text: 'cd', font: { sizeRatio: 3 } },
    ]);
    expect(big).toBeGreaterThan(plain);
    const cellBold = { bold: true as const };
    // 書式のある区切りはセルの太字を継がない
    expect(estimateCellTextWidth('ab', cellBold, [{ text: 'ab', font: {} }])).toBe(
      estimateCellTextWidth('ab'),
    );
    expect(estimateCellTextWidth('ab', cellBold, [{ text: 'ab' }])).toBe(
      estimateCellTextWidth('ab', cellBold),
    );
  });
});
