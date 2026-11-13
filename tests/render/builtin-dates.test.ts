// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { cleanupViews, mountWorkbook, stubClientSize, wait } from './render-sheet.helpers';

/** renderWorkbook の builtinDates（M7-8・ADR-0007 §6）。表示の文字は core が読み取りのときに作る */

let restoreSize: () => void;
beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
});
afterEach(() => {
  cleanupViews();
  restoreSize();
});

const spec = {
  sheets: [{ name: 'S', rows: [[{ v: 46337, s: 1 }]], cols: [{ min: 1, max: 1, width: 40 }] }],
  xfs: [0, 28],
};
const cellText = (root: HTMLElement) =>
  root.querySelector('tbody tr[aria-rowindex="2"] td:not(.xs-rownum)')?.textContent;

describe('renderWorkbook の builtinDates', () => {
  it("'japanese-era' で組み込みの 28 を和暦で描く。既定は西暦", async () => {
    const era = await mountWorkbook(spec, { builtinDates: 'japanese-era' });
    const plain = await mountWorkbook(spec);
    await wait();
    expect(cellText(era.root)).toBe('令和8年11月11日');
    expect(cellText(plain.root)).toBe('2026/11/11');
  });
});
