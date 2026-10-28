// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createGeometry, type CellFormat } from '../../src/core/index';
import { emptyLayout } from '../../src/core/internal';
import { buildGridTable } from '../../src/render/grid-table';
import { buildFormatCss } from '../../src/render/format-css';
import type { GridCell, GridRow } from '../../src/render/grid';

/**
 * 書式の重さの計測（ISSUE-100・ADR-0030 §決定 11。InterHub `xlsx-format.perf.test.ts` を写したもの）：全セルに書式（塗り 50 色・太字・文字の色・下線を混ぜる）を持つ
 * 10,000 セルの合成シートで、table の構築時間を書式なしと比べる（`borders.perf.test.ts` と同じ形）。
 * 書式のあるセルの追加の書き込みは td の className の 1 語（と、はみ出しの塗りの判定）だけ。
 * 加えて、`maxCellFormats`（4,096）個の書式の CSS の生成時間を測る。jsdom は実ブラウジングより遅いので余裕を見た値。
 * 実機（TV-44-4）のスクロールの 200ms 超のフレームは目視で確かめる。
 */

const ROWS = 100;
const COLS = 100;
const geo = createGeometry(emptyLayout(), ROWS, COLS);

/** 塗り 50 色 × 太字・文字の色・下線の有無を混ぜた書式の表（件数は 50×4＝200） */
function formatTable(): CellFormat[] {
  const out: CellFormat[] = [];
  for (let i = 0; i < 50; i++) {
    const fill = { rgb: 0x808080 + i * 0x010203, alpha: 1 };
    out.push({ fill });
    out.push({ fill, bold: true });
    out.push({ fill, color: { rgb: 0xff0000, alpha: 1 }, underline: 'single' });
    out.push({ fill, bold: true, italic: true, underline: 'double', sizeRatio: 1.25 });
  }
  return out;
}

/** 100 行 × 100 列の文字・数値のセル。withFormats のとき全セルに書式の番号を付ける */
function synthetic(withFormats: boolean, formatCount: number): GridRow[] {
  const rows: GridRow[] = [];
  for (let r = 1; r <= ROWS; r++) {
    const cells: GridCell[] = [];
    for (let c = 0; c < COLS; c++) {
      cells.push(
        (r + c) % 3 === 0
          ? { text: String(r * c), kind: 'n', colSpan: 1, rowSpan: 1 }
          : { text: `r${r}c${c}`, kind: 's', colSpan: 1, rowSpan: 1 },
      );
    }
    const row: GridRow = { r, cells };
    if (withFormats) {
      row.formats = Array.from(
        { length: COLS },
        (_, c) => (r * 7 + c * 3 + ((r * c) % 5)) % formatCount,
      );
    }
    rows.push(row);
  }
  return rows;
}

/** table の構築を runs 回測り、最小の時間（ms）を返す（ウォームアップ 1 回を除く） */
function bestOf(runs: number, build: () => HTMLTableElement): { best: number; times: number[] } {
  build();
  const times: number[] = [];
  for (let i = 0; i < runs; i++) {
    const t0 = performance.now();
    build();
    times.push(performance.now() - t0);
  }
  return { best: Math.min(...times), times };
}

describe('TD-100-20 重さ：10,000 セル（全セル書式あり）', () => {
  it('table の構築が、書式なしの目安 1.2 倍以内（jsdom の誤差の余裕 5ms を足す）', () => {
    const table = formatTable();
    const plain = synthetic(false, table.length);
    const formatted = synthetic(true, table.length);
    const range = { first: 1, last: COLS };
    const build = (rows: GridRow[], formats: CellFormat[]) => () =>
      buildGridTable((c) => `C${c}`, COLS, rows, geo, ROWS, 7, range, formats);

    // 交互に測って、ゆらぎ（GC・JIT）を両方に均す
    const a = bestOf(4, build(plain, []));
    const b = bestOf(4, build(formatted, table));
    const a2 = bestOf(4, build(plain, []));
    const b2 = bestOf(4, build(formatted, table));
    const noFmt = Math.min(a.best, a2.best);
    const withFmt = Math.min(b.best, b2.best);
    console.log(
      `[TD-100-20] 10,000 セルの table 構築（最小）：書式なし ${noFmt.toFixed(1)}ms／全セル書式 ${withFmt.toFixed(1)}ms` +
        `（比 ${(withFmt / noFmt).toFixed(2)}）。書式なし ${[...a.times, ...a2.times].map((t) => t.toFixed(0)).join('/')}／書式あり ${[...b.times, ...b2.times].map((t) => t.toFixed(0)).join('/')} ms`,
    );

    // 書式のクラスが全 td に付いている（測っているものが正しい）
    const t = build(formatted, table)();
    expect(t.querySelectorAll('td[class^="xs-xf-"], td[class*=" xs-xf-"]').length).toBe(
      ROWS * COLS,
    );
    expect(withFmt).toBeLessThanOrEqual(noFmt * 1.2 + 5);
  }, 60_000); // jsdom の 10,000 セルの構築は 1 回 0.25 秒前後（計 20 回ほど）

  it('maxCellFormats（4,096）個の書式の CSS の生成が目安 20ms 以内。規則は 1 書式 1 行', () => {
    const formats: CellFormat[] = [];
    for (let i = 0; i < 4096; i++) {
      formats.push({
        fill: { rgb: (i * 0x010203) & 0xffffff, alpha: 1 },
        color: i % 3 === 0 ? { rgb: 0xff0000, alpha: 1 } : undefined,
        bold: i % 2 === 0 ? true : undefined,
        underline: i % 5 === 0 ? 'double' : undefined,
        sizeRatio: i % 7 === 0 ? 1.5 : undefined,
      });
    }
    buildFormatCss(formats, 1); // ウォームアップ
    const times: number[] = [];
    let css = '';
    for (let i = 0; i < 5; i++) {
      const t0 = performance.now();
      css = buildFormatCss(formats, 1);
      times.push(performance.now() - t0);
    }
    const best = Math.min(...times);
    console.log(
      `[TD-100-20] 4,096 個の書式の CSS 生成：${times.map((t) => t.toFixed(1)).join(' / ')} ms（最小 ${best.toFixed(1)}ms・${css.length} 文字）`,
    );
    expect(css.split('\n')).toHaveLength(4096);
    expect(best).toBeLessThan(20);
  });
});
