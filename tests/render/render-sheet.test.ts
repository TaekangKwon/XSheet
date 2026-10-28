// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createGeometry, openXlsx, XlsxError, type XlsxBook } from '../../src/core/index';
import { MESSAGES_JA, formatMessage } from '../../src/render/messages';
import { renderSheet } from '../../src/render/render-sheet';
import { createSheetView, xlsxErrorMessage } from '../../src/render/sheet-view';
import { buildXlsx, buildZip, type CellSpec } from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  dataRows,
  headTexts,
  mount,
  openBook,
  px,
  renderSpec,
  stubClientSize,
  wait,
} from './render-sheet.helpers';

/**
 * `renderSheet` の DOM 描画の振る舞いテスト（InterHub `spreadsheet.render.test.ts` を写したもの。ISSUE-82・84・ADR-0025）。
 * jsdom はレイアウトを持たない（高さ 0）ので、実寸（行の位置・畳みの描画）・スクロール連動は目視で確認し、
 * ここでは DOM の構造（見出し・結合・仮想化の行数・エラー文言・固定レイアウトの colgroup／行高／余白の値・
 * 非表示の畳み・はみ出しの入れ物の幅）を検証する。文言は既定（`MESSAGES_JA`）。
 * タブ・シートタブ列・.xls の経路のテストは写さない（タブは M2-4 の renderWorkbook、.xls は対象外）。
 */

afterEach(() => {
  cleanupViews();
});

describe('renderSheet（.xlsx）', () => {
  it('TD-82-2: 見出しは列記号・行番号は Excel の行番号（先頭行を見出しにしない）', async () => {
    const rows: CellSpec[][] = [
      ['表題', null, null],
      ['名前', '値', '備考'],
      ['a', 1, 'x'],
    ];
    const root = await renderSpec({ sheets: [{ name: 'S', rows }] });
    expect(root.className).toBe('xs-sheet');
    expect(headTexts(root)).toEqual(['', 'A', 'B', 'C']);
    const trs = dataRows(root);
    expect(trs.map((tr) => tr.querySelector('.xs-rownum')?.textContent)).toEqual(['1', '2', '3']);
    // 先頭行（表題）も本文の 1 行目として出る
    expect(trs[0]?.cells[1]?.textContent).toBe('表題');
    expect(trs[2]?.cells[2]?.textContent).toBe('1');
  });

  it('TD-82-2: セル値は textContent のみ（HTML 風の文字列が要素にならない）・結合は colSpan/rowSpan', async () => {
    const rows: CellSpec[][] = [
      ['<img src=x onerror=alert(1)>', null, 'D1'],
      [null, null, 'D2'],
    ];
    const root = await renderSpec({ sheets: [{ name: 'S', rows, merges: ['A1:B2'] }] });
    expect(root.querySelector('img')).toBeNull();
    const first = dataRows(root)[0]!;
    // 行番号 td + 結合セル + D1（結合で隠れる B1 の td は無い）
    expect(first.cells).toHaveLength(3);
    const merged = first.cells[1]!;
    expect(merged.textContent).toBe('<img src=x onerror=alert(1)>');
    expect(merged.colSpan).toBe(2);
    expect(merged.rowSpan).toBe(2);
    // 2 行目は 結合に覆われるため 行番号 td + D2 のみ
    expect(dataRows(root)[1]?.cells).toHaveLength(2);
  });

  it('TD-82-2: 根は container の子の div.xs-sheet（container には何も付けない）。子は通知の行・スクロール枠・<style>', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: [['a']] }] });
    const { root, container } = await mount(book);
    expect(container.getAttribute('class')).toBeNull();
    expect(container.hasAttribute('data-xs-fmt')).toBe(false);
    expect(root.parentElement).toBe(container);
    expect([...root.children].map((c) => c.tagName.toLowerCase() + '.' + c.className)).toEqual([
      'div.xs-notice',
      'div.xs-body',
      'style.',
    ]);
    expect(root.querySelector<HTMLElement>('.xs-notice')!.hidden).toBe(true);
  });

  it('TD-82-2: 空シートは noData（シートの番号は SheetInfo.index。非表示のシートも番号で開ける）', async () => {
    const book = await openBook({
      sheets: [
        { name: 'Hid', state: 'hidden', rows: [['h']] },
        { name: 'E', rows: [] },
      ],
    });
    const { root } = await mount(book, 2);
    expect(root.querySelector('.xs-empty')?.textContent).toBe(MESSAGES_JA.noData);
    expect(root.querySelector('.xs-grid')).toBeNull();
    const hid = await mount(book, 1);
    expect(dataRows(hid.root)[0]?.cells[1]?.textContent).toBe('h');
  });

  it('TD-82-2: ブックに無い番号は RangeError（container に何も足さない）', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: [['a']] }] });
    const none = await openBook({ sheets: [] });
    const container = document.createElement('div');
    await expect(renderSheet(container, book, 0)).rejects.toBeInstanceOf(RangeError);
    await expect(renderSheet(container, book, 2)).rejects.toBeInstanceOf(RangeError);
    await expect(renderSheet(container, none, 1)).rejects.toBeInstanceOf(RangeError);
    expect(container.childElementCount).toBe(0);
    book.close();
    none.close();
  });

  it('TD-82-3: 大きなシートでも DOM に出す行は見えている行＋前後だけ（行仮想化）', async () => {
    const parts: string[] = [];
    for (let r = 1; r <= 30000; r++) {
      parts.push(`<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>row${r}</t></is></c></row>`);
    }
    const root = await renderSpec({ sheets: [{ name: 'Big', sheetDataXml: parts.join('') }] });
    const rows = dataRows(root);
    expect(rows.length).toBeGreaterThan(0);
    expect(rows.length).toBeLessThan(200);
    expect(rows[0]?.cells[1]?.textContent).toBe('row1');
    // 描かない下側の行は余白行の高さで確保する（総行数 30000 より小さいと崩れる）
    const spacer = root.querySelector<HTMLElement>('.xs-spacer:last-child td');
    expect(parseFloat(spacer?.style.height ?? '0')).toBeGreaterThan(100000);
    expect(root.querySelector('table')?.getAttribute('aria-rowcount')).toBe('30001');
  });

  it('TD-82-4: 読み取りの例外を種類ごとの文言にする（暗号化・zip でない・Strict・壊れた zip）', async () => {
    const reason = async (bytes: Uint8Array): Promise<string> => {
      try {
        await openXlsx(bytes);
      } catch (e) {
        return xlsxErrorMessage(e, MESSAGES_JA);
      }
      throw new Error('開けてしまった');
    };
    expect(await reason(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0, 0, 0]))).toBe(
      MESSAGES_JA.encrypted,
    );
    const parseFailed = formatMessage(MESSAGES_JA.parseFailed, { detail: '' });
    expect(await reason(new TextEncoder().encode('plain text'))).toMatch(
      new RegExp(`^${parseFailed}xlsx: not-zip`),
    );
    const strict = await buildXlsx({ sheets: [{ name: 'S', rows: [['a']] }], strict: true });
    expect(await reason(strict)).toBe(MESSAGES_JA.unsupported);
    const brokenZip = await buildZip({ 'hello.txt': 'no workbook here' });
    expect(await reason(brokenZip)).toMatch(new RegExp(`^${parseFailed}xlsx: broken`));
    expect(xlsxErrorMessage(new XlsxError('too-large'), MESSAGES_JA)).toBe(MESSAGES_JA.tooLarge);
    expect(xlsxErrorMessage('x', { ...MESSAGES_JA, parseFailed: 'E:{detail}' })).toBe('E:x');
  });

  it('TD-82-4: シートを開けない（openSheet の例外）ときは、スクロール枠の中に種類ごとの文言の div.xs-banner を出す', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: [['a']] }] });
    const failing: XlsxBook = {
      sheets: book.sheets,
      openSheet: () => Promise.reject(new XlsxError('too-large')),
      openDrawings: (i) => book.openDrawings(i),
      readImage: (i, p) => book.readImage(i, p),
      close: () => book.close(),
    };
    const { root } = await mount(failing);
    const banner = root.querySelector('.xs-body > .xs-banner');
    expect(banner?.textContent).toBe(MESSAGES_JA.tooLarge);
    expect(root.querySelector('.xs-grid')).toBeNull();
  });

  it('TD-82-5: 読み込み中に destroy されたら結果を捨てる（根は外れ、表は描かない）', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: [['a']] }] });
    const container = document.createElement('div');
    const view = createSheetView(container, book, { messages: MESSAGES_JA, drawings: true });
    const pending = view.showSheet(1);
    view.destroy();
    await pending;
    expect(container.childElementCount).toBe(0);
    expect(view.root.childElementCount).toBe(0);
    expect(view.root.querySelector('table')).toBeNull();
    expect(view.activeIndex).toBeNull();
    book.close();
  });

  it('TD-82-5: destroy で描画済みの本文を空にし、根を container から外す（ブックは閉じず、再描画できる）', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: [['a']] }] });
    const { root, view, container } = await mount(book);
    expect(root.querySelector('table')).not.toBeNull();
    view.destroy();
    view.destroy(); // 2 回目は何もしない
    expect(root.querySelector('table')).toBeNull();
    expect(root.hasAttribute('data-xs-fmt')).toBe(false);
    expect(container.childElementCount).toBe(0);
    const again = await mount(book);
    expect(dataRows(again.root)[0]?.cells[1]?.textContent).toBe('a');
  });
});

describe('renderSheet（.xlsx・固定レイアウト／ISSUE-84）', () => {
  it('TD-84-7: colgroup は「行番号の列＋シートの列数」の col を持ち、幅は geometry の列幅（非表示は 0＋畳みクラス）', async () => {
    const rows: CellSpec[][] = [['a', 'b', 'c', 'd', 'e']];
    const bytes = await buildXlsx({
      font: { name: 'Calibri', size: 11 },
      sheets: [
        {
          name: 'S',
          rows,
          cols: [
            { min: 2, max: 2, width: 20 },
            { min: 3, max: 3, hidden: true },
            { min: 4, max: 4, width: 6 },
          ],
        },
      ],
    });
    const { root } = await mount(await openXlsx(bytes));
    const table = root.querySelector<HTMLTableElement>('table.xs-grid')!;
    const cols = [...table.querySelectorAll<HTMLElement>('colgroup > col')];
    expect(cols).toHaveLength(1 + 5);

    // geometry と同じ値（同じシートを開き直して算出）
    const book = await openXlsx(bytes);
    const sh = await book.openSheet(1);
    const geo = createGeometry(sh.layout, sh.rowCount, sh.colCount);
    book.close();
    const widths = cols.slice(1).map((c) => px(c.style.width));
    expect(widths).toEqual([1, 2, 3, 4, 5].map((c) => geo.colWidth(c)));
    // Calibri 11（MDW 7）：既定 64px・width=20 は 140px・非表示は 0・width=6 は 42px
    expect(widths).toEqual([64, 140, 0, 42, 64]);
    expect(cols.map((c) => c.classList.contains('xs-collapsed'))).toEqual([
      false,
      false,
      false,
      true,
      false,
      false,
    ]);
    // 表の幅は「行番号の列＋totalWidth」。行番号の列は 1 行（1 桁）なので最小の 32px
    expect(px(cols[0]!.style.width)).toBe(32);
    expect(px(table.style.width)).toBe(32 + geo.totalWidth);
  });

  it('TD-84-8: 行の高さは tr の height とセル内側の要素（罫線 1px を引いた高さ）で固定し、非表示の行は DOM に出さない', async () => {
    const rows: CellSpec[][] = [['r1'], ['r2'], ['r3'], ['r4'], ['r5']];
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      sheets: [
        { name: 'S', rows, rowAttrs: { 2: { ht: 30 }, 3: { hidden: true }, 4: { ht: 0.75 } } },
      ],
    });
    const trs = dataRows(root);
    // 行 3（非表示）は tr を作らない。位置は上下の余白（rowTop の差）と各 tr の高さで決まる
    expect(trs.map((tr) => tr.getAttribute('aria-rowindex'))).toEqual(['2', '3', '5', '6']);
    expect(trs.map((tr) => px(tr.style.height))).toEqual([20, 40, 1, 20]);
    expect(root.querySelector('tr.xs-collapsed')).toBeNull();
    // セル内側の要素の高さ＝行高 − 下罫線 1px（行番号の td も同じ）
    const boxHeights = trs.map((tr) =>
      [...tr.querySelectorAll<HTMLElement>('.xs-cell')].map((b) => px(b.style.height)),
    );
    expect(boxHeights).toEqual([
      [19, 19],
      [39, 39],
      [0, 0],
      [19, 19],
    ]);
    // 内側の要素の中身は textContent のみ
    expect(trs[1]!.cells[1]!.querySelector('.xs-cell')?.textContent).toBe('r2');
    expect(trs[1]!.cells[1]!.textContent).toBe('r2');
  });

  it('TD-84-9: 上の余白＋描いた行の高さ＋下の余白＝geometry の totalHeight（可変の行高・非表示を含む）', async () => {
    const parts: string[] = [];
    for (let r = 1; r <= 5000; r++) {
      const attrs = r === 3 ? ' hidden="1"' : r === 7 ? ' ht="60" customHeight="1"' : '';
      parts.push(
        `<row r="${r}"${attrs}><c r="A${r}" t="inlineStr"><is><t>v${r}</t></is></c></row>`,
      );
    }
    const bytes = await buildXlsx({
      font: { name: 'Calibri', size: 11 },
      sheets: [{ name: 'Big', sheetDataXml: parts.join('') }],
    });
    const { root } = await mount(await openXlsx(bytes));
    const book = await openXlsx(bytes);
    const sh = await book.openSheet(1);
    const geo = createGeometry(sh.layout, sh.rowCount, sh.colCount);
    book.close();
    // 行 3 は 0・行 7 は 60pt＝80px・その他 20px
    expect([geo.rowHeight(2), geo.rowHeight(3), geo.rowHeight(7)]).toEqual([20, 0, 80]);

    const [top, bottom] = [...root.querySelectorAll<HTMLElement>('.xs-spacer')].map((tr) =>
      px(tr.style.height),
    );
    const drawn = dataRows(root).reduce((sum, tr) => sum + px(tr.style.height), 0);
    expect(top).toBe(0);
    expect(top! + drawn + bottom!).toBe(geo.totalHeight);
    expect(bottom).toBeGreaterThan(80000);
    // 余白の td も同じ高さ
    expect(px(root.querySelector<HTMLElement>('.xs-spacer td')!.style.height)).toBe(0);
  });

  it('TD-84-13: フィルターで大半を畳んだ大きいシートでも、DOM の tr は見える行の数（前後 40 行ぶん）だけ', async () => {
    const parts: string[] = [];
    for (let r = 1; r <= 100000; r++) {
      const hid = r % 500 === 1 ? '' : ' hidden="1"';
      parts.push(
        `<row r="${r}"${hid}><c r="A${r}" t="inlineStr"><is><t>row${r}</t></is></c></row>`,
      );
    }
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      sheets: [{ name: 'F', sheetDataXml: parts.join('') }],
    });
    const trs = dataRows(root);
    // 見える行は 1・501・1001・…（jsdom は枠の高さ 0 なので、先頭から overscan の 40 行 + 先頭の 1 行）
    expect(trs).toHaveLength(41);
    expect(trs.map((tr) => Number(tr.getAttribute('aria-rowindex')) - 1).slice(0, 3)).toEqual([
      1, 501, 1001,
    ]);
    // 上＋描いた行＋下の余白＝総高さ（見える行 200 個 × 20px）
    const [top, bottom] = [...root.querySelectorAll<HTMLElement>('.xs-spacer')].map((tr) =>
      px(tr.style.height),
    );
    expect(top! + trs.length * 20 + bottom!).toBe(200 * 20);
  });

  it('TD-84-14: 結合の先頭の行が非表示でも、見える部分に値が出る（最初の見える行へ移り rowSpan を縮める）', async () => {
    const rows: CellSpec[][] = [];
    for (let r = 1; r <= 6; r++) rows.push([r === 2 ? 'MERGED' : null, `b${r}`]);
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      sheets: [
        {
          name: 'S',
          rows,
          merges: ['A2:A5'],
          rowAttrs: { 2: { hidden: true }, 3: { hidden: true } },
        },
      ],
    });
    const trs = dataRows(root);
    expect(trs.map((tr) => tr.getAttribute('aria-rowindex'))).toEqual(['2', '5', '6', '7']);
    // 行 4 の A に値。見える行は 4・5 の 2 行ぶん（rowSpan 2）。行 5 には A の td が無い
    const a4 = trs[1]!.cells[1]!;
    expect(a4.textContent).toBe('MERGED');
    expect(a4.rowSpan).toBe(2);
    expect(px(a4.querySelector<HTMLElement>('.xs-cell')!.style.height)).toBe(2 * 20 - 1);
    expect(trs[2]!.cells[1]!.textContent).toBe('b5'); // 行 5 は B のみ
  });

  it('TD-84-15: 非表示の列のセルは文字の入れ物（内側の要素）を作らない（隣へ文字が出ない）', async () => {
    const rows: CellSpec[][] = [['a1', 'HIDDEN-TEXT', null, 'd1']];
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      sheets: [{ name: 'S', rows, cols: [{ min: 2, max: 2, hidden: true }] }],
    });
    const tr = dataRows(root)[0]!;
    expect(tr.cells[1]!.querySelector('.xs-cell')).not.toBeNull(); // A
    expect(tr.cells[2]!.textContent).toBe(''); // B（非表示）は文字を出さない
    expect(tr.cells[2]!.querySelector('.xs-cell')).toBeNull();
    expect(tr.cells[4]!.querySelector('.xs-cell')).not.toBeNull(); // D
  });

  it('TD-84-16: 表はスクロール枠の幅・高さを埋めるまで空セルで広げる（足した列・行にも罫線とはみ出し）', async () => {
    const restore = stubClientSize(800, 400);
    try {
      const root = await renderSpec({
        font: { name: 'Calibri', size: 11 },
        sheets: [{ name: 'S', rows: [['タイトル', 'b'], ['x']] }],
      });
      // 幅 800・行番号の列 32px（行数 21 → 2 桁）→ 64px 列を 13 列、高さ 400px → 20px 行を 21 行
      expect(headTexts(root)).toHaveLength(1 + 13);
      expect(root.querySelectorAll('colgroup > col')).toHaveLength(1 + 13);
      const trs = dataRows(root);
      expect(trs).toHaveLength(21);
      expect(root.querySelector('table')?.getAttribute('aria-rowcount')).toBe('22');
      // 足した空セルも td（罫線あり）。どの行も行番号 td ＋ 13 個
      expect(trs.every((tr) => tr.cells.length === 14)).toBe(true);
      // 最後の列（足した列）の文字もはみ出せる：A1 は B1 が値なので 1 列だけ。B1 は右の足した空列 11 個まで
      const b1 = trs[0]!.cells[2]!.querySelector<HTMLElement>('.xs-cell')!;
      expect(b1.classList.contains('is-spill')).toBe(true);
      expect(px(b1.style.width)).toBe(12 * 64 - 1);
    } finally {
      restore();
    }
  });

  it('TD-84-10: td は 1 セル 1 個のまま（罫線を残す）。文字列は内側の要素だけを右の空セルまで広げる', async () => {
    const rows: CellSpec[][] = [
      ['長い文字列', null, null, '値', null],
      [1234567, null, null, 'x', null],
      ['a', 'b', null, null, null],
    ];
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      sheets: [{ name: 'S', rows }],
    });
    const trs = dataRows(root);
    const colCount = 4; // 値のある最後の列 D まで
    // どの行も、行番号 td ＋ 列数ぶんの td（colSpan で束ねない）
    for (const tr of trs) {
      expect(tr.cells).toHaveLength(1 + colCount);
      expect([...tr.cells].every((td) => td.colSpan === 1)).toBe(true);
    }
    const box = (r: number, c: number): HTMLElement | null =>
      trs[r]!.cells[1 + c]!.querySelector<HTMLElement>('.xs-cell');
    // 1 行目：A1 は B1・C1（空）の列幅を足した 64×3−1＝191px まで出せる。D1 は右隣が無いので広げない
    expect(box(0, 0)!.classList.contains('is-spill')).toBe(true);
    expect(px(box(0, 0)!.style.width)).toBe(64 * 3 - 1);
    expect(box(0, 0)!.querySelector('.xs-text')?.textContent).toBe('長い文字列');
    expect(trs[0]!.cells[1]!.textContent).toBe('長い文字列');
    expect(box(0, 3)!.classList.contains('is-spill')).toBe(false);
    // 2 行目：数値は自分の列だけ（広げない）。右隣は値のある D2 までの B2・C2 も空 td のまま
    expect(box(1, 0)!.classList.contains('is-spill')).toBe(false);
    expect(box(1, 1)).toBeNull();
    // 3 行目：値のあるセルが隣にある A3 は広げず、B3 は C3・D3（空）まで 3 列
    expect(box(2, 0)!.classList.contains('is-spill')).toBe(false);
    expect(px(box(2, 1)!.style.width)).toBe(64 * 3 - 1);
  });

  it('TD-84-18: はみ出しはセルの型で決まる（文字列は数字だけでもはみ出す・数値は切る・数式の "" の手前で止まる）', async () => {
    const rows: CellSpec[][] = [
      ['12345678901234', null, null, null],
      ['▲1,234', null, { t: 'str', f: 'IF(1,"","x")', v: '' }, 'z'],
      [12345678901234, null, null, null],
    ];
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      sheets: [{ name: 'S', rows }],
    });
    const trs = dataRows(root);
    const box = (r: number, c: number): HTMLElement | null =>
      trs[r]!.cells[1 + c]!.querySelector<HTMLElement>('.xs-cell');
    // 数字だけの文字列（inlineStr）は文字列なので、右の空セル 3 つ（合計 4 列）まではみ出す
    expect(px(box(0, 0)!.style.width)).toBe(64 * 4 - 1);
    // 「▲1,234」も文字列。数式の結果が "" のセル（C 列）は値のあるセルなので、手前（B 列まで）で止まる
    expect(px(box(1, 0)!.style.width)).toBe(64 * 2 - 1);
    expect(box(1, 2)).toBeNull(); // "" のセルは文字の入れ物を作らない（見た目は空）
    // 数値セルは自分の列で切る
    expect(box(2, 0)!.classList.contains('is-spill')).toBe(false);
  });

  it('TD-84-19: 型ごとの既定の配置（数値・日付は右、真偽・エラーは中央、文字列は左）', async () => {
    const rows: CellSpec[][] = [[1234, { v: 45306, s: 1 }, true, { t: 'e', v: '#N/A' }, '文字列']];
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      xfs: [0, 14],
      sheets: [{ name: 'S', rows }],
    });
    const tr = dataRows(root)[0]!;
    const cls = [1, 2, 3, 4, 5].map((i) => [...tr.cells[i]!.querySelector('.xs-cell')!.classList]);
    expect(cls[0]).toContain('is-right');
    expect(cls[1]).toContain('is-right');
    expect(cls[2]).toContain('is-center');
    expect(cls[3]).toContain('is-center');
    expect(cls[4]).not.toContain('is-right');
    expect(cls[4]).not.toContain('is-center');
    // 行番号の列は従来の配置（右）のまま
    expect(tr.cells[0]!.textContent).toBe('1');
  });

  it('TD-84-20: 収まらない数値・日付は ### にし、元の値を title に入れる（文字列は ### にしない）', async () => {
    const rows: CellSpec[][] = [[12345678901234, 1234567, '12345678901234', { v: 45306, s: 1 }]];
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      xfs: [0, 14],
      sheets: [{ name: 'S', rows, cols: [{ min: 4, max: 4, width: 4 }] }],
    });
    const tr = dataRows(root)[0]!;
    const box = (i: number): HTMLElement => tr.cells[i]!.querySelector<HTMLElement>('.xs-cell')!;
    // A：12345678901234（数値）は 64px 列に収まらない → ########（8 個）。元の値は title
    expect(box(1).textContent).toBe('########');
    expect(box(1).title).toBe('1.23457E+13');
    // B：7 桁は収まる
    expect(box(2).textContent).toBe('1234567');
    expect(box(2).title).toBe('');
    // C：文字列は ### にしない
    expect(box(3).textContent).toBe('12345678901234');
    // D：狭い列（width 4）の日付は収まらず # ＋ title に日付
    expect(box(4).textContent).toMatch(/^#+$/);
    expect(box(4).title).toBe('2024/1/15');
  });

  it('TD-84-11: 結合セルは colSpan／rowSpan のまま（はみ出しで広げない）・内側の高さは結合範囲の合計', async () => {
    const rows: CellSpec[][] = [
      ['M', null, 'x'],
      [null, null, 'y'],
      ['z', null, null],
    ];
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      sheets: [{ name: 'S', rows, merges: ['A1:B2'] }],
    });
    const merged = dataRows(root)[0]!.cells[1]!;
    expect([merged.colSpan, merged.rowSpan]).toEqual([2, 2]);
    expect(px(merged.querySelector<HTMLElement>('.xs-cell')!.style.height)).toBe(2 * 20 - 1);
  });

  it('枠の大きさが後から決まっても、取り付け後の次のフレームで実寸で描き直す', async () => {
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      sheets: [{ name: 'S', rows: [['a']] }],
    });
    expect(root.querySelectorAll('colgroup > col')).toHaveLength(1 + 1); // 枠の幅 0 では 1 列
    const restore = stubClientSize(800, 400);
    try {
      await wait();
      expect(root.querySelectorAll('colgroup > col')).toHaveLength(1 + 13);
    } finally {
      restore();
    }
  });
});
