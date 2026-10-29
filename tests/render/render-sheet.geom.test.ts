// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { openXlsx, type SheetDrawing } from '../../src/core/index';
import { MESSAGES_JA } from '../../src/render/messages';
import { buildXlsx, shapeXml, twoCellAnchor, wsDr } from '../../src/testing/make-xlsx';
import { stubClientSize, wait } from './render-sheet.helpers';

/**
 * 形の表（geom）の動的 import（M2-3）。render は geom を型だけ静的に引き、形の表と評価器は `loadGeom` で読む。
 * - 読む前の `renderShape` は、形のある図形も代替（種類名の破線枠）を返す（評価できない形と同じ道）。
 * - 表示は形のある図形があるシートを開く時だけ読み、読めなければ代替で描いて通知（`libLoadFailed`）を出す。
 * 前準備（`setup.ts`）が読み込んだ状態を外すため、`vi.resetModules` で render を読み直す。
 */

let restoreSize: (() => void) | null = null;

afterEach(() => {
  restoreSize?.();
  restoreSize = null;
  vi.doUnmock('../../src/geom/index');
  vi.resetModules();
  document.body.replaceChildren();
});

/** 枠線のある矩形の図形 */
const RECT: SheetDrawing = {
  anchor: {
    type: 'oneCell',
    from: { col: 0, colOff: 0, row: 0, rowOff: 0 },
    cx: 9525 * 40,
    cy: 9525 * 20,
  },
  kind: 'shape',
  name: 'r',
  descr: '',
  mediaPath: null,
  geom: { type: 'preset', prst: 'rect', adj: [] },
  hasOutline: true,
};

describe('geom の動的 import', () => {
  it('読む前の renderShape は代替を返し、loadGeom の後は形で描く（読むのは 1 回）', async () => {
    vi.resetModules();
    const shapes = await import('../../src/render/shapes');
    const rect = { x: 0, y: 0, w: 40, h: 20 };
    const sheet = { totalWidth: 1000, totalHeight: 1000 };
    expect(shapes.isGeomLoaded()).toBe(false);
    expect(shapes.renderShape(rect, RECT, sheet).kind).toBe('fallback');
    const p1 = shapes.loadGeom();
    expect(shapes.loadGeom()).toBe(p1);
    await p1;
    expect(shapes.isGeomLoaded()).toBe(true);
    expect(shapes.renderShape(rect, RECT, sheet).kind).toBe('drawn');
  });

  it('形の表を読めなければ、図形は代替（種類名の破線枠）で描き、通知の行に libLoadFailed を出す。表は出る', async () => {
    vi.resetModules();
    vi.doMock('../../src/geom/index', () => {
      throw new Error('読めない');
    });
    const { renderSheet } = await import('../../src/render/render-sheet');
    const bytes = await buildXlsx({
      font: { name: 'Calibri', size: 11 },
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: {
            xml: wsDr([
              twoCellAnchor(
                { col: 1, row: 1 },
                { col: 3, row: 3 },
                shapeXml({ name: 's', prst: 'rect' }),
              ),
            ]),
          },
        },
      ],
    });
    const book = await openXlsx(bytes);
    const container = document.createElement('div');
    document.body.append(container);
    restoreSize = stubClientSize(800, 400);
    const view = await renderSheet(container, book, 1);
    await wait();
    const notice = container.querySelector<HTMLElement>('.xs-notice')!;
    expect(notice.hidden).toBe(false);
    expect(notice.textContent).toBe(MESSAGES_JA.libLoadFailed);
    expect(container.querySelector('.xs-figure .xs-frame')?.textContent).toBe(MESSAGES_JA.shape);
    expect(container.querySelector('table.xs-grid')).not.toBeNull();
    view.destroy();
    book.close();
  });

  it('形のある図形が無いシートは geom を読まない', async () => {
    vi.resetModules();
    const shapes = await import('../../src/render/shapes');
    const { renderSheet } = await import('../../src/render/render-sheet');
    const book = await openXlsx(await buildXlsx({ sheets: [{ name: 'S', rows: [['a']] }] }));
    const container = document.createElement('div');
    const view = await renderSheet(container, book, 1);
    expect(shapes.isGeomLoaded()).toBe(false);
    view.destroy();
    book.close();
  });
});
