import { describe, expect, it } from 'vitest';
import { openXlsx } from '../../src/core/index';
import { makePng } from '../../src/testing/make-image';
import {
  alternateContent,
  buildXlsx,
  chartXml,
  graphicFrameXml,
  grpSpXml,
  picXml,
  spXml,
  twoCellAnchor,
  wsDr,
  type DrawingRelSpec,
} from '../../src/testing/make-xlsx';

const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CHARTEX_URI = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const A1 = { col: 0, row: 0 };
const B2 = { col: 1, row: 1 };
const CHART_REL: DrawingRelSpec = { id: 'rId1', target: '../charts/chart1.xml', type: 'chart' };

/** 図の XML（アンカーの並び）と rels を持つ 1 シートのブックを作り、openDrawings の結果を返す（グラフの部品は既定で chart1.xml を持つ） */
async function drawingsOf(
  anchors: string[],
  opts: { rels?: DrawingRelSpec[]; charts?: Record<string, string> } = {},
) {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: [['a']],
        drawing: { xml: wsDr(anchors), rels: opts.rels ?? [CHART_REL] },
      },
    ],
    charts: opts.charts ?? { 'chart1.xml': chartXml() },
    extraFiles: { 'xl/media/image1.png': makePng(40, 30) },
  });
  const book = await openXlsx(bytes);
  return book.openDrawings(1);
}

describe('TD-101-1 openDrawings：グラフのパス（chartPath）', () => {
  it('chart の graphicFrame は c:chart r:id を rels で解いた zip のパスを持つ。kind は chart のまま', async () => {
    const { drawings, notices } = await drawingsOf([
      twoCellAnchor(A1, B2, graphicFrameXml('chart', CHART_URI, 5, [0, 0, 0, 0], 'rId1')),
    ]);
    expect(notices).toEqual([]);
    expect(drawings).toHaveLength(1);
    expect(drawings[0]).toMatchObject({
      kind: 'chart',
      name: 'chart',
      mediaPath: null,
      chartPath: 'xl/charts/chart1.xml',
    });
  });

  it('c:chart の名前空間の接頭辞は問わない（局所名で引く）', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(
        A1,
        B2,
        graphicFrameXml('chart', CHART_URI, 5, [0, 0, 0, 0], 'rId1', 'ch:chart'),
      ),
    ]);
    expect(drawings[0]!.chartPath).toBe('xl/charts/chart1.xml');
  });

  it('c:chart 無し・r:id が空・rels に無い Id・外部・zip に無い・chartex・SmartArt は持たない（kind は今どおり）', async () => {
    const { drawings, notices } = await drawingsOf(
      [
        twoCellAnchor(A1, B2, graphicFrameXml('noChild')),
        twoCellAnchor(A1, B2, graphicFrameXml('emptyId', CHART_URI, 5, [0, 0, 0, 0], '')),
        twoCellAnchor(A1, B2, graphicFrameXml('unknownId', CHART_URI, 5, [0, 0, 0, 0], 'rId9')),
        twoCellAnchor(A1, B2, graphicFrameXml('external', CHART_URI, 5, [0, 0, 0, 0], 'rId2')),
        twoCellAnchor(A1, B2, graphicFrameXml('missing', CHART_URI, 5, [0, 0, 0, 0], 'rId3')),
        twoCellAnchor(
          A1,
          B2,
          graphicFrameXml('chartex', CHARTEX_URI, 5, [0, 0, 0, 0], 'rId1', 'cx:chart'),
        ),
        twoCellAnchor(A1, B2, graphicFrameXml('dgm', DIAGRAM_URI, 5, [0, 0, 0, 0], 'rId1')),
      ],
      {
        rels: [
          CHART_REL,
          {
            id: 'rId2',
            target: 'https://example.invalid/chart1.xml',
            type: 'chart',
            external: true,
          },
          { id: 'rId3', target: '../charts/none.xml', type: 'chart' },
        ],
      },
    );
    expect(notices).toEqual([]);
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([
      ['chart', 'noChild'],
      ['chart', 'emptyId'],
      ['chart', 'unknownId'],
      ['chart', 'external'],
      ['chart', 'missing'],
      ['chart', 'chartex'],
      ['smartart', 'dgm'],
    ]);
    for (const d of drawings) expect('chartPath' in d).toBe(false);
  });

  it('グループの中・AlternateContent の中のグラフも chartPath を持つ', async () => {
    const frame = (name: string) => graphicFrameXml(name, CHART_URI, 5, [0, 0, 100, 100], 'rId1');
    const { drawings } = await drawingsOf([
      twoCellAnchor(A1, B2, grpSpXml('grp', [frame('inGroup')], 6, { box: [0, 0, 100, 100] })),
      twoCellAnchor(A1, B2, alternateContent(frame('inChoice'), spXml('fallback'))),
    ]);
    expect(drawings.map((d) => [d.name, d.chartPath])).toEqual([
      ['inGroup', 'xl/charts/chart1.xml'],
      ['inChoice', 'xl/charts/chart1.xml'],
    ]);
  });

  it('図の一覧（順・kind・名前・画像のパス）と通知は、グラフに r:id があっても今と同じ', async () => {
    const { drawings, notices } = await drawingsOf(
      [
        twoCellAnchor(A1, B2, picXml({ name: 'pic', embed: 'rId2' })),
        twoCellAnchor(A1, B2, graphicFrameXml('chart', CHART_URI, 5, [0, 0, 0, 0], 'rId1')),
        twoCellAnchor(A1, B2, spXml('rect')),
      ],
      { rels: [CHART_REL, { id: 'rId2', target: '../media/image1.png' }] },
    );
    expect(notices).toEqual([]);
    expect(drawings.map((d) => [d.kind, d.name, d.mediaPath])).toEqual([
      ['image', 'pic', 'xl/media/image1.png'],
      ['chart', 'chart', null],
      ['shape', 'rect', null],
    ]);
  });
});
