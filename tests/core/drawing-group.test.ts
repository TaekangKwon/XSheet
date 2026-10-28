/**
 * ISSUE-89：グループ（`xdr:grpSp`）の中身の展開（ADR-0027 §決定 5・6）。
 * 子（sp・cxnSp・pic・graphicFrame・入れ子の grpSp）を一覧に足し、グループの子の座標系（`SheetDrawing.group`）を持たせる。
 * 展開できないグループ（`chExt` が 0・座標系の無いもの・深さの上限超え）は、今までどおり破線枠（kind が group）。
 */
import { describe, expect, it } from 'vitest';
import { openXlsx, type XlsxLimits } from '../../src/core/index';
import { makePng } from '../../src/testing/make-image';
import {
  alternateContent,
  buildXlsx,
  graphicFrameXml,
  grpSpXml,
  lineShapeXml,
  picXml,
  shapeXml,
  twoCellAnchor,
  wsDr,
  type Box,
  type DrawingRelSpec,
  type GroupFrameSpec,
} from '../../src/testing/make-xlsx';

const PNG = makePng(40, 30);
const PIC_RELS: DrawingRelSpec[] = [{ id: 'rId1', target: '../media/image1.png' }];
const A1 = { col: 0, row: 0 };
const B2 = { col: 1, row: 1 };

/** 図の XML（アンカーの並び）を持つ 1 シートのブックを作り、openDrawings の結果を返す */
async function drawingsOf(anchors: string[], limits?: Partial<XlsxLimits>) {
  const bytes = await buildXlsx({
    sheets: [{ name: 'S', rows: [['a']], drawing: { xml: wsDr(anchors), rels: PIC_RELS } }],
    extraFiles: { 'xl/media/image1.png': PNG },
  });
  const book = await openXlsx(bytes, { limits });
  return book.openDrawings(1);
}

/** グループ 1 つを A1〜B2 のアンカーに置いて読む */
const readGroup = (group: string, limits?: Partial<XlsxLimits>) =>
  drawingsOf([twoCellAnchor(A1, B2, group)], limits);

const rect = (b: Box) => ({ x: b[0], y: b[1], cx: b[2], cy: b[3] });
/** 子の座標系が箱と同じグループの座標系 */
const frameOf = (box: Box, ch?: Box): GroupFrameSpec => ({ box, ch });

/** 4 つの子を持つ小さなグループの箱 */
const G: Box = [1000, 2000, 4000, 3000];

describe('TD-89-1 展開：子の種類 5 通り・XML の順・グループ自身は数えない', () => {
  it('sp（文字付き）・cxnSp・pic・graphicFrame・入れ子の grpSp の中の図が、XML の順に一覧へ出る。グループ自身は出ない', async () => {
    const { drawings, notices } = await readGroup(
      grpSpXml(
        'g',
        [
          shapeXml({ name: 'tb', txBox: true, box: [1000, 2000, 1000, 500], paras: ['hello'] }),
          lineShapeXml({ name: 'ln', box: [2000, 2000, 1000, 500] }),
          picXml({ name: 'pic', embed: 'rId1', box: [3000, 2000, 1000, 500] }),
          graphicFrameXml('chart', undefined, 5, [1000, 3000, 1000, 500]),
          grpSpXml(
            'inner',
            [shapeXml({ name: 'deep', box: [100, 200, 300, 400] })],
            7,
            frameOf([2000, 3000, 1000, 1000], [0, 0, 500, 500]),
          ),
        ],
        6,
        frameOf(G),
      ),
    );
    expect(notices).toEqual([]);
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([
      ['textbox', 'tb'],
      ['line', 'ln'],
      ['image', 'pic'],
      ['chart', 'chart'],
      ['shape', 'deep'],
    ]);
    expect(drawings.some((d) => d.kind === 'group')).toBe(false);
    // 子の画像は、グループの外と同じ経路（mediaPath）で読める
    expect(drawings[2]!.mediaPath).toBe('xl/media/image1.png');
    // 文字も ISSUE-87 の読み取りをそのまま使う
    expect(drawings[0]!.text?.paragraphs[0]?.runs[0]?.text).toBe('hello');
  });

  it('子の anchor は最も外側のグループのアンカー。group は座標系の連なりと自身の xfrm を持つ', async () => {
    const { drawings } = await readGroup(
      grpSpXml(
        'g',
        [shapeXml({ name: 'c', box: [1500, 2500, 800, 600] })],
        6,
        frameOf(G, [500, 1000, 2000, 1500]),
      ),
    );
    expect(drawings).toHaveLength(1);
    expect(drawings[0]!.anchor).toEqual({
      type: 'twoCell',
      from: { col: 0, colOff: 0, row: 0, rowOff: 0 },
      to: { col: 1, colOff: 0, row: 1, rowOff: 0 },
    });
    expect(drawings[0]!.group).toEqual({
      levels: [{ ch: rect([500, 1000, 2000, 1500]), own: null }],
      xfrm: rect([1500, 2500, 800, 600]),
    });
  });

  it('トップレベルの図（グループの外）は group を持たない', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(A1, B2, shapeXml({ name: 'top', box: [0, 0, 100, 100] })),
    ]);
    expect(drawings[0]!.group).toBeUndefined();
  });

  it('グループの外の図との重ね順は XML の順（グループの子は、そのグループの位置に入る）', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(A1, B2, shapeXml({ name: 'before' })),
      twoCellAnchor(
        A1,
        B2,
        grpSpXml(
          'g',
          [
            picXml({ name: 'a', embed: 'rId1', box: [1000, 2000, 10, 10] }),
            picXml({ name: 'b', embed: 'rId1', box: [1000, 2000, 10, 10] }),
          ],
          6,
          frameOf(G),
        ),
      ),
      twoCellAnchor(A1, B2, shapeXml({ name: 'after' })),
    ]);
    expect(drawings.map((d) => d.name)).toEqual(['before', 'a', 'b', 'after']);
  });

  it('子の図・線は、グループの外と同じ上限（maxDrawings・maxLines）で数える。グループ自身は数えない', async () => {
    const children = [
      shapeXml({ name: 's1', box: [1000, 2000, 10, 10] }),
      lineShapeXml({ name: 'l1', box: [1000, 2000, 10, 10] }),
      shapeXml({ name: 's2', box: [1000, 2000, 10, 10] }),
      lineShapeXml({ name: 'l2', box: [1000, 2000, 10, 10] }),
      shapeXml({ name: 's3', box: [1000, 2000, 10, 10] }),
    ];
    // 図の上限 2・線の上限 1：グループ自身が図に数えられるなら s1 しか入らない
    const r = await readGroup(grpSpXml('g', children, 6, frameOf(G)), {
      maxDrawings: 2,
      maxLines: 1,
    });
    expect(r.drawings.map((d) => d.name)).toEqual(['s1', 'l1', 's2']);
    expect(r.notices).toEqual(['too-many']);
  });

  it('隠したグループ（cNvPr hidden）は、中の子も出さない。隠した子だけを出さないこともできる', async () => {
    const hiddenGroup = grpSpXml(
      'g',
      [shapeXml({ name: 'c', box: [1000, 2000, 10, 10] })],
      6,
      frameOf(G),
    ).replace('<xdr:cNvPr id="6" name="g"/>', '<xdr:cNvPr id="6" name="g" hidden="1"/>');
    expect(hiddenGroup).toContain('hidden="1"');
    expect((await readGroup(hiddenGroup)).drawings).toEqual([]);
    const { drawings } = await readGroup(
      grpSpXml(
        'g',
        [
          shapeXml({ name: 'hid', hidden: true, box: [1000, 2000, 10, 10] }),
          shapeXml({ name: 'vis', box: [1000, 2000, 10, 10] }),
        ],
        6,
        frameOf(G),
      ),
    );
    expect(drawings.map((d) => d.name)).toEqual(['vis']);
  });

  it('グループの中の mc:AlternateContent（採る枝の図）も展開する', async () => {
    const { drawings } = await readGroup(
      grpSpXml(
        'g',
        [
          alternateContent(
            graphicFrameXml('cx', undefined, 5, [1000, 2000, 10, 10]),
            shapeXml({ name: 'fallback', box: [1000, 2000, 10, 10] }),
          ),
        ],
        6,
        frameOf(G),
      ),
    );
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([['chart', 'cx']]);
    expect(drawings[0]!.group?.xfrm).toEqual(rect([1000, 2000, 10, 10]));
  });

  it('xfrm の無い子は、位置が決まらないので出さない（兄弟は出す）', async () => {
    const { drawings } = await readGroup(
      grpSpXml(
        'g',
        [shapeXml({ name: 'noXfrm' }), shapeXml({ name: 'ok', box: [1000, 2000, 10, 10] })],
        6,
        frameOf(G),
      ),
    );
    expect(drawings.map((d) => d.name)).toEqual(['ok']);
  });
});

describe('TD-89-5 壊れたグループの頭（nvGrpSpPr・grpSpPr・cNvPr が無い）で、子のものを自分のものにしない', () => {
  const inner = grpSpXml(
    'inner',
    [shapeXml({ name: 'deep', box: [10, 20, 30, 40] })],
    7,
    frameOf([2000, 2500, 1000, 800], [0, 0, 100, 80]),
  );

  it('nvGrpSpPr・grpSpPr が無いグループは、子の入れ子のグループの座標系を借りて展開せず、破線枠（名前は空）になる', async () => {
    const { drawings } = await readGroup(`<xdr:grpSp>${inner}</xdr:grpSp>`);
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([['group', '']]);
    expect(drawings[0]!.group).toBeUndefined();
  });

  it('nvGrpSpPr だけあって grpSpPr が直後に無い（子のグループの grpSpPr が先に来る）グループも、展開しない', async () => {
    const { drawings } = await readGroup(
      `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="6" name="g"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>${inner}</xdr:grpSp>`,
    );
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([['group', 'g']]);
  });

  it('cNvPr の無いグループは、隠した子の hidden を自分のものにしない（グループは隠れず、隠した子だけが出ない）', async () => {
    const xml = grpSpXml(
      'g',
      [
        shapeXml({ name: 'hid', hidden: true, box: [1000, 2000, 10, 10] }),
        shapeXml({ name: 'vis', box: [1000, 2000, 10, 10] }),
      ],
      6,
      frameOf(G),
    ).replace('<xdr:cNvPr id="6" name="g"/>', '');
    expect(xml).not.toContain('name="g"');
    const { drawings } = await readGroup(xml);
    expect(drawings.map((d) => d.name)).toEqual(['vis']);
  });

  it('入れ子の中でも同じ：cNvPr の無い内側のグループは、子の hidden で隠れない', async () => {
    const innerNoName = grpSpXml(
      'x',
      [
        shapeXml({ name: 'hid', hidden: true, box: [10, 20, 30, 40] }),
        shapeXml({ name: 'vis', box: [10, 20, 30, 40] }),
      ],
      7,
      frameOf([2000, 2500, 1000, 800], [0, 0, 100, 80]),
    ).replace('<xdr:cNvPr id="7" name="x"/>', '');
    const { drawings } = await readGroup(grpSpXml('outer', [innerNoName], 6, frameOf(G)));
    expect(drawings.map((d) => d.name)).toEqual(['vis']);
  });
});

describe('TD-89-2 座標系：入れ子 2 段・chOff が 0 でない・拡大縮小', () => {
  it('入れ子の子は、外側から順の levels を持つ（最も外側の own は null、内側の own は親の座標系での位置）', async () => {
    const { drawings } = await readGroup(
      grpSpXml(
        'outer',
        [
          grpSpXml(
            'inner',
            [shapeXml({ name: 'deep', box: [10, 20, 30, 40] })],
            7,
            frameOf([2000, 2500, 1000, 800], [0, 0, 100, 80]),
          ),
        ],
        6,
        frameOf(G, [500, 1000, 8000, 6000]),
      ),
    );
    expect(drawings.map((d) => d.name)).toEqual(['deep']);
    expect(drawings[0]!.group).toEqual({
      levels: [
        { ch: rect([500, 1000, 8000, 6000]), own: null },
        { ch: rect([0, 0, 100, 80]), own: rect([2000, 2500, 1000, 800]) },
      ],
      xfrm: rect([10, 20, 30, 40]),
    });
  });

  it('負の座標・巨大な座標も読める（位置は符号付き）', async () => {
    const { drawings } = await readGroup(
      grpSpXml(
        'g',
        [shapeXml({ name: 'neg', box: [-500, -300, 100, 100] })],
        6,
        frameOf(G, [-1000, -1000, 4000, 4000]),
      ),
    );
    expect(drawings[0]!.group?.levels[0]!.ch).toEqual(rect([-1000, -1000, 4000, 4000]));
    expect(drawings[0]!.group?.xfrm).toEqual(rect([-500, -300, 100, 100]));
  });

  it('グループ自身の回転・反転は読まない：levels にも子の rot・flip にも反映されない（子は回転前の xfrm のまま）', async () => {
    const child = shapeXml({ name: 'c', txBox: true, box: [1000, 2000, 10, 10], paras: ['x'] });
    const rotated = await readGroup(
      grpSpXml('g', [child], 6, {
        box: G,
        ch: [500, 600, 8000, 6000],
        attrs: 'rot="1800000" flipH="1" flipV="1"',
      }),
    );
    const plain = await readGroup(
      grpSpXml('g', [child], 6, { box: G, ch: [500, 600, 8000, 6000] }),
    );
    // 子自身の回転は 0（グループの 30° が移らない）。座標系の連なりは回転の有無で変わらず、回転の項目を持たない
    expect(rotated.drawings[0]).toMatchObject({ rot: 0, flipH: false, flipV: false });
    expect(rotated.drawings[0]!.group).toEqual(plain.drawings[0]!.group);
    expect(rotated.drawings[0]!.group).toEqual({
      levels: [{ ch: rect([500, 600, 8000, 6000]), own: null }],
      xfrm: rect([1000, 2000, 10, 10]),
    });
  });

  it('子の回転・反転は ISSUE-88 の読み取りのまま（rot・flipH・flipV）。文字の無い図形・画像は持たない', async () => {
    const { drawings } = await readGroup(
      grpSpXml(
        'g',
        [
          shapeXml({
            name: 'rotTb',
            txBox: true,
            xfrm: 'rot="5400000" flipV="1"',
            box: [1000, 2000, 10, 10],
            paras: ['x'],
          }),
          lineShapeXml({
            name: 'rotLn',
            xfrm: 'rot="16200000" flipH="1"',
            box: [1000, 2000, 10, 10],
          }),
        ],
        6,
        frameOf(G),
      ),
    );
    expect(drawings[0]).toMatchObject({ rot: 5400000, flipH: false, flipV: true });
    expect(drawings[1]).toMatchObject({ rot: 16200000, flipH: true, flipV: false });
  });
});

describe('TD-89-3 写せないグループは、今までどおり破線枠（group 1 件）', () => {
  const child = shapeXml({ name: 'c', box: [1000, 2000, 10, 10] });

  it('chExt の幅か高さが 0：展開しない', async () => {
    for (const ch of [
      [0, 0, 0, 3000],
      [0, 0, 4000, 0],
      [0, 0, 0, 0],
    ] as Box[]) {
      const { drawings } = await readGroup(grpSpXml('g', [child], 6, frameOf(G, ch)));
      expect(
        drawings.map((d) => [d.kind, d.name]),
        JSON.stringify(ch),
      ).toEqual([['group', 'g']]);
      expect(drawings[0]!.group).toBeUndefined();
    }
  });

  it('grpSpPr に xfrm が無い（または chOff・chExt が無い）：展開しない', async () => {
    expect((await readGroup(grpSpXml('g', [child]))).drawings.map((d) => d.kind)).toEqual([
      'group',
    ]);
    const noCh = grpSpXml('g', [child], 6, frameOf(G)).replace(
      /<a:chOff[^>]*\/><a:chExt[^>]*\/>/,
      '',
    );
    expect(noCh).not.toContain('chExt');
    expect((await readGroup(noCh)).drawings.map((d) => d.kind)).toEqual(['group']);
  });

  it('閉じていないグループ：今までどおり破線枠（例外にしない）', async () => {
    const broken = grpSpXml('g', [child], 6, frameOf(G)).replace('</xdr:grpSp>', '');
    const { drawings, notices } = await readGroup(broken);
    expect(notices).toEqual([]);
    expect(drawings.map((d) => d.kind)).toEqual(['group']);
  });

  it('入れ子の中の写せないグループは、親の座標系の位置を持つ破線枠として出る（親の展開は続く）', async () => {
    const { drawings } = await readGroup(
      grpSpXml(
        'outer',
        [
          grpSpXml('zero', [child], 7, frameOf([2000, 2500, 1000, 800], [0, 0, 0, 80])),
          shapeXml({ name: 'sibling', box: [1000, 2000, 10, 10] }),
        ],
        6,
        frameOf(G),
      ),
    );
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([
      ['group', 'zero'],
      ['shape', 'sibling'],
    ]);
    expect(drawings[0]!.group).toEqual({
      levels: [{ ch: rect(G), own: null }],
      xfrm: rect([2000, 2500, 1000, 800]),
    });
  });
});

/** 深さ n のグループ（トップレベルを 1 と数える）。いちばん内側に子 `leaf` を 1 つ置く */
function nested(n: number): string {
  let xml = shapeXml({ name: 'leaf', box: [10, 10, 80, 80] });
  for (let d = n; d >= 1; d--) {
    xml = grpSpXml(
      `g${d}`,
      [xml],
      10 + d,
      frameOf(d === 1 ? G : [0, 0, 100, 100], [0, 0, 100, 100]),
    );
  }
  return xml;
}

describe('TD-89-4 入れ子の深さ：8 まで展開し、9 は破線枠', () => {
  it('深さ 8 のグループの子（leaf）まで展開する。levels は 8 段', async () => {
    const { drawings } = await readGroup(nested(8));
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([['shape', 'leaf']]);
    expect(drawings[0]!.group?.levels).toHaveLength(8);
  });

  it('深さ 9 のグループは展開せず、8 段目までの座標系を持つ破線枠（group）になる。中の leaf は出ない', async () => {
    const { drawings } = await readGroup(nested(9));
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([['group', 'g9']]);
    expect(drawings[0]!.group?.levels).toHaveLength(8);
    expect(drawings[0]!.group?.xfrm).toEqual(rect([0, 0, 100, 100]));
  });

  it('上限は limits で替えられる（maxGroupDepth=2：入れ子 2 段目は展開、3 段目は破線枠）', async () => {
    const two = await readGroup(nested(2), { maxGroupDepth: 2 });
    expect(two.drawings.map((d) => d.name)).toEqual(['leaf']);
    const three = await readGroup(nested(3), { maxGroupDepth: 2 });
    expect(three.drawings.map((d) => [d.kind, d.name])).toEqual([['group', 'g3']]);
  });

  it('深いグループを展開できなくても、兄弟の子は出す', async () => {
    const { drawings } = await readGroup(
      grpSpXml(
        'g1',
        [nested(9).replace('g1', 'gx'), shapeXml({ name: 'sib', box: [1000, 2000, 10, 10] })],
        6,
        frameOf(G),
      ),
    );
    // 外側の g1 が深さ 1、nested(9) が 2〜10 段なので、全体の深さ 9 の g8 が破線枠（g9 は中なので出ない）
    expect(drawings.map((d) => d.name)).toEqual(['g8', 'sib']);
  });
});
