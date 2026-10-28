import { describe, expect, it } from 'vitest';
import { openXlsx, XlsxError, type XlsxLimits } from '../../src/core/index';
import { makePng } from '../../src/testing/make-image';
import {
  absoluteAnchor,
  alternateContent,
  buildXlsx,
  cxnSpXml,
  graphicFrameXml,
  grpSpXml,
  oneCellAnchor,
  picXml,
  spXml,
  twoCellAnchor,
  wsDr,
  type DrawingRelSpec,
  type SheetSpec,
} from '../../src/testing/make-xlsx';

const PNG = makePng(40, 30);
const PIC_RELS: DrawingRelSpec[] = [{ id: 'rId1', target: '../media/image1.png' }];
const A1 = { col: 0, row: 0 };
const B2 = { col: 1, row: 1 };

/** 図の XML（アンカーの並び）を持つ 1 シートのブックを作り、openDrawings の結果を返す */
async function drawingsOf(
  anchors: string[],
  opts: {
    rels?: DrawingRelSpec[];
    xml?: string;
    files?: Record<string, string | Uint8Array>;
    limits?: Partial<XlsxLimits>;
  } = {},
) {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: [['a']],
        drawing: { xml: opts.xml ?? wsDr(anchors), rels: opts.rels ?? PIC_RELS },
      },
    ],
    extraFiles: { 'xl/media/image1.png': PNG, ...opts.files },
  });
  const book = await openXlsx(bytes, { limits: opts.limits });
  return book.openDrawings(1);
}

describe('TD-85-3 openDrawings：図の無いシート・入口の引数', () => {
  it('drawing の無いシートは空の一覧（通知なし）', async () => {
    const book = await openXlsx(await buildXlsx({ sheets: [{ name: 'S', rows: [['a']] }] }));
    expect(await book.openDrawings(1)).toEqual({ drawings: [], notices: [] });
  });

  it('グラフシート（chartsheet）は空の一覧', async () => {
    const book = await openXlsx(
      await buildXlsx({
        sheets: [
          { name: 'S', rows: [['a']] },
          { name: 'C', chartsheet: true },
        ],
      }),
    );
    expect(await book.openDrawings(2)).toEqual({ drawings: [], notices: [] });
  });

  it('シートの番号が範囲外なら RangeError、閉じた後は XlsxError(broken)', async () => {
    const book = await openXlsx(await buildXlsx({ sheets: [{ name: 'S', rows: [['a']] }] }));
    await expect(book.openDrawings(0)).rejects.toBeInstanceOf(RangeError);
    await expect(book.openDrawings(2)).rejects.toBeInstanceOf(RangeError);
    book.close();
    await expect(book.openDrawings(1)).rejects.toMatchObject({ kind: 'broken' });
    await expect(book.openDrawings(1)).rejects.toBeInstanceOf(XlsxError);
  });

  it('openSheet の前でも後でも同じ結果（シートの XML から読んだ drawing の指定を覚える）', async () => {
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: { xml: wsDr([twoCellAnchor(A1, B2, spXml('r'))]) },
        },
      ],
    });
    const fresh = await (await openXlsx(bytes)).openDrawings(1);
    const book = await openXlsx(bytes);
    await book.openSheet(1);
    expect(await book.openDrawings(1)).toEqual(fresh);
    expect(fresh.drawings).toHaveLength(1);
  });

  it('シートの rels が指す drawing が zip に無いときは空の一覧（捨てる）', async () => {
    const bytes = await buildXlsx({
      sheets: [{ name: 'S', rows: [['a']], drawing: { xml: wsDr([]), omitPart: true } }],
    });
    expect(await (await openXlsx(bytes)).openDrawings(1)).toEqual({ drawings: [], notices: [] });
  });

  it('ほかの要素に drawing という語が入っていても、<drawing> だけを拾う（legacyDrawing・セルの文字）', async () => {
    const spec: SheetSpec = {
      name: 'S',
      xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><legacyDrawing r:id="rId9"/></worksheet>',
      drawing: { xml: wsDr([twoCellAnchor(A1, B2, spXml('r'))]) },
    };
    const book = await openXlsx(await buildXlsx({ sheets: [spec] }));
    expect((await book.openDrawings(1)).drawings).toEqual([]);
  });
});

describe('TD-85-4 openDrawings：3 種のアンカー', () => {
  it('twoCellAnchor：from・to の番号とずれ。editAs は無視する', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(
        { col: 2, colOff: 38100, row: 3, rowOff: 25400 },
        { col: 4, colOff: 358140, row: 5, rowOff: 128270 },
        picXml({ name: 'p', embed: 'rId1' }),
        'oneCell',
      ),
      twoCellAnchor(A1, B2, spXml('q'), 'absolute'),
    ]);
    expect(drawings.map((d) => d.anchor)).toEqual([
      {
        type: 'twoCell',
        from: { col: 2, colOff: 38100, row: 3, rowOff: 25400 },
        to: { col: 4, colOff: 358140, row: 5, rowOff: 128270 },
      },
      {
        type: 'twoCell',
        from: { col: 0, colOff: 0, row: 0, rowOff: 0 },
        to: { col: 1, colOff: 0, row: 1, rowOff: 0 },
      },
    ]);
  });

  it('oneCellAnchor：from と ext（cx・cy の EMU）', async () => {
    const { drawings } = await drawingsOf([
      oneCellAnchor({ col: 5, colOff: 100, row: 7, rowOff: 200 }, 952500, 476250, spXml('s')),
    ]);
    expect(drawings).toHaveLength(1);
    expect(drawings[0]!.anchor).toEqual({
      type: 'oneCell',
      from: { col: 5, colOff: 100, row: 7, rowOff: 200 },
      cx: 952500,
      cy: 476250,
    });
  });

  it('absoluteAnchor：pos（x・y の EMU）と ext', async () => {
    const { drawings } = await drawingsOf([
      absoluteAnchor(609600, 182880, 381000, 285750, spXml('s')),
    ]);
    expect(drawings[0]!.anchor).toEqual({
      type: 'absolute',
      x: 609600,
      y: 182880,
      cx: 381000,
      cy: 285750,
    });
  });

  it('接頭辞が xdr: でない drawing（既定の名前空間・別の接頭辞）も読む', async () => {
    const xml =
      '<?xml version="1.0"?><wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><twoCellAnchor><from><col>1</col><colOff>2</colOff><row>3</row><rowOff>4</rowOff></from><to><col>5</col><colOff>6</colOff><row>7</row><rowOff>8</rowOff></to><sp><nvSpPr><cNvPr id="2" name="n"/><cNvSpPr/></nvSpPr></sp><clientData/></twoCellAnchor></wsDr>';
    const { drawings } = await drawingsOf([], { xml });
    expect(drawings).toHaveLength(1);
    expect(drawings[0]).toMatchObject({
      kind: 'shape',
      name: 'n',
      anchor: {
        type: 'twoCell',
        from: { col: 1, colOff: 2, row: 3, rowOff: 4 },
        to: { col: 5, colOff: 6, row: 7, rowOff: 8 },
      },
    });
  });

  it('不正なアンカー（from が無い・値が数字でない・列が範囲外）の図は一覧に出さず、後ろの図は読む', async () => {
    const good = twoCellAnchor(A1, B2, spXml('good'));
    const noFrom =
      '<xdr:twoCellAnchor><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
      spXml('x') +
      '</xdr:twoCellAnchor>';
    const nan = twoCellAnchor({ col: Number.NaN, row: 0 }, B2, spXml('y'));
    const farCol = twoCellAnchor({ col: 16384, row: 0 }, B2, spXml('z'));
    const { drawings, notices } = await drawingsOf([noFrom, nan, farCol, good]);
    expect(drawings.map((d) => d.name)).toEqual(['good']);
    expect(notices).toEqual([]);
  });

  it('負のずれ・巨大なずれは範囲に丸める（0 以上・2^31-1 以下）', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(
        { col: 0, colOff: -5, row: 0, rowOff: 99999999999 },
        { col: 1, colOff: 0, row: 1, rowOff: 0 },
        spXml('s'),
      ),
    ]);
    expect(drawings[0]!.anchor).toMatchObject({
      type: 'twoCell',
      from: { colOff: 0, rowOff: 2147483647 },
    });
  });
});

describe('TD-85-5 openDrawings：要素の種類（ADR-0026 §決定 6 の表）', () => {
  it('画像・グラフ・SmartArt・図形・テキストボックス・グループを判別し、名前と説明を持つ。順序は drawing XML の順', async () => {
    const { drawings, notices } = await drawingsOf([
      twoCellAnchor(A1, B2, picXml({ name: 'pic', descr: '画面写し', embed: 'rId1' })),
      twoCellAnchor(A1, B2, graphicFrameXml('chart')),
      twoCellAnchor(
        A1,
        B2,
        graphicFrameXml('dgm', 'http://schemas.openxmlformats.org/drawingml/2006/diagram'),
      ),
      twoCellAnchor(A1, B2, spXml('rect')),
      twoCellAnchor(A1, B2, spXml('note', true)),
      twoCellAnchor(
        A1,
        B2,
        grpSpXml('grp', [spXml('inner'), picXml({ name: 'inPic', embed: 'rId1' })]),
      ),
    ]);
    expect(notices).toEqual([]);
    expect(drawings.map((d) => [d.kind, d.name, d.descr, d.mediaPath])).toEqual([
      ['image', 'pic', '画面写し', 'xl/media/image1.png'],
      ['chart', 'chart', '', null],
      ['smartart', 'dgm', '', null],
      ['shape', 'rect', '', null],
      ['textbox', 'note', '', null],
      ['group', 'grp', '', null],
    ]);
  });

  it('名前・説明の実体参照を戻し、長すぎる文字は 1,000 文字で切る', async () => {
    const long = 'あ'.repeat(5000);
    const { drawings } = await drawingsOf([
      twoCellAnchor(A1, B2, picXml({ name: 'a&b<c', descr: long, embed: 'rId1' })),
    ]);
    expect(drawings[0]!.name).toBe('a&b<c');
    expect(drawings[0]!.descr).toHaveLength(1000);
  });

  it('ISSUE-89：座標系（grpSpPr の a:xfrm）の無いグループは展開できず、group 1 件のまま（中の画像は出ない）', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(A1, B2, grpSpXml('g', [picXml({ name: 'inPic', embed: 'rId1' })])),
    ]);
    expect(drawings.map((d) => [d.kind, d.mediaPath])).toEqual([['group', null]]);
  });

  it('ISSUE-89：座標系のあるグループは展開され、グループの中の画像が出る（グループ自身は出ない）', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(
        A1,
        B2,
        grpSpXml('g', [picXml({ name: 'inPic', embed: 'rId1' })], 6, { box: [0, 0, 100, 100] }),
      ),
    ]);
    expect(drawings.map((d) => [d.kind, d.name, d.mediaPath])).toEqual([
      ['image', 'inPic', 'xl/media/image1.png'],
    ]);
  });

  it('隠した図（cNvPr hidden="1"）は出さない', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(A1, B2, picXml({ name: 'h', hidden: true, embed: 'rId1' })),
      twoCellAnchor(A1, B2, spXml('v')),
    ]);
    expect(drawings.map((d) => d.name)).toEqual(['v']);
  });

  it('トリミング（srcRect）・回転（rot）の指定があっても読まず、画像として同じように出す', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(A1, B2, picXml({ name: 'plain', embed: 'rId1' })),
      twoCellAnchor(
        A1,
        B2,
        picXml({
          name: 'rot',
          embed: 'rId1',
          srcRect: 'l="16667"',
          xfrmAttrs: 'rot="1800000" flipH="1"',
        }),
      ),
    ]);
    const [a, b] = drawings;
    expect(b).toEqual({ ...a, name: 'rot' });
  });
});

describe('TD-85-6 openDrawings：線・矢印（cxnSp）は line で出し、図とは別に数える（ISSUE-87 で期待を改めた）', () => {
  it('cxnSp は種類 line で一覧に出る（図と同じ一覧に XML の順で並ぶ）', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(A1, B2, cxnSpXml('line')),
      twoCellAnchor(A1, B2, spXml('s')),
      oneCellAnchor(A1, 100, 100, cxnSpXml('arrow')),
    ]);
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([
      ['line', 'line'],
      ['shape', 's'],
      ['line', 'arrow'],
    ]);
  });

  it('図の数の上限に数えない：上限ちょうどの図と大量の線があっても、通知は出ない（線は別の上限で数える）', async () => {
    const lines = Array.from({ length: 50 }, (_v, i) => twoCellAnchor(A1, B2, cxnSpXml(`l${i}`)));
    const shapes = Array.from({ length: 3 }, (_v, i) => twoCellAnchor(A1, B2, spXml(`s${i}`)));
    const { drawings, notices } = await drawingsOf([...lines, ...shapes], {
      limits: { maxDrawings: 3 },
    });
    expect(drawings.filter((d) => d.kind !== 'line')).toHaveLength(3);
    expect(drawings.filter((d) => d.kind === 'line')).toHaveLength(50);
    expect(notices).toEqual([]);
  });
});

describe('TD-85-7 openDrawings：mc:AlternateContent', () => {
  it('アンカーの中：xdr:pic を含む枝（Choice でも Fallback でも）を採る', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(A1, B2, alternateContent(picXml({ name: 'cp', embed: 'rId1' }), spXml('fs'))),
      twoCellAnchor(A1, B2, alternateContent(spXml('cs'), picXml({ name: 'fp', embed: 'rId1' }))),
    ]);
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([
      ['image', 'cp'],
      ['image', 'fp'],
    ]);
  });

  it('アンカーの中：xdr:pic が無ければ、graphicFrame を持つ Choice（uri を問わない）。無ければ Fallback の種類で出す', async () => {
    const { drawings } = await drawingsOf([
      twoCellAnchor(
        A1,
        B2,
        alternateContent(graphicFrameXml('chart-choice'), spXml('fallback-shape')),
      ),
      twoCellAnchor(A1, B2, alternateContent(spXml('shape-choice'), spXml('fallback-shape'))),
    ]);
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([
      ['chart', 'chart-choice'],
      ['shape', 'fallback-shape'],
    ]);
  });

  it('アンカーの外（スライサーなど）：Choice と Fallback の両方にアンカーがあっても 1 つだけ出す', async () => {
    const choice = twoCellAnchor(A1, B2, graphicFrameXml('slicer'));
    const fallback = twoCellAnchor({ col: 3, row: 3 }, { col: 4, row: 4 }, spXml('fallback-shape'));
    const { drawings } = await drawingsOf([
      alternateContent(choice, fallback),
      twoCellAnchor(A1, B2, spXml('after')),
    ]);
    expect(drawings.map((d) => d.name)).toEqual(['fallback-shape', 'after']);
    expect(drawings[0]!.anchor).toMatchObject({ from: { col: 3, row: 3 } });
  });

  it('アンカーの外：pic を含むアンカーを持つ枝があればそちら', async () => {
    const choice = twoCellAnchor(A1, B2, picXml({ name: 'cpic', embed: 'rId1' }));
    const fallback = twoCellAnchor(A1, B2, spXml('fshape'));
    const { drawings } = await drawingsOf([alternateContent(choice, fallback)]);
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([['image', 'cpic']]);
  });
});

describe('TD-85-16 openDrawings：mc:AlternateContent の入れ子', () => {
  it('アンカーの外の AlternateContent の Choice の中に、アンカー内の AlternateContent（pic）があっても、画像 1 件が出る', async () => {
    const inner = twoCellAnchor(
      A1,
      B2,
      alternateContent(picXml({ name: 'innerPic', embed: 'rId1' }), spXml('innerFb')),
    );
    const outerFb = twoCellAnchor(A1, B2, spXml('outerFb'));
    const { drawings } = await drawingsOf([
      alternateContent(inner, outerFb),
      twoCellAnchor(A1, B2, spXml('after')),
    ]);
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([
      ['image', 'innerPic'],
      ['shape', 'after'],
    ]);
  });

  it('Fallback は AlternateContent 直下のものだけを採る（Choice の中の入れ子の Fallback を、外側の Fallback と取り違えない）', async () => {
    // 外側の Choice（pic なし）の中の入れ子の Fallback は 'nestedFb'。外側の Fallback は 'outerFb'
    const nested = alternateContent(graphicFrameXml('nestedChoice'), spXml('nestedFb'));
    const choice = twoCellAnchor(A1, B2, nested);
    const outerFb = twoCellAnchor(A1, B2, spXml('outerFb'));
    const { drawings } = await drawingsOf([alternateContent(choice, outerFb)]);
    expect(drawings.map((d) => d.name)).toEqual(['outerFb']);
  });

  it('入れ子の AlternateContent の Choice・Fallback の閉じを、外側の枝の終わりと取り違えない', async () => {
    const inner = alternateContent(spXml('i1'), picXml({ name: 'i2', embed: 'rId1' }));
    const { drawings } = await drawingsOf([
      twoCellAnchor(A1, B2, alternateContent(inner, spXml('outerFb'))),
    ]);
    expect(drawings.map((d) => [d.kind, d.name])).toEqual([['image', 'i2']]);
  });
});

describe('TD-85-8 openDrawings：画像の参照（media のパス・svgBlip・外部リンク）', () => {
  it('asvg:svgBlip の併記は SVG を使わず、a:blip の r:embed（PNG の予備）のパスを返す', async () => {
    const { drawings } = await drawingsOf(
      [twoCellAnchor(A1, B2, picXml({ name: 'svg', embed: 'rId1', svgEmbed: 'rId2' }))],
      {
        rels: [
          { id: 'rId1', target: '../media/image1.png' },
          { id: 'rId2', target: '../media/image2.svg' },
        ],
        files: { 'xl/media/image2.svg': '<svg xmlns="http://www.w3.org/2000/svg"/>' },
      },
    );
    expect(drawings[0]).toMatchObject({ kind: 'image', mediaPath: 'xl/media/image1.png' });
  });

  it('SVG だけ（a:blip に r:embed が無く svgBlip だけ）の画像は mediaPath が null（SVG は使わない）', async () => {
    const { drawings } = await drawingsOf(
      [twoCellAnchor(A1, B2, picXml({ name: 'svgonly', svgEmbed: 'rId2' }))],
      {
        rels: [{ id: 'rId2', target: '../media/image2.svg' }],
        files: { 'xl/media/image2.svg': '<svg xmlns="http://www.w3.org/2000/svg"/>' },
      },
    );
    expect(drawings[0]).toMatchObject({ kind: 'image', mediaPath: null });
  });

  it('外部リンク（TargetMode="External"）・r:link だけ・zip に無い Target・r:embed 無しは、辿らず mediaPath null（図は残る）', async () => {
    const { drawings, notices } = await drawingsOf(
      [
        twoCellAnchor(A1, B2, picXml({ name: 'ext', embed: 'rId5' })),
        twoCellAnchor(A1, B2, picXml({ name: 'link', link: 'rId5' })),
        twoCellAnchor(A1, B2, picXml({ name: 'gone', embed: 'rId6' })),
        twoCellAnchor(A1, B2, picXml({ name: 'noembed' })),
        twoCellAnchor(A1, B2, picXml({ name: 'norel', embed: 'rId99' })),
        twoCellAnchor(A1, B2, picXml({ name: 'ok', embed: 'rId1' })),
      ],
      {
        rels: [
          { id: 'rId1', target: '../media/image1.png' },
          { id: 'rId5', target: 'https://example.com/a.png', external: true },
          { id: 'rId6', target: '../media/missing.png' },
        ],
      },
    );
    expect(notices).toEqual([]);
    expect(drawings.map((d) => [d.name, d.mediaPath])).toEqual([
      ['ext', null],
      ['link', null],
      ['gone', null],
      ['noembed', null],
      ['norel', null],
      ['ok', 'xl/media/image1.png'],
    ]);
  });

  it('外部指定の無い file:/// の Target は、zip の相対パスとして解決した先が無ければ捨てる（外部を取りに行かない）', async () => {
    const { drawings } = await drawingsOf(
      [twoCellAnchor(A1, B2, picXml({ name: 'f', embed: 'rId5' }))],
      {
        rels: [{ id: 'rId5', target: 'file:///C:/secret/a.png' }],
      },
    );
    // 外部指定が無くても、相対パスとして解決した先が zip に無ければ捨てる
    expect(drawings[0]!.mediaPath).toBe(null);
  });

  it('Target の正規化：drawing の位置からの相対（..）と、/ 始まりの絶対パス。重複する / ・. も解く', async () => {
    const { drawings } = await drawingsOf(
      [
        twoCellAnchor(A1, B2, picXml({ name: 'rel', embed: 'rId1' })),
        twoCellAnchor(A1, B2, picXml({ name: 'abs', embed: 'rId2' })),
        twoCellAnchor(A1, B2, picXml({ name: 'dots', embed: 'rId3' })),
      ],
      {
        rels: [
          { id: 'rId1', target: '../media/image1.png' },
          { id: 'rId2', target: '/xl/media/image1.png' },
          { id: 'rId3', target: './../drawings/../media//image1.png' },
        ],
      },
    );
    expect(drawings.map((d) => d.mediaPath)).toEqual([
      'xl/media/image1.png',
      'xl/media/image1.png',
      'xl/media/image1.png',
    ]);
  });

  it('drawing の rels が無くても読める（図は出し、画像は null）', async () => {
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: { xml: wsDr([twoCellAnchor(A1, B2, picXml({ name: 'p', embed: 'rId1' }))]) },
        },
      ],
    });
    const { drawings } = await (await openXlsx(bytes)).openDrawings(1);
    expect(drawings.map((d) => [d.kind, d.mediaPath])).toEqual([['image', null]]);
  });

  it('拡張子・Content_Types を信じない：拡張子が何でも media のパスをそのまま返す（形式は readImage が先頭バイトで決める）', async () => {
    const { drawings } = await drawingsOf(
      [twoCellAnchor(A1, B2, picXml({ name: 'p', embed: 'rId1' }))],
      {
        rels: [{ id: 'rId1', target: '../media/image1.emf' }],
        files: { 'xl/media/image1.emf': PNG },
      },
    );
    expect(drawings[0]!.mediaPath).toBe('xl/media/image1.emf');
  });
});

describe('TD-85-9 openDrawings：上限・解析失敗は XlsxError にせず、通知で返す', () => {
  it('図の数の上限（maxDrawings）：先頭の分だけ出し、too-many を通知', async () => {
    const shapes = Array.from({ length: 5 }, (_v, i) => twoCellAnchor(A1, B2, spXml(`s${i}`)));
    const { drawings, notices } = await drawingsOf(shapes, { limits: { maxDrawings: 3 } });
    expect(drawings.map((d) => d.name)).toEqual(['s0', 's1', 's2']);
    expect(notices).toEqual(['too-many']);
  });

  it('図の数の既定の上限は 1,000：1,001 個目から出さない', async () => {
    const shapes = Array.from({ length: 1001 }, (_v, i) => twoCellAnchor(A1, B2, spXml(`s${i}`)));
    const { drawings, notices } = await drawingsOf(shapes);
    expect(drawings).toHaveLength(1000);
    expect(notices).toEqual(['too-many']);
  });

  it('drawing XML が展開後の上限（maxSmallPartBytes）を超えたら図を出さず part-too-large。表は開ける', async () => {
    const many = Array.from({ length: 200 }, (_v, i) => twoCellAnchor(A1, B2, spXml(`s${i}`)));
    const big = await buildXlsx({
      sheets: [{ name: 'S', rows: [['cell']], drawing: { xml: wsDr(many) } }],
    });
    const limit = 4000; // workbook.xml・styles.xml・rels より大きく、drawing より小さい
    const b2 = await openXlsx(big, { limits: { maxSmallPartBytes: limit } });
    expect(await b2.openDrawings(1)).toEqual({ drawings: [], notices: ['part-too-large'] });
    const sheet = await b2.openSheet(1);
    expect(sheet.getRows(1, 1)[0]!.cells[0]).toMatchObject({ text: 'cell' });
  });

  it('drawing の rels が上限を超えたら図を出さず part-too-large', async () => {
    const rels = Array.from({ length: 300 }, (_v, i) => ({
      id: `rId${i + 10}`,
      target: `../media/image${i + 10}.png`,
    }));
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: { xml: wsDr([twoCellAnchor(A1, B2, spXml('s'))]), rels },
        },
      ],
    });
    const book = await openXlsx(bytes, { limits: { maxSmallPartBytes: 4000 } });
    expect(await book.openDrawings(1)).toEqual({ drawings: [], notices: ['part-too-large'] });
  });

  it('drawing の部品の展開に失敗（壊れた圧縮データ）したら図を出さず broken。例外にしない', async () => {
    const good = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: {
            xml: wsDr(Array.from({ length: 30 }, () => twoCellAnchor(A1, B2, spXml('s')))),
          },
        },
      ],
    });
    const bytes = corruptEntry(good, 'xl/drawings/drawing1.xml');
    const book = await openXlsx(bytes);
    expect(await book.openDrawings(1)).toEqual({ drawings: [], notices: ['broken'] });
    // 表は読める
    expect((await book.openSheet(1)).rowCount).toBe(1);
  });

  it('閉じていない・壊れた XML でも例外を出さず、読めた分だけ返す', async () => {
    const xml = wsDr([twoCellAnchor(A1, B2, spXml('ok'))]).replace(
      '</xdr:wsDr>',
      '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col></xdr:wsDr>',
    );
    const { drawings } = await drawingsOf([], { xml });
    expect(drawings.map((d) => d.name)).toEqual(['ok']);
  });

  it('閉じの無い要素が大量にあっても二乗時間にならない（線形の切り出し）', async () => {
    const opens = '<xdr:twoCellAnchor>'.repeat(20000);
    const t0 = performance.now();
    const { drawings } = await drawingsOf([], { xml: wsDr([opens]) });
    expect(drawings).toEqual([]);
    expect(performance.now() - t0).toBeLessThan(3000);
  });

  it('図の名前・画像のパスを console に出さない', async () => {
    const calls: unknown[][] = [];
    const orig = { log: console.log, warn: console.warn, error: console.error, info: console.info };
    console.log =
      console.warn =
      console.error =
      console.info =
        (...a: unknown[]) => {
          calls.push(a);
        };
    try {
      await drawingsOf([twoCellAnchor(A1, B2, picXml({ name: 'SECRET-NAME', embed: 'rId1' }))]);
      await drawingsOf([twoCellAnchor(A1, B2, picXml({ name: 'x', embed: 'rId9' }))]);
    } finally {
      Object.assign(console, orig);
    }
    expect(calls).toEqual([]);
  });
});

/**
 * zip の中の 1 ファイルの圧縮データの先頭を壊す（deflate の不正なブロックにする）。
 * 中央ディレクトリは無傷なので zip としては開け、そのファイルの展開だけが失敗する。
 */
function corruptEntry(zip: Uint8Array, name: string): Uint8Array {
  const out = zip.slice();
  const dv = new DataView(out.buffer);
  const nameBytes = new TextEncoder().encode(name);
  for (let i = 0; i + 30 < out.length; i++) {
    if (dv.getUint32(i, true) !== 0x04034b50) continue;
    const nameLen = dv.getUint16(i + 26, true);
    const extraLen = dv.getUint16(i + 28, true);
    const got = out.subarray(i + 30, i + 30 + nameLen);
    if (got.length === nameBytes.length && got.every((v, k) => v === nameBytes[k])) {
      out.fill(0xff, i + 30 + nameLen + extraLen, i + 30 + nameLen + extraLen + 8);
      return out;
    }
  }
  throw new Error('entry not found');
}
