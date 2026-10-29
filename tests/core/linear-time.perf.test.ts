import { describe, expect, it } from 'vitest';
import { findDrawingRelId } from '../../src/core/drawing';
import { openXlsx, type XlsxLimits } from '../../src/core/index';
import { buildXlsx, wsDr } from '../../src/testing/make-xlsx';
import type { SheetDrawings } from '../../src/core/types';
import { findElement, parseAttrs } from '../../src/core/xml';

/**
 * 細工したファイル（`>` の無い開きタグ・閉じの無い要素の大量の繰り返し）で、読み取りの時間が大きさに比例することの確認。
 * `[^>]*>` は後ろに `>` が無いと候補ごとに末尾まで読んで二乗時間になる（371KiB で 5.9 秒だった）ので、`[^<>]*>` で打ち切る。
 * 線形なら本件の大きさ（約 400KB〜1MB）は数十 ms。二乗なら秒単位になるので、緩い固定の時間（1.5 秒）で判定する。
 * 時間を測るため、perf の project（ほかのテストの後に 1 ファイルずつ）で走らせる（*.perf.test.ts）。
 */
const LIMIT_MS = 1500;
/** 繰り返しの数（二乗なら 1 秒を大きく超える大きさ） */
const N = 40_000;

/**
 * fn を実行して、かかった時間（ms）を返す。上限を超えたときだけ 3 回まで測り直し、いちばん短い時間を返す。
 * ほかの処理と重なった 1 回の遅れで落ちないため。二乗なら毎回秒単位になるので、測り直しても上限を超える。
 */
async function elapsed(fn: () => unknown): Promise<number> {
  let best = Infinity;
  for (let i = 0; i < 3 && best >= LIMIT_MS; i++) {
    const t0 = performance.now();
    await fn();
    best = Math.min(best, performance.now() - t0);
  }
  return best;
}

const MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

describe('TD-85-17 細工した XML でも時間が線形（> の無い開きタグ・閉じの無い要素）', () => {
  it('drawing：アンカーの開きタグに > が無い（トップレベル）', async () => {
    const xml = wsDr(['<xdr:twoCellAnchor '.repeat(N)]);
    const bytes = await buildXlsx({
      sheets: [{ name: 'S', rows: [['a']], drawing: { xml, rels: [] } }],
    });
    const book = await openXlsx(bytes);
    expect(await elapsed(() => book.openDrawings(1))).toBeLessThan(LIMIT_MS);
  });

  it('drawing：アンカーの中の図の要素・cNvPr・blip の開きタグに > が無い', async () => {
    const pt =
      '<xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:to>';
    const anchors = [
      `<xdr:twoCellAnchor>${pt}${'<xdr:sp '.repeat(N)}</xdr:twoCellAnchor>`,
      `<xdr:twoCellAnchor>${pt}<xdr:pic>${'<xdr:cNvPr '.repeat(N)}${'<a:blip '.repeat(N)}</xdr:pic></xdr:twoCellAnchor>`,
      `<xdr:twoCellAnchor>${pt}<xdr:graphicFrame>${'<a:graphicData '.repeat(N)}</xdr:graphicFrame></xdr:twoCellAnchor>`,
    ];
    const bytes = await buildXlsx({
      sheets: [{ name: 'S', rows: [['a']], drawing: { xml: wsDr(anchors), rels: [] } }],
    });
    const book = await openXlsx(bytes);
    expect(await elapsed(() => book.openDrawings(1))).toBeLessThan(LIMIT_MS);
  });

  it('drawing：mc:AlternateContent の開きタグに > が無い・閉じが無い・入れ子が深い', async () => {
    const ac = '<mc:AlternateContent>';
    const xmls = [
      wsDr(['<mc:AlternateContent '.repeat(N)]),
      wsDr([ac.repeat(N)]),
      wsDr([`${ac}<mc:Choice>`.repeat(N)]),
      wsDr([`<xdr:twoCellAnchor><xdr:from/>${ac.repeat(N)}</xdr:twoCellAnchor>`]),
    ];
    for (const xml of xmls) {
      const bytes = await buildXlsx({
        sheets: [{ name: 'S', rows: [['a']], drawing: { xml, rels: [] } }],
      });
      const book = await openXlsx(bytes);
      expect(await elapsed(() => book.openDrawings(1))).toBeLessThan(LIMIT_MS);
    }
  });

  it('drawing の rels：<Relationship に > が無い', async () => {
    const bytes = await buildXlsx({
      sheets: [{ name: 'S', rows: [['a']], drawing: { xml: wsDr([]), rels: [] } }],
      extraFiles: { 'xl/drawings/_rels/drawing1.xml.rels': '<Relationship '.repeat(N) },
    });
    const book = await openXlsx(bytes);
    expect(await elapsed(() => book.openDrawings(1))).toBeLessThan(LIMIT_MS);
  });

  it('findDrawingRelId（openSheet が毎回呼ぶ）：末尾に <drawing を > 無しで並べる／閉じの無い <drawing>', async () => {
    const head = `<worksheet ${MAIN}><sheetData></sheetData>`;
    expect(await elapsed(() => findDrawingRelId(head + '<drawing '.repeat(N)))).toBeLessThan(
      LIMIT_MS,
    );
    expect(await elapsed(() => findDrawingRelId(head + '<drawing>'.repeat(N)))).toBeLessThan(
      LIMIT_MS,
    );
    // 正しい <drawing> は、細工が前にあっても拾える
    expect(findDrawingRelId(`${head}<drawing r:id="rId7"/>`)).toBe('rId7');
  });

  it('findElement（xml.ts）：> の無い開きタグ・閉じの無い要素の繰り返し', async () => {
    expect(await elapsed(() => findElement('<a '.repeat(N), 'a'))).toBeLessThan(LIMIT_MS);
    expect(await elapsed(() => findElement('<a>'.repeat(N), 'a'))).toBeLessThan(LIMIT_MS);
  });

  it('parseAttrs（xml.ts）：長い属性名（= が無い）・多数の属性・閉じていない引用符でも線形', async () => {
    // 名前の途中から照合し直すと二乗時間になる（名前 4 万字で数秒）。先頭の後読みで、名前の始まりだけから照合する
    const long = 'a'.repeat(N);
    expect(await elapsed(() => parseAttrs('<a:rPr ' + long))).toBeLessThan(LIMIT_MS);
    expect(
      await elapsed(() => parseAttrs('<a:rPr ' + long + ' ' + long + '=1 ' + long)),
    ).toBeLessThan(LIMIT_MS);
    expect(await elapsed(() => parseAttrs('<a:rPr ' + 'a:b.c-d'.repeat(N / 7)))).toBeLessThan(
      LIMIT_MS,
    );
    expect(await elapsed(() => parseAttrs('<a:rPr ' + 'a  '.repeat(N)))).toBeLessThan(LIMIT_MS);
    expect(await elapsed(() => parseAttrs('<a:rPr ' + 'a="x" '.repeat(N)))).toBeLessThan(LIMIT_MS);
    expect(await elapsed(() => parseAttrs('<a:rPr a="' + 'b '.repeat(N)))).toBeLessThan(LIMIT_MS);
    // 正しい属性は、細工が前にあっても拾える
    expect(parseAttrs('<a:rPr ' + long + ' sz="1200" b="1"/>')).toMatchObject({
      sz: '1200',
      b: '1',
    });
  });

  it('sheet：結合セル <mergeCell に > と ref が無い／<sheetData に > が無い', async () => {
    const merge = `<worksheet ${MAIN}><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row></sheetData><mergeCells>${'<mergeCell '.repeat(N)}</mergeCells></worksheet>`;
    const sd = `<worksheet ${MAIN}>${'<sheetData '.repeat(N)}</worksheet>`;
    for (const xml of [merge, sd]) {
      const book = await openXlsx(await buildXlsx({ sheets: [{ name: 'S', xml }] }));
      // 開けても開けなくてもよい。固まらず、時間が線形であること
      const ms = await elapsed(() => book.openSheet(1).catch(() => undefined));
      expect(ms).toBeLessThan(LIMIT_MS);
    }
  });

  it('styles：<cellXfs・<xf・<numFmt・<font に > が無い／閉じの無い <cellXfs>', async () => {
    const styles = [
      `<styleSheet ${MAIN}>${'<cellXfs '.repeat(N)}</styleSheet>`,
      `<styleSheet ${MAIN}><cellXfs>${'<xf '.repeat(N)}</cellXfs></styleSheet>`,
      `<styleSheet ${MAIN}><cellXfs>${'<xf>'.repeat(N)}</styleSheet>`,
      `<styleSheet ${MAIN}>${'<numFmt '.repeat(N)}<fonts><font>${'<sz '.repeat(N)}${'<name '.repeat(N)}</font></fonts></styleSheet>`,
    ];
    for (const xml of styles) {
      const bytes = await buildXlsx({
        sheets: [{ name: 'S', rows: [['a']] }],
        extraFiles: { 'xl/styles.xml': xml },
      });
      const ms = await elapsed(async () => {
        const book = await openXlsx(bytes);
        await book.openSheet(1);
      });
      expect(ms).toBeLessThan(LIMIT_MS);
    }
  });

  it('workbook・rels：<sheet・<Relationship に > が無い', async () => {
    const wb = `<workbook ${MAIN}><sheets>${'<sheet '.repeat(N)}</sheets></workbook>`;
    const bytes = await buildXlsx({
      sheets: [{ name: 'S', rows: [['a']] }],
      extraFiles: {
        'xl/workbook.xml': wb,
        'xl/_rels/workbook.xml.rels': '<Relationship '.repeat(N),
      },
    });
    const ms = await elapsed(() => openXlsx(bytes).catch(() => undefined));
    expect(ms).toBeLessThan(LIMIT_MS);
  });

  /** sheetData の中身（と共有文字列の XML）を渡して、開く・全行の展開・共有文字列の取り出しにかかる時間を測る */
  async function readTime(rows: string, sstXml?: string): Promise<number> {
    const bytes = await buildXlsx({
      sheets: [{ name: 'S', sheetDataXml: rows }],
      sharedStrings: sstXml === undefined ? undefined : [],
      extraFiles: sstXml === undefined ? {} : { 'xl/sharedStrings.xml': sstXml },
    });
    return elapsed(async () => {
      const book = await openXlsx(bytes);
      const sheet = await book.openSheet(1);
      sheet.getRows(1, Math.max(1, sheet.rowCount));
      sheet.getRows(1, 2);
    });
  }

  it('sheet のセル：<c に > が無い／閉じの無い <c>／セルの中の <v・<is に > が無い・閉じが無い', async () => {
    const rows = [
      `<row r="1">${'<c '.repeat(N)}<c r="A1"><v>1</v></c></row>`,
      `<row r="1">${'<c>'.repeat(N)}</row>`,
      `<row r="1"><c r="A1">${'<v '.repeat(N)}</c></row>`,
      `<row r="1"><c r="A1">${'<v>'.repeat(N)}</c></row>`,
      `<row r="1"><c r="A1" t="inlineStr">${'<is '.repeat(N)}</c></row>`,
      `<row r="1"><c r="A1" t="inlineStr">${'<is>'.repeat(N)}</c></row>`,
      `<row r="1"><c r="A1" t="inlineStr"><is>${'<t '.repeat(N)}${'<t>'.repeat(N)}</is></c></row>`,
    ];
    for (const r of rows) expect(await readTime(r)).toBeLessThan(LIMIT_MS);
  });

  it('共有文字列：<si に > が無い／閉じの無い <si>／<si> の中の <t・<rPh に > が無い・閉じが無い', async () => {
    const head = `<sst ${MAIN}>`;
    const cell = '<row r="1"><c r="A1" t="s"><v>0</v></c></row>';
    const ssts = [
      `${head}${'<si '.repeat(N)}</sst>`,
      `${head}${'<si>'.repeat(N)}</sst>`,
      `${head}<si>${'<t '.repeat(N)}</si></sst>`,
      `${head}<si>${'<t>'.repeat(N)}</si></sst>`,
      `${head}<si>${'<rPh '.repeat(N)}<t>x</t></si></sst>`,
      `${head}<si>${'<rPh>'.repeat(N)}<t>x</t></si></sst>`,
    ];
    for (const sst of ssts) expect(await readTime(cell, sst)).toBeLessThan(LIMIT_MS);
  });

  it('細工が混ざっていても、正しい要素は読める（セル・共有文字列・ふりがな除去）', async () => {
    const sst = `<sst ${MAIN}><si><t>plain</t></si><si><r><t>ab</t></r><rPh sb="0" eb="1"><t>XX</t></rPh><r><t>cd</t></r></si><si/></sst>`;
    const rows =
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" s="1"/><c r="D1" t="inlineStr"><is><t>in</t></is></c><c r="E1"><v>2.5</v></c></row>';
    const bytes = await buildXlsx({
      sheets: [{ name: 'S', sheetDataXml: rows }],
      sharedStrings: [],
      extraFiles: { 'xl/sharedStrings.xml': sst },
    });
    const sheet = await (await openXlsx(bytes)).openSheet(1);
    expect(sheet.getRows(1, 1)[0]!.cells.map((c) => c?.text)).toEqual([
      'plain',
      'abcd',
      undefined,
      'in',
      '2.5',
    ]);
  });
});

describe('TD-89-6 グループの展開でも時間が線形（深い入れ子・閉じの無いグループ・大量の子）', () => {
  const pt =
    '<xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:to>';
  /** 座標系を持つ（展開できる）グループの開きタグ */
  const OPEN =
    '<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="1" name="g"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/><a:chOff x="0" y="0"/><a:chExt cx="10" cy="10"/></a:xfrm></xdr:grpSpPr>';
  const SP =
    '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="s"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></xdr:spPr></xdr:sp>';

  /** 1 つのアンカーに content を入れた drawing を読み、かかった時間（ms）と結果を返す */
  async function read(content: string) {
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: {
            xml: wsDr([`<xdr:twoCellAnchor>${pt}${content}<xdr:clientData/></xdr:twoCellAnchor>`]),
            rels: [],
          },
        },
      ],
    });
    const book = await openXlsx(bytes);
    let drawings: readonly unknown[] = [];
    const ms = await elapsed(async () => {
      drawings = (await book.openDrawings(1)).drawings;
    });
    return { ms, drawings };
  }

  it('深い入れ子（閉じあり・10,000 段）：深さの上限で止まり、XML の大きさ×深さの上限の時間', async () => {
    const M = 10_000;
    const { ms, drawings } = await read(OPEN.repeat(M) + SP + '</xdr:grpSp>'.repeat(M));
    expect(ms).toBeLessThan(LIMIT_MS);
    // 深さ 9 のグループは展開せず破線枠 1 件
    expect(drawings).toHaveLength(1);
  });

  it('深い入れ子（閉じの無い・10,000 段）：例外にせず、時間も線形', async () => {
    const { ms } = await read(OPEN.repeat(10_000));
    expect(ms).toBeLessThan(LIMIT_MS);
  });

  it('グループの開きタグに > が無い・grpSpPr や xfrm が閉じていないものの繰り返し', async () => {
    for (const content of [
      '<xdr:grpSp '.repeat(N),
      '<xdr:grpSp>'.repeat(N),
      `<xdr:grpSp><xdr:grpSpPr>${'<a:xfrm><a:chOff x="0" y="0"/>'.repeat(N)}</xdr:grpSpPr></xdr:grpSp>`,
      OPEN + '<xdr:sp '.repeat(N) + '</xdr:grpSp>',
      OPEN + '<xdr:pic>'.repeat(N) + '</xdr:grpSp>',
      OPEN + '<mc:AlternateContent>'.repeat(N) + '</xdr:grpSp>',
    ]) {
      expect((await read(content)).ms).toBeLessThan(LIMIT_MS);
    }
  });

  it('1 つのグループに子が大量（20,000 個）：上限（図 1,000）を超えても走査は線形', async () => {
    const { ms, drawings } = await read(OPEN + SP.repeat(20_000) + '</xdr:grpSp>');
    expect(ms).toBeLessThan(LIMIT_MS);
    expect(drawings).toHaveLength(1000);
  });
});

describe('TD-91-8 自由図形（a:custGeom）の読み取りでも時間が線形（巨大な pathLst・gdLst・> の無い開きタグ・閉じの無い要素）', () => {
  const pt =
    '<xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:to>';
  const sp = (geom: string): string =>
    `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="s"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm>${geom}</xdr:spPr></xdr:sp>`;

  /** 1 つのアンカーに content を入れた drawing を読み、かかった時間（ms）と結果を返す */
  async function read(content: string, limits?: Partial<XlsxLimits>) {
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: {
            xml: wsDr([`<xdr:twoCellAnchor>${pt}${content}<xdr:clientData/></xdr:twoCellAnchor>`]),
            rels: [],
          },
        },
      ],
    });
    const book = await openXlsx(bytes, { limits });
    let found: SheetDrawings = { drawings: [], notices: [] };
    const ms = await elapsed(async () => {
      found = await book.openDrawings(1);
    });
    return { ms, ...found };
  }

  it('巨大な pathLst（命令 40,000 個・1 つの path）：上限で打ち切って shape-simplified。時間は線形', async () => {
    const cmds = '<a:lnTo><a:pt x="1" y="1"/></a:lnTo>'.repeat(N);
    const { ms, drawings, notices } = await read(
      sp(
        `<a:custGeom><a:pathLst><a:path w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${cmds}</a:path></a:pathLst></a:custGeom>`,
      ),
    );
    expect(ms).toBeLessThan(LIMIT_MS);
    expect(drawings[0]!.geom).toBeUndefined();
    expect(notices).toEqual(['shape-simplified']);
  });

  it('巨大な gdLst（guide 40,000 個）・空の path の大量の並び（40,000 個）', async () => {
    const guides = '<a:gd name="g" fmla="val 1"/>'.repeat(N);
    const a = await read(
      sp(`<a:custGeom><a:gdLst>${guides}</a:gdLst><a:pathLst><a:path/></a:pathLst></a:custGeom>`),
    );
    expect(a.ms).toBeLessThan(LIMIT_MS);
    expect(a.notices).toEqual(['shape-simplified']);
    const b = await read(
      sp(`<a:custGeom><a:pathLst>${'<a:path/>'.repeat(N)}</a:pathLst></a:custGeom>`),
    );
    expect(b.ms).toBeLessThan(LIMIT_MS);
    expect(b.notices).toEqual(['shape-simplified']);
  });

  it('小さな上限でも、命令が大量な path を打ち切って走査が終わる（上限の外の分は読まない）', async () => {
    const cmds = '<a:lnTo><a:pt x="1" y="1"/></a:lnTo>'.repeat(N);
    const { ms, notices } = await read(
      sp(
        `<a:custGeom><a:pathLst><a:path><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${cmds}</a:path></a:pathLst></a:custGeom>`,
      ),
      { maxGeomCommands: 10 },
    );
    expect(ms).toBeLessThan(LIMIT_MS);
    expect(notices).toEqual(['shape-simplified']);
  });

  it('> の無い開きタグ・閉じの無い要素・a:pt の大量の繰り返し：例外にせず、時間も線形', async () => {
    for (const geom of [
      `<a:custGeom><a:pathLst><a:path>${'<a:moveTo '.repeat(N)}</a:path></a:pathLst></a:custGeom>`,
      `<a:custGeom><a:pathLst><a:path>${'<a:pt '.repeat(N)}</a:path></a:pathLst></a:custGeom>`,
      `<a:custGeom><a:pathLst><a:path>${'<a:moveTo>'.repeat(N)}</a:path></a:pathLst></a:custGeom>`,
      `<a:custGeom><a:pathLst>${'<a:path '.repeat(N)}</a:pathLst></a:custGeom>`,
      `<a:custGeom><a:pathLst>${'<a:path>'.repeat(N)}</a:pathLst></a:custGeom>`,
      `<a:custGeom><a:gdLst>${'<a:gd '.repeat(N)}</a:gdLst></a:custGeom>`,
      `<a:custGeom><a:gdLst>${'<a:gd name="x" fmla="val 1">'.repeat(N)}</a:gdLst></a:custGeom>`,
      `<a:custGeom><a:avLst>${'<a:gd '.repeat(N)}</a:avLst></a:custGeom>`,
      `<a:custGeom>${'<a:avLst>'.repeat(N)}`,
      `<a:prstGeom prst="rect"><a:avLst>${'<a:gd '.repeat(N)}</a:avLst></a:prstGeom>`,
      `<a:prstGeom prst="rect"><a:avLst>${'<a:gd name="a" fmla="val 1"/>'.repeat(N)}</a:avLst></a:prstGeom>`,
      `<a:custGeom>${'<a:rect '.repeat(N)}`,
    ]) {
      const r = await read(sp(geom));
      expect(r.ms).toBeLessThan(LIMIT_MS);
      expect(r.drawings).toHaveLength(1);
    }
  });
});
