/**
 * 図形の文字・線の上限と線形時間（ISSUE-87・ADR-0027 §決定 6）。
 * 上限を超えても XlsxError にせず、表は必ず読める。通知 `text-truncated` は 1 シートに 1 回。
 */
import { describe, expect, it } from 'vitest';
import { openXlsx, type XlsxLimits } from '../../src/core/index';
import {
  buildXlsx,
  lineShapeXml,
  shapeXml,
  twoCellAnchor,
  wsDr,
  type ParaSpec,
  type ShapeSpec,
} from '../../src/testing/make-xlsx';

const A1 = { col: 0, row: 0 };
const B2 = { col: 1, row: 1 };

/** 図の要素（XML）を並べた 1 シートのブックを開き、図と、表の読み取りの成否を返す */
async function open(contents: string[], limits?: Partial<XlsxLimits>) {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: [['cell']],
        drawing: { xml: wsDr(contents.map((c) => twoCellAnchor(A1, B2, c))), rels: [] },
      },
    ],
  });
  const book = await openXlsx(bytes, { limits });
  const found = await book.openDrawings(1);
  const sheet = await book.openSheet(1);
  return { ...found, cell: sheet.getRows(1, 1)[0]!.cells[0] };
}

const shape = (name: string, paras: (string | ParaSpec)[]): string =>
  shapeXml({ name, paras } as ShapeSpec);
const textOf = (d: {
  text?: { paragraphs: { runs: { text: string; br: boolean }[] }[] };
}): string[] => d.text!.paragraphs.map((p) => p.runs.map((r) => (r.br ? '\n' : r.text)).join(''));

describe('TD-87-20 線の上限：図とは別に数える', () => {
  it('線 1,001 本：先頭の 1,000 本だけ出し、too-many。表は読める', async () => {
    const lines = Array.from({ length: 1001 }, (_v, i) => lineShapeXml({ name: `l${i}` }));
    const { drawings, notices, cell } = await open(lines);
    expect(drawings).toHaveLength(1000);
    expect(drawings[0]!.name).toBe('l0');
    expect(drawings[999]!.name).toBe('l999');
    expect(notices).toEqual(['too-many']);
    expect(cell).toMatchObject({ text: 'cell' });
  });

  it('図 1,000 個と線 1,000 本は互いに押し出さない（順序は XML の順）。通知なし', async () => {
    const mixed: string[] = [];
    for (let i = 0; i < 1000; i++)
      mixed.push(lineShapeXml({ name: `l${i}` }), shapeXml({ name: `s${i}` }));
    const { drawings, notices } = await open(mixed);
    expect(drawings.filter((d) => d.kind === 'line')).toHaveLength(1000);
    expect(drawings.filter((d) => d.kind !== 'line')).toHaveLength(1000);
    expect(notices).toEqual([]);
  });

  it('線が上限に達しても図の走査は続く（先に線 1,001 本、あとに図 3 個）。逆も同じ', async () => {
    const lines = Array.from({ length: 1001 }, (_v, i) => lineShapeXml({ name: `l${i}` }));
    const shapes = [1, 2, 3].map((i) => shapeXml({ name: `s${i}` }));
    const a = await open([...lines, ...shapes]);
    expect(a.drawings.filter((d) => d.kind !== 'line').map((d) => d.name)).toEqual([
      's1',
      's2',
      's3',
    ]);
    expect(a.drawings.filter((d) => d.kind === 'line')).toHaveLength(1000);
    expect(a.notices).toEqual(['too-many']);

    const manyShapes = Array.from({ length: 1001 }, (_v, i) => shapeXml({ name: `s${i}` }));
    const lines3 = [1, 2, 3].map((i) => lineShapeXml({ name: `l${i}` }));
    const b = await open([...manyShapes, ...lines3]);
    expect(b.drawings.filter((d) => d.kind === 'line').map((d) => d.name)).toEqual([
      'l1',
      'l2',
      'l3',
    ]);
    expect(b.drawings.filter((d) => d.kind !== 'line')).toHaveLength(1000);
    expect(b.notices).toEqual(['too-many']);
  });

  it('小さな上限（maxLines・maxDrawings）でも同じ。ちょうどなら通知なし', async () => {
    const six = [0, 1, 2].flatMap((i) => [
      lineShapeXml({ name: `l${i}` }),
      shapeXml({ name: `s${i}` }),
    ]);
    const exact = await open(six, { maxLines: 3, maxDrawings: 3 });
    expect(exact.drawings).toHaveLength(6);
    expect(exact.notices).toEqual([]);
    const over = await open(six, { maxLines: 2, maxDrawings: 3 });
    expect(over.drawings.filter((d) => d.kind === 'line')).toHaveLength(2);
    expect(over.notices).toEqual(['too-many']);
  });
});

describe('TD-87-21 文字の上限：切る・捨てる。XlsxError にせず、text-truncated は 1 回だけ', () => {
  it('図形 1 つの文字数 10,000：ちょうどは切らない。超えたら 10,000 文字で切って末尾に「…」', async () => {
    const exact = await open([shape('s', ['あ'.repeat(10000)])]);
    expect(textOf(exact.drawings[0]!)).toEqual(['あ'.repeat(10000)]);
    expect(exact.notices).toEqual([]);

    const over = await open([shape('s', ['あ'.repeat(10001)])]);
    expect(textOf(over.drawings[0]!)).toEqual(['あ'.repeat(10000) + '…']);
    expect(over.notices).toEqual(['text-truncated']);
    expect(over.cell).toMatchObject({ text: 'cell' });
  });

  it('文字数は図形全体で数える（ラン・段落をまたいで切る。以降のランは捨てる）', async () => {
    const r = await open([shape('s', [{ runs: ['a'.repeat(6), 'b'.repeat(6), 'c'] }, 'd'])], {
      maxShapeTextChars: 10,
    });
    expect(textOf(r.drawings[0]!)).toEqual(['aaaaaabbbb…']);
    expect(r.notices).toEqual(['text-truncated']);
  });

  it('サロゲートペアの途中では切らない', async () => {
    const r = await open([shape('s', ['😀'.repeat(4)])], { maxShapeTextChars: 3 });
    // 3 コード単位まで＝絵文字 1 つ（2 単位）＋高位サロゲート 1 つ → 高位サロゲートは外す
    expect(textOf(r.drawings[0]!)).toEqual(['😀…']);
  });

  it('図形 1 つの段落数 500：ちょうどは残し、501 段落目から捨てる', async () => {
    const exact = await open([
      shape(
        's',
        Array.from({ length: 500 }, (_v, i) => `p${i}`),
      ),
    ]);
    expect(exact.drawings[0]!.text!.paragraphs).toHaveLength(500);
    expect(exact.notices).toEqual([]);

    const over = await open([
      shape(
        's',
        Array.from({ length: 501 }, (_v, i) => `p${i}`),
      ),
    ]);
    expect(over.drawings[0]!.text!.paragraphs).toHaveLength(500);
    expect(textOf(over.drawings[0]!)[499]).toBe('p499');
    expect(over.notices).toEqual(['text-truncated']);
  });

  it('図形 1 つのラン 2,000（a:br を含む）：ちょうどは残し、2,001 個目から捨てる', async () => {
    const runs = (n: number) =>
      Array.from({ length: n }, (_v, i) => (i % 2 === 0 ? 'x' : { br: true }));
    const exact = await open([shape('s', [{ runs: runs(2000) }])]);
    expect(exact.drawings[0]!.text!.paragraphs[0]!.runs).toHaveLength(2000);
    expect(exact.notices).toEqual([]);

    const over = await open([shape('s', [{ runs: runs(2001) }])]);
    expect(over.drawings[0]!.text!.paragraphs[0]!.runs).toHaveLength(2000);
    expect(over.notices).toEqual(['text-truncated']);
  });

  it('ランの数は段落をまたいで数える（500 段落×5 ランで 2,000 を超える）', async () => {
    const paras = Array.from({ length: 500 }, () => ({ runs: ['a', 'b', 'c', 'd', 'e'] }));
    const r = await open([shape('s', paras)]);
    const total = r.drawings[0]!.text!.paragraphs.reduce((n, p) => n + p.runs.length, 0);
    expect(total).toBe(2000);
    expect(r.notices).toEqual(['text-truncated']);
  });

  it('1 シートの文字の合計（小さな上限で確認）：残りで切り、尽きた後の図形は文字を持たせない（文字の無い図形と同じ）', async () => {
    const shapes = [1, 2, 3, 4].map((i) => shape(`s${i}`, ['0123456789']));
    // 図形 1 つ＝文字 10＋ラン 1＋段落 1＝12。2 つで 24、3 つ目は段落 1・ラン 1 を引いて残り 5 文字
    const r = await open(shapes, { maxShapeTextChars: 10, maxSheetTextChars: 31 });
    expect(r.drawings.map((d) => (d.text ? textOf(d)[0] : null))).toEqual([
      '0123456789',
      '0123456789',
      '01234…',
      null,
    ]);
    // 文字を持たない図形も、図形の項目（回転・反転・枠線と塗りの有無・形・枠線の太さと線種。ISSUE-91）は持つ。無いのは text だけ
    expect(Object.keys(r.drawings[3]!).sort()).toEqual([
      'anchor',
      'descr',
      'fillColor',
      'flipH',
      'flipV',
      'geom',
      'hasFill',
      'hasOutline',
      'kind',
      'mediaPath',
      'name',
      'outline',
      'outlineColor',
      'rot',
    ]);
    expect(r.notices).toEqual(['text-truncated']);
  });

  it('1 シートの文字の合計 1,000,000（既定。ラン・段落も 1 つずつ数える）：図形 100 個（各 10,000 文字）まで。101 個目は文字なし', async () => {
    const shapes = Array.from({ length: 101 }, (_v, i) => shape(`s${i}`, ['a'.repeat(10000)]));
    const r = await open(shapes);
    expect(r.drawings).toHaveLength(101);
    expect(r.drawings.filter((d) => d.text !== undefined)).toHaveLength(100);
    expect(r.drawings[100]!.text).toBeUndefined();
    expect(r.notices).toEqual(['text-truncated']);
    expect(r.cell).toMatchObject({ text: 'cell' });
  });

  it('合計がちょうどのとき、次の図形に文字が無ければ通知しない', async () => {
    const r = await open([shape('s1', ['0123456789']), shapeXml({ name: 'empty' })], {
      maxShapeTextChars: 10,
      maxSheetTextChars: 12,
    });
    expect(r.notices).toEqual([]);
  });

  it('合計の上限にラン（a:br を含む）と段落も 1 つずつ数える：図形 1,000 個 × a:br 2,000 個でも合計 1,000,000 で止まる', async () => {
    // 1 図形＝段落 1＋ラン 2,000（x 1＋a:br 1,999）＋文字 1＝2,002。合計の上限（既定 1,000,000）で 500 個目あたりから文字を持たせない
    const raw = '<a:r><a:t>x</a:t></a:r>' + '<a:br/>'.repeat(1999);
    const shapes = Array.from({ length: 1000 }, (_v, i) => shape(`s${i}`, [{ raw }]));
    const r = await open(shapes);
    const withText = r.drawings.filter((d) => d.text !== undefined);
    expect(r.drawings).toHaveLength(1000);
    expect(withText.length).toBeGreaterThanOrEqual(499);
    expect(withText.length).toBeLessThanOrEqual(500);
    expect(r.notices).toEqual(['text-truncated']);
    expect(r.cell).toMatchObject({ text: 'cell' });
  }, 30_000);

  it('合計の上限（小さな値）：段落とランも数えて残りで止まる。上限ちょうどは通知しない', async () => {
    // 図形 1 つ＝段落 1＋ラン 1＋文字 1＝3。上限 9 なら 3 つまでちょうど
    const shapes = [1, 2, 3, 4].map((i) => shape(`s${i}`, ['x']));
    const exact = await open(shapes.slice(0, 3), { maxSheetTextChars: 9 });
    expect(exact.drawings.map((d) => d.text !== undefined)).toEqual([true, true, true]);
    expect(exact.notices).toEqual([]);
    const over = await open(shapes, { maxSheetTextChars: 9 });
    expect(over.drawings.map((d) => d.text !== undefined)).toEqual([true, true, true, false]);
    expect(over.notices).toEqual(['text-truncated']);
  });

  it('空のラン（文字も a:br も無いもの）は一覧に入れず、ランの数にも数えない', async () => {
    const raw =
      '<a:r><a:t></a:t></a:r>'.repeat(3000) +
      '<a:r><a:rPr/></a:r>'.repeat(10) +
      '<a:r><a:t>x</a:t></a:r>';
    const r = await open([shape('s', [{ raw }])]);
    expect(r.drawings[0]!.text!.paragraphs[0]!.runs).toHaveLength(1);
    expect(textOf(r.drawings[0]!)).toEqual(['x']);
    expect(r.notices).toEqual([]);
  });

  it('文字の無い図形は、段落・ランの上限を空の要素が超えただけでは text-truncated を立てない', async () => {
    const paras = await open([
      shapeXml({ name: 'p', paras: Array.from({ length: 600 }, () => ({ runs: [] })) }),
    ]);
    expect(paras.drawings[0]!.text).toBeUndefined();
    expect(paras.notices).toEqual([]);
    const brs = await open([
      shapeXml({
        name: 'r',
        paras: [{ runs: Array.from({ length: 2500 }, () => ({ br: true })) }],
      }),
    ]);
    expect(brs.drawings[0]!.text).toBeUndefined();
    expect(brs.notices).toEqual([]);
    // 空の txBody と、文字の無いランだけの図形
    const empty = await open([
      shapeXml({ name: 'e', paras: [{ raw: '<a:r><a:t/></a:r>'.repeat(3000) }] }),
    ]);
    expect(empty.notices).toEqual([]);
  });

  it('文字の無い図形でも、上限のために失った文字があれば text-truncated（シート合計が尽きた後の図形・上限 0）', async () => {
    const afterFull = await open([shape('a', ['0123456789']), shape('b', ['lost'])], {
      maxSheetTextChars: 12,
    });
    expect(afterFull.drawings[1]!.text).toBeUndefined();
    expect(afterFull.notices).toEqual(['text-truncated']);
    const zero = await open([shape('z', ['abc'])], { maxShapeTextChars: 0 });
    expect(zero.drawings[0]!.text).toBeUndefined();
    expect(zero.notices).toEqual(['text-truncated']);
  });

  it('text-truncated は 1 シートに 1 回だけ（複数の図形が切れても）。too-many と一緒に出る', async () => {
    const shapes = [1, 2, 3].map((i) => shape(`s${i}`, ['x'.repeat(20)]));
    const r = await open(shapes, { maxShapeTextChars: 5, maxDrawings: 2 });
    expect(r.drawings).toHaveLength(2);
    expect(r.notices).toEqual(['too-many', 'text-truncated']);
  });

  it('図の上限を超えた図形の文字は読まない（切れても通知しない）', async () => {
    const r = await open([shapeXml({ name: 'a' }), shape('over', ['x'.repeat(20)])], {
      maxDrawings: 1,
      maxShapeTextChars: 5,
    });
    expect(r.notices).toEqual(['too-many']);
  });
});

describe('TD-87-22 細工した XML でも時間が線形（文字・線）', () => {
  /** 緩い固定の時間（linear-time.perf.test.ts と同じ。線形なら数十 ms、二乗なら秒単位） */
  const LIMIT_MS = 1500;
  const N = 40_000;

  /** 図形 1 つの txBody に inner を入れて読み、かかった時間（ms）を返す */
  async function timeText(inner: string): Promise<number> {
    const sp = `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="n"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/><xdr:txBody><a:bodyPr/>${inner}</xdr:txBody></xdr:sp>`;
    const bytes = await buildXlsx({
      sheets: [
        { name: 'S', rows: [['a']], drawing: { xml: wsDr([twoCellAnchor(A1, B2, sp)]), rels: [] } },
      ],
    });
    const book = await openXlsx(bytes);
    const t0 = performance.now();
    await book.openDrawings(1);
    return performance.now() - t0;
  }

  it('段落・ランの開きタグに > が無い・閉じが無い・入れ子が深い', async () => {
    const inners = [
      '<a:p '.repeat(N),
      '<a:p>'.repeat(N),
      '<a:p><a:pPr '.repeat(N),
      '<a:p><a:r>'.repeat(N),
      '<a:p><a:r><a:t>'.repeat(N),
      '<a:p><a:fld>'.repeat(N),
      `<a:p>${'<a:r '.repeat(N)}</a:p>`,
      `<a:p>${'<a:br>'.repeat(N)}</a:p>`,
      `<a:p><a:pPr>${'<a:buChar '.repeat(N)}</a:pPr></a:p>`,
      `<a:p><a:r>${'<a:t '.repeat(N)}</a:r></a:p>`,
    ];
    for (const inner of inners) expect(await timeText(inner)).toBeLessThan(LIMIT_MS);
  });

  it('巨大な 1 本の文字列（5MB）・大量の段落（10 万）・大量のラン', async () => {
    // 量が多い入力なので、負荷（並列実行・CI）で遅くなっても落ちないよう、固定の時間を 4 倍に広げる。二乗なら桁違いに遅く、なお検出できる
    const volumeMs = LIMIT_MS * 4;
    const huge = `<a:p><a:r><a:t>${'a&amp;'.repeat(800_000)}</a:t></a:r></a:p>`;
    expect(await timeText(huge)).toBeLessThan(volumeMs);
    expect(await timeText('<a:p><a:r><a:t>x</a:t></a:r></a:p>'.repeat(100_000))).toBeLessThan(
      volumeMs,
    );
    expect(await timeText(`<a:p>${'<a:r><a:t>x</a:t></a:r>'.repeat(200_000)}</a:p>`)).toBeLessThan(
      volumeMs,
    );
    expect(await timeText('<a:p/>'.repeat(200_000))).toBeLessThan(volumeMs);
  });

  it('図形・線の中の要素（spPr・a:ln・a:gd・矢じり・xfrm）の開きタグに > が無い', async () => {
    const lineBodies = [
      '<xdr:spPr>' + '<a:ln '.repeat(N) + '</xdr:spPr>',
      '<xdr:spPr><a:prstGeom prst="bentConnector3"><a:avLst>' +
        '<a:gd '.repeat(N) +
        '</a:avLst></a:prstGeom></xdr:spPr>',
      '<xdr:spPr><a:ln>' + '<a:headEnd '.repeat(N) + '<a:tailEnd '.repeat(N) + '</a:ln></xdr:spPr>',
      '<xdr:spPr>' + '<a:xfrm '.repeat(N) + '<a:prstGeom '.repeat(N) + '</xdr:spPr>',
      '<xdr:spPr>' + '<a:ln>'.repeat(N) + '</xdr:spPr>',
      '<xdr:spPr>' + '<a:extLst>'.repeat(N) + '</xdr:spPr>',
      '<xdr:style>' + '<a:lnRef '.repeat(N) + '<a:fillRef '.repeat(N) + '</xdr:style>',
    ];
    for (const body of lineBodies) {
      for (const tag of ['cxnSp', 'sp']) {
        const xml = `<xdr:${tag} macro=""><xdr:nv${tag === 'sp' ? 'SpPr' : 'CxnSpPr'}><xdr:cNvPr id="2" name="n"/></xdr:nv${tag === 'sp' ? 'SpPr' : 'CxnSpPr'}>${body}</xdr:${tag}>`;
        const bytes = await buildXlsx({
          sheets: [
            {
              name: 'S',
              rows: [['a']],
              drawing: { xml: wsDr([twoCellAnchor(A1, B2, xml)]), rels: [] },
            },
          ],
        });
        const book = await openXlsx(bytes);
        const t0 = performance.now();
        await book.openDrawings(1);
        expect(performance.now() - t0).toBeLessThan(LIMIT_MS);
      }
    }
  });

  it('閉じの無い cxnSp・sp・txBody を大量に並べても、全体が線形で終わる', async () => {
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: {
            xml: wsDr([
              `<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:to>${'<xdr:cxnSp>'.repeat(N)}${'<xdr:sp>'.repeat(N)}${'<xdr:txBody>'.repeat(N)}</xdr:twoCellAnchor>`,
            ]),
            rels: [],
          },
        },
      ],
    });
    const book = await openXlsx(bytes);
    const t0 = performance.now();
    await book.openDrawings(1);
    expect(performance.now() - t0).toBeLessThan(LIMIT_MS);
  });

  it('図 5,000 個＋線 5,000 本（それぞれ上限を大きく超える）でも線形', async () => {
    const many = Array.from({ length: 5_000 }, (_v, i) => [
      lineShapeXml({ name: `l${i}`, ln: '<a:ln w="9525"><a:tailEnd type="triangle"/></a:ln>' }),
      shape(`s${i}`, ['text']),
    ]).flat();
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: { xml: wsDr(many.map((c) => twoCellAnchor(A1, B2, c))), rels: [] },
        },
      ],
    });
    const book = await openXlsx(bytes);
    const t0 = performance.now();
    const { drawings, notices } = await book.openDrawings(1);
    expect(performance.now() - t0).toBeLessThan(LIMIT_MS);
    expect(drawings).toHaveLength(2000);
    expect(notices).toEqual(['too-many']);
  });
});

describe('TD-87-23 console に図形の文字・図の名前を出さない', () => {
  it('文字のある図形・線・上限超過の読み取りで console が呼ばれない', async () => {
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
      await open(
        [shape('SECRET-NAME', ['SECRET-TEXT'.repeat(5)]), lineShapeXml({ name: 'SECRET-LINE' })],
        {
          maxShapeTextChars: 3,
        },
      );
    } finally {
      Object.assign(console, orig);
    }
    expect(calls).toEqual([]);
  });
});
