// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import {
  createGeometry,
  type ChartFailReason,
  type ChartResult,
  type ImageResult,
  type SheetDrawing,
} from '../../src/core/index';
import { DEFAULT_THEME, type Theme } from '../../src/core/theme';
import { emptyLayout, type ChartModel } from '../../src/core/types';
import { barModel, series } from './chart-fixtures';
import { EMU_PER_PX, layoutDrawings, viewWindow } from '../../src/render/drawings';
import { createImageLayer, type ImageLayer } from '../../src/render/image-layer';
import { MESSAGES_JA } from '../../src/render/messages';

/**
 * 画像レイヤーのグラフ（ISSUE-103・ADR-0031 §決定 11・P1・P8）：見える範囲に入った時に読み、描けたら破線枠の代わりに `<svg>`。
 * 描けなければ破線枠のまま（文言は今と同じ）。キャッシュ・読み込みの順序・シートの切替／dispose の競合を確かめる。
 * 文言は既定（`MESSAGES_JA`。InterHub の `t()` のキー文字列を置き換えた）。
 */

const px = (n: number): number => n * EMU_PER_PX;
const geo = createGeometry(emptyLayout(), 1000, 100);

/** oneCell の図（A1 から w×h px。col・row は 0 始まり） */
function fig(
  col: number,
  row: number,
  over: Partial<SheetDrawing> = {},
  w = 200,
  h = 120,
): SheetDrawing {
  return {
    anchor: { type: 'oneCell', from: { col, colOff: 0, row, rowOff: 0 }, cx: px(w), cy: px(h) },
    kind: 'chart',
    name: `chart${col}_${row}`,
    descr: '',
    mediaPath: null,
    chartPath: `xl/charts/chart${col}_${row}.xml`,
    ...over,
  };
}

const okChart = (model: ChartModel = barModel(), theme?: Theme): ChartResult => ({
  ok: true,
  chart: model,
  ...(theme ? { theme } : {}),
});

/** 面が塗りなしのグラフ（場面の backdrop が真＝ダーク表示で白い下地が要る） */
const NONE = barModel({ chartArea: { fill: { kind: 'none' } } });

const ALL = viewWindow(0, 0, 3000, 3000, 0, 0);
/** 図から遠い範囲（全部外れる） */
const FAR = viewWindow(0, 100000, 3000, 3000, 0, 0);

const flush = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

// XSheet: グラフの描画は遅れて読む部品（render/chart-draw.ts）。最初の読み込みを先に済ませ、flush 1 回で描き終わるようにする
beforeAll(async () => {
  await import('../../src/render/chart-draw');
});

function makeLayer() {
  let n = 0;
  const revoked: string[] = [];
  const layer = createImageLayer({
    createUrl: () => `blob:l/${++n}`,
    revokeUrl: (u) => revoked.push(u),
  });
  document.body.append(layer.element);
  return { layer, revoked };
}

const imageOk = (): ImageResult => ({
  ok: true,
  format: 'png',
  mime: 'image/png',
  width: 40,
  height: 30,
  bytes: new Uint8Array([1]),
});

function setSheet(
  layer: ImageLayer,
  drawings: SheetDrawing[],
  readChart?: (p: string) => Promise<ChartResult>,
  readImage: (p: string) => Promise<ImageResult> = async () => imageOk(),
) {
  layer.setSheet({ drawings, boxes: layoutDrawings(geo, drawings), readImage, readChart });
}

/** 外から解決できる Promise */
function deferred<T>() {
  let resolve!: (v: T) => void;
  const promise = new Promise<T>((r) => (resolve = r));
  return { promise, resolve };
}

const figs = (layer: ImageLayer): HTMLElement[] => [
  ...layer.element.querySelectorAll<HTMLElement>('.xs-figure'),
];

afterEach(() => {
  document.body.replaceChildren();
});

describe('TD-103-10 グラフが見える範囲に入ると svg に置き換わる', () => {
  it('読み込み中は破線枠「グラフ」・読めたら破線枠が消えて <svg> 1 枚になる（位置と大きさは図の矩形）', async () => {
    const { layer } = makeLayer();
    const d = deferred<ChartResult>();
    setSheet(layer, [fig(1, 2)], () => d.promise);
    layer.update(ALL);
    expect(layer.element.querySelector('.xs-frame')!.textContent).toBe(MESSAGES_JA.chart);
    expect(layer.element.querySelector('svg')).toBeNull();
    d.resolve(okChart());
    await flush();
    const f = figs(layer)[0]!;
    expect(f.querySelector('.xs-frame')).toBeNull();
    expect(f.children).toHaveLength(1);
    const svg = f.querySelector('svg.xs-chart')!;
    expect([svg.getAttribute('width'), svg.getAttribute('height')]).toEqual(['200', '120']);
    expect([f.style.left, f.style.top, f.style.width, f.style.height]).toEqual([
      '64px',
      '40px',
      '200px',
      '120px',
    ]);
    expect(svg.querySelectorAll('path, text').length).toBeGreaterThan(0);
  });

  it('読み上げ：role="img"・<title> は i18n の「グラフ」（図の名前・説明は使わない）・aria-hidden なし', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [fig(0, 0, { name: '<秘密>', descr: '説明' })], async () => okChart());
    layer.update(ALL);
    await flush();
    const svg = layer.element.querySelector('svg.xs-chart')!;
    expect(svg.getAttribute('role')).toBe('img');
    expect(svg.hasAttribute('aria-hidden')).toBe(false);
    expect(svg.querySelector('title')!.textContent).toBe(MESSAGES_JA.chart);
    expect(svg.outerHTML).not.toContain('秘密');
  });

  it('グループの中のグラフ（drawing.group）も、グループの子の矩形（box.rect）に svg で出る', async () => {
    const { layer } = makeLayer();
    const inner = fig(2, 3, {
      group: {
        levels: [{ ch: { x: 0, y: 0, cx: 1000, cy: 1000 }, own: null }],
        xfrm: { x: 250, y: 250, cx: 500, cy: 500 },
      },
    });
    const drawings = [inner];
    const boxes = layoutDrawings(geo, drawings);
    expect(boxes).toHaveLength(1);
    const r = boxes[0]!.rect;
    setSheet(layer, drawings, async () => okChart());
    layer.update(ALL);
    await flush();
    const f = figs(layer)[0]!;
    expect([f.style.left, f.style.top, f.style.width, f.style.height]).toEqual([
      `${r.x}px`,
      `${r.y}px`,
      `${r.w}px`,
      `${r.h}px`,
    ]);
    expect(r.w).toBeLessThan(200); // 外側のアンカー（200×120）より小さい子の矩形
    const svg = f.querySelector('svg.xs-chart')!;
    expect(svg.getAttribute('width')).toBe(String(Math.round(r.w * 1000) / 1000));
    expect(f.querySelector('.xs-frame')).toBeNull();
  });

  it('グラフの面に塗りがあれば、図の層に下地（xs-backdrop の兄弟 svg）は足さない', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [fig(0, 0)], async () => okChart());
    layer.update(ALL);
    await flush();
    expect(layer.element.querySelector('svg.xs-backdrop')).toBeNull();
  });

  it('result.theme がレイアウトに渡る（系列の既定の色＝テーマの accent1）', async () => {
    const { layer } = makeLayer();
    const theme = new Map(DEFAULT_THEME) as Map<string, number>;
    theme.set('accent1', 0x123456);
    setSheet(layer, [fig(0, 0)], async () => okChart(barModel(), theme as Theme));
    layer.update(ALL);
    await flush();
    const fills = [...layer.element.querySelectorAll('svg path')].map((p) =>
      p.getAttribute('fill'),
    );
    expect(fills).toContain('#123456');
  });

  it('グラフの面が塗りなしなら、下地は svg の中ではなく図の直前の兄弟 svg.xs-backdrop（図の矩形と同じ位置と大きさ・色は CSS）', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [fig(1, 2)], async () => okChart(NONE));
    layer.update(ALL);
    await flush();
    const svg = layer.element.querySelector('svg.xs-chart')!;
    // グラフの svg の中に下地の矩形は無い（中にあると、グラフより下の図形・画像まで隠す）
    expect(svg.querySelector('.xs-chart-backdrop')).toBeNull();
    expect(svg.querySelector('.xs-backdrop-part')).toBeNull();
    const kids = [...layer.element.children];
    expect(kids.map((k) => k.getAttribute('class'))).toEqual(['xs-backdrop', 'xs-figure']);
    const back = kids[0] as SVGSVGElement;
    expect([back.style.left, back.style.top, back.style.width, back.style.height]).toEqual([
      '64px',
      '40px',
      '200px',
      '120px',
    ]);
    const part = back.querySelector('rect')!;
    expect(part.getAttribute('class')).toBe('xs-backdrop-part is-fill');
    expect([part.getAttribute('width'), part.getAttribute('height')]).toEqual(['200', '120']);
    // 色・表示は CSS のテーマ変数（style・属性に色を入れない）
    expect(part.hasAttribute('fill')).toBe(false);
    expect(part.hasAttribute('style')).toBe(false);
    expect(back.hasAttribute('fill')).toBe(false);
  });
});

describe('TD-103-16 塗りなしのグラフの下地は、図形と同じ最背面の層（グラフより下の図形・画像を隠さない）', () => {
  const colored = (over: Partial<SheetDrawing>): SheetDrawing =>
    fig(0, 0, {
      kind: 'shape',
      mediaPath: null,
      chartPath: undefined,
      geom: { type: 'preset', prst: 'rect', adj: [] },
      hasOutline: true,
      hasFill: true,
      outlineColor: { rgb: 0, alpha: 1 },
      fillColor: { rgb: 0xff0000, alpha: 1 },
      ...over,
    });
  const classes = (layer: ImageLayer): (string | null)[] =>
    [...layer.element.children].map((k) => k.getAttribute('class'));

  it('下の図形（先の drawing）と塗りなしのグラフ：下地はどちらも最背面の層の兄弟で、グラフの svg の中には入らない', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [colored({}), fig(0, 0)], async () => okChart(NONE));
    layer.update(ALL);
    await flush();
    expect(classes(layer)).toEqual([
      'xs-backdrop', // 図形の下地
      'xs-figure', // 図形（グラフの下）
      'xs-backdrop', // グラフの下地
      'xs-figure', // グラフ
    ]);
    const kids = [...layer.element.children];
    // グラフの下地は、図形の要素の中にも、グラフの svg の中にも入っていない（z-index: -1 の兄弟として図形より後ろに描かれる）
    expect(kids[2]!.parentElement).toBe(layer.element);
    expect(kids[3]!.contains(kids[2]!)).toBe(false);
    expect(kids[1]!.contains(kids[2]!)).toBe(false);
    expect(kids[3]!.querySelector('[class*="backdrop"]')).toBeNull(); // グラフの svg の中に下地は無い
    expect(kids[3]!.querySelector('.xs-chart')).not.toBeNull();
  });

  it('読み込み中は下地なし。読み込みが終わると下地が図の直前に入る', async () => {
    const { layer } = makeLayer();
    const d = deferred<ChartResult>();
    setSheet(layer, [fig(0, 0)], () => d.promise);
    layer.update(ALL);
    expect(classes(layer)).toEqual(['xs-figure']); // 読み込み中は破線枠だけ
    d.resolve(okChart(NONE));
    await flush();
    expect(classes(layer)).toEqual(['xs-backdrop', 'xs-figure']);
  });

  it('範囲を出入りしても下地は 1 つだけ・同じ要素が戻る。外れている間は残らない', async () => {
    const { layer } = makeLayer();
    const read = vi.fn(async () => okChart(NONE));
    setSheet(layer, [fig(0, 0)], read);
    layer.update(ALL);
    await flush();
    const back1 = layer.element.querySelector('svg.xs-backdrop')!;
    layer.update(FAR);
    expect(layer.element.children).toHaveLength(0); // 下地も図と一緒に外れる
    layer.update(ALL);
    expect(layer.element.querySelectorAll('svg.xs-backdrop')).toHaveLength(1);
    expect(layer.element.querySelector('svg.xs-backdrop')).toBe(back1); // 作り直さない
    expect(classes(layer)).toEqual(['xs-backdrop', 'xs-figure']);
    await flush();
    expect(read).toHaveBeenCalledTimes(1);
    expect(layer.element.querySelectorAll('svg.xs-backdrop')).toHaveLength(1);
  });

  it('読み込み中に外れて、外れている間に読み込みが終わっても、下地は足さない。戻した時に図と一緒に入る', async () => {
    const { layer } = makeLayer();
    const d = deferred<ChartResult>();
    setSheet(layer, [fig(0, 0)], () => d.promise);
    layer.update(ALL);
    layer.update(FAR);
    d.resolve(okChart(NONE));
    await flush();
    expect(layer.element.children).toHaveLength(0);
    layer.update(ALL);
    expect(classes(layer)).toEqual(['xs-backdrop', 'xs-figure']);
  });

  it('描けなかったグラフ（失敗）には下地を敷かない', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [fig(0, 0)], async () => ({ ok: false, reason: 'broken' }));
    layer.update(ALL);
    await flush();
    expect(layer.element.querySelector('svg.xs-backdrop')).toBeNull();
    expect(classes(layer)).toEqual(['xs-figure']);
  });

  it('白い塗りのあるグラフは下地なし。塗りなしのグラフも svg の中身は下地の有無で変わらない', async () => {
    const solid = barModel({
      chartArea: { fill: { kind: 'solid', color: { rgb: 0xffffff, alpha: 1 } } },
    });
    const a = makeLayer().layer;
    setSheet(a, [fig(0, 0)], async () => okChart(solid));
    a.update(ALL);
    await flush();
    expect(a.element.querySelector('svg.xs-backdrop')).toBeNull();

    const b = makeLayer().layer;
    setSheet(b, [fig(0, 0)], async () => okChart(NONE));
    b.update(ALL);
    await flush();
    const svgB = b.element.querySelector('svg.xs-chart')!;
    expect(svgB.querySelector('[class*="backdrop"]')).toBeNull();
  });

  it('setSheet・dispose で下地も DOM に残さない。後から結果が返っても足さない', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [fig(0, 0)], async () => okChart(NONE));
    layer.update(ALL);
    await flush();
    expect(layer.element.querySelectorAll('svg.xs-backdrop')).toHaveLength(1);
    setSheet(layer, [fig(0, 0)], async () => okChart(NONE));
    expect(layer.element.children).toHaveLength(0);

    const d = deferred<ChartResult>();
    setSheet(layer, [fig(0, 0)], () => d.promise);
    layer.update(ALL);
    layer.dispose();
    d.resolve(okChart(NONE));
    await flush();
    expect(layer.element.children).toHaveLength(0);
  });
});

describe('TD-103-11 描けないグラフは今の破線枠のまま（P8）', () => {
  const FRAME = MESSAGES_JA.chart;

  it('chartPath の無いグラフは読まずに破線枠。readChart の無いシートも破線枠', async () => {
    const { layer } = makeLayer();
    const read = vi.fn(async () => okChart());
    setSheet(layer, [fig(0, 0, { chartPath: undefined }), fig(0, 5)], read);
    layer.update(ALL);
    await flush();
    expect(read).toHaveBeenCalledTimes(1); // chartPath のある方だけ
    const frames = layer.element.querySelectorAll('.xs-frame');
    expect([...frames].map((f) => f.textContent)).toEqual([FRAME]);

    const second = makeLayer().layer;
    setSheet(second, [fig(0, 0)]); // readChart なし
    second.update(ALL);
    await flush();
    expect(second.element.querySelector('.xs-frame')!.textContent).toBe(FRAME);
    expect(second.element.querySelector('svg')).toBeNull();
  });

  it.each<ChartFailReason>(['unsupported', 'too-large', 'broken', 'no-data'])(
    'readChart が %s で失敗しても、破線枠のまま・文言は同じ・svg は無い',
    async (reason) => {
      const { layer } = makeLayer();
      setSheet(layer, [fig(0, 0)], async () => ({ ok: false, reason }));
      layer.update(ALL);
      await flush();
      expect(layer.element.querySelector('.xs-frame')!.textContent).toBe(FRAME);
      expect(layer.element.querySelector('svg')).toBeNull();
    },
  );

  it('readChart が例外を投げても破線枠のまま（表示は止まらない）', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [fig(0, 0)], () => Promise.reject(new Error('closed')));
    layer.update(ALL);
    await flush();
    expect(layer.element.querySelector('.xs-frame')!.textContent).toBe(FRAME);
  });

  it('場面が null（要素が上限を超える）なら破線枠のまま', async () => {
    const { layer } = makeLayer();
    const huge = barModel({
      type: 'pie',
      barDir: undefined,
      grouping: 'standard',
      varyColors: true,
      catAxis: undefined,
      valAxis: undefined,
      categories: Array.from({ length: 5001 }, (_, i) => `c${i}`),
      series: [
        series(
          0,
          's',
          Array.from({ length: 5001 }, () => 1),
        ),
      ],
    });
    setSheet(layer, [fig(0, 0)], async () => okChart(huge));
    layer.update(ALL);
    await flush();
    expect(layer.element.querySelector('.xs-frame')!.textContent).toBe(FRAME);
    expect(layer.element.querySelector('svg')).toBeNull();
  });

  it('失敗は覚える（範囲の出入りで読み直さない）', async () => {
    const { layer } = makeLayer();
    const read = vi.fn(async (): Promise<ChartResult> => ({ ok: false, reason: 'unsupported' }));
    setSheet(layer, [fig(0, 0)], read);
    layer.update(ALL);
    await flush();
    layer.update(FAR);
    layer.update(ALL);
    await flush();
    expect(read).toHaveBeenCalledTimes(1);
    expect(layer.element.querySelector('.xs-frame')!.textContent).toBe(FRAME);
  });
});

describe('TD-103-12 キャッシュ・読み込みの順序（P1）', () => {
  it('範囲を出入りしても readChart は 1 回・svg は作り直さず同じ要素が戻る', async () => {
    const { layer } = makeLayer();
    const read = vi.fn(async () => okChart());
    setSheet(layer, [fig(0, 0)], read);
    layer.update(ALL);
    await flush();
    const svg1 = layer.element.querySelector('svg')!;
    layer.update(FAR);
    expect(layer.element.querySelector('svg')).toBeNull();
    layer.update(ALL);
    expect(layer.element.querySelector('svg')).toBe(svg1); // await 無しで即座に戻る
    await flush();
    expect(read).toHaveBeenCalledTimes(1);
    expect(layer.element.querySelector('.xs-frame')).toBeNull();
  });

  it('範囲の外にある図は読まない。見える範囲に入った時に初めて読む', async () => {
    const { layer } = makeLayer();
    const read = vi.fn(async () => okChart());
    // 100 行目（≒ 2,000px 下）の図
    setSheet(layer, [fig(0, 0), fig(0, 100)], read);
    layer.update(viewWindow(0, 0, 400, 400, 0, 0));
    await flush();
    expect(read).toHaveBeenCalledTimes(1);
    layer.update(viewWindow(0, 1800, 400, 600, 0, 0));
    await flush();
    expect(read).toHaveBeenCalledTimes(2);
  });

  it('読み込みは 1 つずつ順に走る（前の結果が返るまで次を呼ばない）', async () => {
    const { layer } = makeLayer();
    const ds = [deferred<ChartResult>(), deferred<ChartResult>(), deferred<ChartResult>()];
    let calls = 0;
    const read = vi.fn(() => ds[calls++]!.promise);
    setSheet(layer, [fig(0, 0), fig(0, 5), fig(0, 10)], read);
    layer.update(ALL);
    await flush();
    expect(read).toHaveBeenCalledTimes(1);
    ds[0]!.resolve(okChart());
    await flush();
    expect(read).toHaveBeenCalledTimes(2);
    ds[1]!.resolve({ ok: false, reason: 'too-large' });
    await flush();
    expect(read).toHaveBeenCalledTimes(3);
    ds[2]!.resolve(okChart());
    await flush();
    expect(layer.element.querySelectorAll('svg')).toHaveLength(2);
    expect(layer.element.querySelectorAll('.xs-frame')).toHaveLength(1);
  });

  it('同じ部品を指す図が複数あっても readChart は 1 回・svg は図ごとに別の要素', async () => {
    const { layer } = makeLayer();
    const read = vi.fn(async () => okChart());
    setSheet(
      layer,
      [fig(0, 0, { chartPath: 'xl/charts/x.xml' }), fig(0, 5, { chartPath: 'xl/charts/x.xml' })],
      read,
    );
    layer.update(ALL);
    await flush();
    expect(read).toHaveBeenCalledTimes(1);
    const svgs = layer.element.querySelectorAll('svg');
    expect(svgs).toHaveLength(2);
    expect(svgs[0]).not.toBe(svgs[1]);
  });

  it('読み込み中に図が範囲から外れても、戻した時にキャッシュした svg を使う（読み直さない）', async () => {
    const { layer } = makeLayer();
    const d = deferred<ChartResult>();
    const read = vi.fn(() => d.promise);
    setSheet(layer, [fig(0, 0)], read);
    layer.update(ALL);
    layer.update(FAR);
    d.resolve(okChart());
    await flush();
    expect(layer.element.querySelector('svg')).toBeNull(); // 外れている間は足さない
    layer.update(ALL);
    expect(layer.element.querySelector('svg')).not.toBeNull();
    expect(read).toHaveBeenCalledTimes(1);
  });
});

describe('TD-103-13 後始末：setSheet・dispose で残さず、後から返った結果を足さない', () => {
  it('setSheet で前のシートの svg を捨てる。新しいシートは新しく読む', async () => {
    const { layer } = makeLayer();
    const read1 = vi.fn(async () => okChart());
    setSheet(layer, [fig(0, 0)], read1);
    layer.update(ALL);
    await flush();
    expect(layer.element.querySelector('svg')).not.toBeNull();
    const read2 = vi.fn(async () => okChart());
    setSheet(layer, [fig(0, 0)], read2);
    expect(layer.element.children).toHaveLength(0);
    layer.update(ALL);
    expect(layer.element.querySelector('svg')).toBeNull(); // キャッシュは持ち越さない
    await flush();
    expect(read2).toHaveBeenCalledTimes(1);
    expect(layer.element.querySelectorAll('svg')).toHaveLength(1);
  });

  it('読み込み中に setSheet しても、古いシートの結果は新しい表示に入らない', async () => {
    const { layer } = makeLayer();
    const d = deferred<ChartResult>();
    setSheet(layer, [fig(0, 0)], () => d.promise);
    layer.update(ALL);
    // 別のシート（同じ位置にグラフがあるが、読み込みは別）
    const d2 = deferred<ChartResult>();
    setSheet(layer, [fig(0, 0)], () => d2.promise);
    layer.update(ALL);
    d.resolve(okChart());
    await flush();
    expect(layer.element.querySelector('svg')).toBeNull(); // 古い結果は捨てる
    expect(layer.element.querySelector('.xs-frame')!.textContent).toBe(MESSAGES_JA.chart);
    d2.resolve(okChart());
    await flush();
    expect(layer.element.querySelectorAll('svg')).toHaveLength(1);
  });

  it('setSheet(null) で未開始の読み込みを取り消す（readChart を呼ばない）', async () => {
    const { layer } = makeLayer();
    const d = deferred<ChartResult>();
    const read = vi.fn(() => d.promise);
    setSheet(layer, [fig(0, 0), fig(0, 5)], read);
    layer.update(ALL);
    await flush();
    expect(read).toHaveBeenCalledTimes(1);
    layer.setSheet(null);
    d.resolve(okChart());
    await flush();
    expect(read).toHaveBeenCalledTimes(1); // 2 つ目は呼ばれない
    expect(layer.element.children).toHaveLength(0);
  });

  it('dispose で DOM に何も残さず、後から結果が返っても足さない。以後 setSheet・update は何もしない', async () => {
    const { layer } = makeLayer();
    const d = deferred<ChartResult>();
    const read = vi.fn(() => d.promise);
    setSheet(layer, [fig(0, 0), fig(0, 5)], read);
    layer.update(ALL);
    await flush();
    layer.dispose();
    expect(layer.element.children).toHaveLength(0);
    d.resolve(okChart());
    await flush();
    expect(layer.element.children).toHaveLength(0);
    expect(read).toHaveBeenCalledTimes(1);
    setSheet(layer, [fig(0, 0)], read);
    layer.update(ALL);
    await flush();
    expect(layer.element.children).toHaveLength(0);
  });

  it('描けた後の dispose で svg を外す', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [fig(0, 0)], async () => okChart());
    layer.update(ALL);
    await flush();
    layer.dispose();
    expect(layer.element.querySelector('svg')).toBeNull();
  });
});

describe('TD-103-14 SmartArt・画像・図形の出方は今と同じ', () => {
  it('SmartArt は chartPath があっても読まずに破線枠。画像は画像のまま。グラフと混ぜても互いに影響しない', async () => {
    const { layer } = makeLayer();
    const readChart = vi.fn(async () => okChart());
    const readImage = vi.fn(async () => imageOk());
    setSheet(
      layer,
      [
        fig(0, 0),
        fig(0, 5, { kind: 'smartart', chartPath: 'xl/charts/ignored.xml' }),
        fig(0, 10, { kind: 'image', mediaPath: 'xl/media/image1.png', chartPath: undefined }),
        fig(0, 15, { kind: 'shape', mediaPath: null, chartPath: undefined }),
      ],
      readChart,
      readImage,
    );
    layer.update(ALL);
    await flush();
    expect(readChart).toHaveBeenCalledTimes(1);
    expect(readChart).toHaveBeenCalledWith('xl/charts/chart0_0.xml');
    expect(readImage).toHaveBeenCalledTimes(1);
    const frames = [...layer.element.querySelectorAll('.xs-frame')].map((f) => f.textContent);
    expect(frames).toEqual([MESSAGES_JA.smartart, MESSAGES_JA.shape]);
    expect(layer.element.querySelectorAll('img')).toHaveLength(1);
    expect(layer.element.querySelectorAll('svg.xs-chart')).toHaveLength(1);
  });

  it('重ね順は drawing の順のまま（グラフが描けても要素の順は変わらない）', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [fig(0, 0), fig(0, 5, { kind: 'smartart' }), fig(0, 10)], async () =>
      okChart(),
    );
    layer.update(ALL);
    await flush();
    // 行 0・5・10 の図が、この順（行 5 は SmartArt の破線枠）
    expect(figs(layer).map((f) => parseFloat(f.style.top))).toEqual([0, 100, 200]);
    expect(figs(layer).map((f) => f.firstElementChild!.tagName.toLowerCase())).toEqual([
      'svg',
      'div',
      'svg',
    ]);
  });
});

describe('TD-109-6 塗りなしのグラフの下地の角丸（面と同じ半径。ISSUE-109）', () => {
  it('下地の矩形は面と同じ半径 rx・ry=11（roundedCorners の無いグラフ＝角丸）。roundedCorners=false は付けない', async () => {
    const round = makeLayer();
    setSheet(round.layer, [fig(1, 2)], async () => okChart(NONE));
    round.layer.update(ALL);
    await flush();
    const part = round.layer.element.querySelector('svg.xs-backdrop rect')!;
    expect([part.getAttribute('rx'), part.getAttribute('ry')]).toEqual(['11', '11']);
    // 色は CSS のまま（属性に入れない）
    expect(part.hasAttribute('fill')).toBe(false);
    document.body.replaceChildren();
    const square = makeLayer();
    const sq = barModel({ chartArea: { fill: { kind: 'none' } }, roundedCorners: false });
    setSheet(square.layer, [fig(1, 2)], async () => okChart(sq));
    square.layer.update(ALL);
    await flush();
    const part2 = square.layer.element.querySelector('svg.xs-backdrop rect')!;
    expect(part2.hasAttribute('rx')).toBe(false);
    expect(part2.hasAttribute('ry')).toBe(false);
  });
});
