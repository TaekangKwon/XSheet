/**
 * 確認用の例のページ（examples/print.html・M7-20）：`renderForPrint`（ADR-0010 §2）。
 * 表示中のシートをページに分けて `#print-host` に描き、「印刷」は `window.print()` を出して、終わったら（afterprint）片付ける。
 * 「プレビュー」は同じ描画を画面に出す。用紙は A4（縦・横）で、`@page` と `page` を同じ大きさにする。
 * 開発サーバー（`npm run examples` → /examples/print.html）で動かす。Pages には載せない。
 */
import {
  renderWorkbook,
  type PrintOptions,
  type PrintResult,
  type WorkbookView,
} from '../src/render/index';

function byId<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (el === null) {
    throw new Error(`#${id} が無い`);
  }
  return el as T;
}

const SAMPLES = [
  'showcase.xlsx',
  'ref-colors.xlsx',
  'ref-charts.xlsx',
  'ref-drawings.xlsx',
  'ref-shapes.xlsx',
];

/**
 * 1 ページの中身の大きさ（CSS px。96px = 1in）。A4（210×297mm）から Excel の「標準」の余白（左右 0.7in・上下 0.75in）を引いた値。
 * `@page` の余白（print.html）と同じ。
 */
const PAGES: Record<string, NonNullable<PrintOptions['page']>> = {
  portrait: { width: 659, height: 978 },
  landscape: { width: 988, height: 649 },
};

const sample = byId<HTMLSelectElement>('sample');
const file = byId<HTMLInputElement>('file');
const paper = byId<HTMLSelectElement>('paper');
const fit = byId<HTMLInputElement>('fit');
const headers = byId<HTMLInputElement>('headers');
const gridlines = byId<HTMLInputElement>('gridlines');
const theme = byId<HTMLSelectElement>('theme');
const previewBtn = byId<HTMLButtonElement>('preview');
const printBtn = byId<HTMLButtonElement>('print');
const statusOut = byId<HTMLOutputElement>('status');
const viewer = byId<HTMLElement>('viewer');
const host = byId<HTMLElement>('print-host');
const nonce =
  document.querySelector<HTMLLinkElement>('link[rel="stylesheet"][nonce]')?.nonce || undefined;

/** 横の用紙の @page（縦は print.html の既定） */
const pageStyle = document.createElement('style');
if (nonce) pageStyle.nonce = nonce;
document.head.append(pageStyle);

let view: WorkbookView | null = null;
let printed: PrintResult | null = null;
let ctl: AbortController | null = null;
/** 画面にプレビューを出している（指定を替えたら描き直す） */
let previewing = false;

/** プレビューの表示を替える */
function setPreview(on: boolean): void {
  previewing = on;
  document.body.classList.toggle('ex-preview', on);
  previewBtn.textContent = on ? '表に戻る' : 'プレビュー';
}

/** 前の印刷の描画を片付ける（描いている途中なら中断する） */
function clearPrint(): void {
  ctl?.abort();
  ctl = null;
  printed?.destroy();
  printed = null;
}

/** 今の指定で印刷の描画を作り直す */
async function renderPrint(): Promise<PrintResult | null> {
  clearPrint();
  if (!view) return null;
  pageStyle.textContent =
    paper.value === 'landscape' ? '@page { size: A4 landscape; margin: 0.75in 0.7in; }' : '';
  const c = new AbortController();
  ctl = c;
  const t0 = performance.now();
  try {
    const r = await view.renderForPrint(host, {
      page: PAGES[paper.value],
      fit: fit.checked ? 'width' : undefined,
      headers: headers.checked,
      gridlines: gridlines.checked,
      signal: c.signal,
    });
    printed = r;
    const ms = Math.round(performance.now() - t0);
    const cells = host.querySelectorAll('.xs-print td').length;
    statusOut.value = `${r.pages} ページ・td ${cells}・${ms}ms${r.truncated ? '・上限で打ち切り' : ''}`;
    return r;
  } catch (e) {
    if ((e as Error).name !== 'AbortError') {
      statusOut.value = `失敗：${(e as Error).message}`;
    }
    return null;
  } finally {
    if (ctl === c) ctl = null;
  }
}

async function open(data: ArrayBuffer): Promise<void> {
  clearPrint();
  setPreview(false);
  view?.destroy();
  view = await renderWorkbook(viewer, data, { nonce });
}

async function openSample(name: string): Promise<void> {
  const res = await fetch(`../fixtures/excel/${name}`);
  await open(await res.arrayBuffer());
}

for (const name of SAMPLES) {
  sample.append(new Option(name, name, name === SAMPLES[0], name === SAMPLES[0]));
}
sample.addEventListener('change', () => void openSample(sample.value));
file.addEventListener('change', () => {
  const f = file.files?.[0];
  if (f) void f.arrayBuffer().then(open);
});
theme.addEventListener('change', () => {
  document.documentElement.dataset.theme = theme.value;
});
for (const el of [paper, fit, headers, gridlines]) {
  el.addEventListener('change', () => {
    if (previewing) void renderPrint();
    else clearPrint();
  });
}

previewBtn.addEventListener('click', () => {
  if (previewing) {
    clearPrint();
    setPreview(false);
  } else {
    setPreview(true);
    void renderPrint();
  }
});

printBtn.addEventListener('click', () => {
  void renderPrint().then((r) => {
    if (!r) return;
    // 印刷の画面を閉じたら片付ける（window.print() が待たないブラウザもあるので afterprint で）
    window.addEventListener(
      'afterprint',
      () => {
        if (!previewing) clearPrint();
      },
      { once: true },
    );
    window.print();
  });
});

document.documentElement.dataset.theme = theme.value;
void openSample(SAMPLES[0]!);
