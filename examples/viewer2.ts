/**
 * 確認用の例のページ（examples/viewer2.html）：ADR-0007 の未決の 2 つを試す `view2`（src/facade/view2.ts）を動かす。
 * - 検索は `viewer.find` だけで済む（ブックを外で開き直さない）。シートが替わったら `onSheetChange` で探し直す。
 * - 「組み込みの表」で、日本語版 Excel の表（view2 の既定）と今の表（view の既定）を比べる。
 * 開発サーバー（`npm run examples` → /examples/viewer2.html）で動かす。Pages には載せない。
 */
import type { Zoom } from '../src/render/index';
import { view2, type Viewer2 } from '../src/facade/view2';

function byId<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (el === null) {
    throw new Error(`#${id} が無い`);
  }
  return el as T;
}

const fileInput = byId<HTMLInputElement>('file');
const sampleSelect = byId<HTMLSelectElement>('sample');
const builtinSelect = byId<HTMLSelectElement>('builtin');
const headersCheck = byId<HTMLInputElement>('headers');
const zoomSelect = byId<HTMLSelectElement>('zoom');
const findForm = byId<HTMLFormElement>('find');
const findInput = byId<HTMLInputElement>('find-text');
const findPrev = byId<HTMLButtonElement>('find-prev');
const findCount = byId<HTMLOutputElement>('find-count');
const statusOut = byId<HTMLOutputElement>('status');
const container = byId<HTMLElement>('viewer');

const nonce =
  document.querySelector<HTMLLinkElement>('link[rel="stylesheet"][nonce]')?.nonce || undefined;

let viewer: Viewer2 | null = null;
let current: { name: string; data: ArrayBuffer } | null = null;
/** 今の検索の結果（何を探したか・ヒットの数・今のヒット） */
let found: { query: string; count: number; current: number } | null = null;

function zoomValue(): Zoom {
  return zoomSelect.value === 'fit' ? { fit: 'width' } : Number(zoomSelect.value);
}

function columnName(c: number): string {
  let name = '';
  for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}

function report(e: unknown): void {
  if (e instanceof Error && e.name === 'AbortError') return; // 次の検索・シートの切り替えで中断した
  statusOut.value = e instanceof Error ? `エラー：${e.message}` : 'エラー';
}

function clearFind(): void {
  found = null;
  findCount.value = '';
  findPrev.disabled = true;
}

function focusHit(index: number, scroll = true): void {
  if (viewer === null || found === null || found.count === 0) return;
  found.current = index;
  viewer.focusHighlight(index, { scroll });
  findCount.value = `${index + 1} / ${found.count}`;
}

/** 表示中のシートで探す（同じ文字なら次のヒットへ） */
async function find(): Promise<void> {
  if (viewer === null) return;
  const query = findInput.value;
  if (found !== null && found.query === query && found.count > 0) {
    focusHit((found.current + 1) % found.count);
    return;
  }
  clearFind();
  if (query === '') {
    await viewer.find('');
    return;
  }
  findCount.value = '検索中…';
  const started = performance.now();
  const result = await viewer.find(query);
  const ms = Math.round(performance.now() - started);
  found = { query, count: result.matches.length, current: 0 };
  if (found.count === 0) {
    findCount.value = `なし（${ms} ms）`;
    return;
  }
  findPrev.disabled = false;
  findCount.value = `1 / ${found.count}`;
  statusOut.value = `シート ${viewer.activeSheet}：${found.count}${result.truncated ? ' 件以上' : ' 件'}（${ms} ms）`;
}

/** 今のファイルを今の指定で表示する */
async function show(): Promise<void> {
  if (current === null) return;
  clearFind();
  const started = performance.now();
  statusOut.value = `${current.name} を開いています…`;
  // 同じ container に呼ぶと、view2 が前の表示を外す
  viewer = await view2(container, current.data, {
    css: 'none', // style.css は <link> で読み込み済み
    nonce,
    zoom: zoomValue(),
    headers: headersCheck.checked,
    builtinDates: builtinSelect.value === 'default' ? 'default' : undefined,
    onSheetChange: (index) => {
      // シートが替わるとハイライトが消えるので、同じ文字で探し直す
      statusOut.value = `シート ${index} を表示`;
      found = null;
      if (findInput.value !== '') find().catch(report);
      else clearFind();
    },
  });
  container.querySelector('.ex-hint')?.remove();
  statusOut.value = `${current.name}（${Math.round(performance.now() - started)} ms・シート ${viewer.activeSheet}）`;
}

async function load(name: string, read: () => Promise<ArrayBuffer>): Promise<void> {
  try {
    current = { name, data: await read() };
    await show();
  } catch (e) {
    report(e);
  }
}

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (file !== undefined) {
    sampleSelect.value = '';
    void load(file.name, () => file.arrayBuffer());
  }
});
sampleSelect.addEventListener('change', () => {
  const file = sampleSelect.value;
  if (file === '') return;
  fileInput.value = '';
  void load(file, async () => {
    const res = await fetch(`../fixtures/excel/${file}`);
    if (!res.ok) throw new Error(`${file} を読めない（${res.status}）`);
    return res.arrayBuffer();
  });
});
builtinSelect.addEventListener('change', () => show().catch(report));
headersCheck.addEventListener('change', () => show().catch(report));
zoomSelect.addEventListener('change', () => viewer?.setZoom(zoomValue()));

findForm.addEventListener('submit', (e) => {
  e.preventDefault();
  find().catch(report);
});
findPrev.addEventListener('click', () => {
  if (found !== null && found.count > 0) focusHit((found.current - 1 + found.count) % found.count);
});

container.addEventListener('click', (e) => {
  if (viewer === null) return;
  const cell = viewer.cellAt(e.clientX, e.clientY);
  if (cell !== null) statusOut.value = `${columnName(cell.c)}${cell.r}`;
  const hit = viewer.highlightAt(e.clientX, e.clientY);
  if (hit >= 0) focusHit(hit, false);
});

const initial = new URLSearchParams(location.search).get('sample');
if (initial !== null && [...sampleSelect.options].some((o) => o.value === initial)) {
  sampleSelect.value = initial;
  sampleSelect.dispatchEvent(new Event('change'));
}
