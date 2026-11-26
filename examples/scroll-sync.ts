/**
 * 確認用の例のページ（examples/scroll-sync.html・M7-19）：`scrollElement`（ADR-0010 §1）で 2 つの表を一緒に動かす。
 * 合わせ方は InterX-Viewer の `scroll`（縦の割合）と同じ。自分が動かした位置の scroll は相手へ送り返さない。
 * 開発サーバー（`npm run examples` → /examples/scroll-sync.html）で動かす。Pages には載せない。
 */
import { renderWorkbook, type WorkbookView } from '../src/render/index';

function byId<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (el === null) {
    throw new Error(`#${id} が無い`);
  }
  return el as T;
}

const SAMPLES = ['showcase.xlsx', 'ref-colors.xlsx', 'ref-charts.xlsx', 'ref-drawings.xlsx'];

const modeSelect = byId<HTMLSelectElement>('mode');
const statusOut = byId<HTMLOutputElement>('status');
const nonce =
  document.querySelector<HTMLLinkElement>('link[rel="stylesheet"][nonce]')?.nonce || undefined;

/** 片側の表示と、自分が動かした（相手から写した）位置 */
interface Side {
  view: WorkbookView | null;
  expected: { top: number; left: number } | null;
  /** scroll を受ける要素（シートを替えても同じ要素だが、開き直したら替わる） */
  bound: HTMLElement | null;
}

const left: Side = { view: null, expected: null, bound: null };
const right: Side = { view: null, expected: null, bound: null };

/** 割合（余りが無ければ 0） */
function ratio(pos: number, size: number, client: number): number {
  const max = size - client;
  return max > 0 ? Math.min(1, Math.max(0, pos / max)) : 0;
}

/** from のスクロールを to へ割合で写す */
function follow(from: Side, to: Side): void {
  const a = from.view?.scrollElement;
  const b = to.view?.scrollElement;
  if (!a || !b || modeSelect.value === 'off') return;
  // 相手から写した位置の scroll なら送り返さない（InterX-Viewer の createScrollSync と同じく 1px 以内）
  const exp = from.expected;
  from.expected = null;
  if (exp && Math.abs(a.scrollTop - exp.top) <= 1 && Math.abs(a.scrollLeft - exp.left) <= 1) return;
  const top = Math.round(
    ratio(a.scrollTop, a.scrollHeight, a.clientHeight) *
      Math.max(0, b.scrollHeight - b.clientHeight),
  );
  const leftPos =
    modeSelect.value === 'xy'
      ? Math.round(
          ratio(a.scrollLeft, a.scrollWidth, a.clientWidth) *
            Math.max(0, b.scrollWidth - b.clientWidth),
        )
      : b.scrollLeft;
  if (top === b.scrollTop && leftPos === b.scrollLeft) return;
  b.scrollTop = top;
  b.scrollLeft = leftPos;
  to.expected = { top: b.scrollTop, left: b.scrollLeft }; // 端で丸められた後の値
  statusOut.value = `${from === left ? '左 → 右' : '右 → 左'}：縦 ${(
    ratio(a.scrollTop, a.scrollHeight, a.clientHeight) * 100
  ).toFixed(1)}%`;
}

const handlers = new Map<Side, () => void>([
  [left, () => follow(left, right)],
  [right, () => follow(right, left)],
]);

async function open(side: Side, host: HTMLElement, data: ArrayBuffer): Promise<void> {
  const handler = handlers.get(side)!;
  side.bound?.removeEventListener('scroll', handler);
  side.bound = null;
  side.view?.destroy();
  side.view = await renderWorkbook(host, data, { nonce });
  side.bound = side.view.scrollElement;
  side.bound?.addEventListener('scroll', handler, { passive: true });
}

async function openSample(side: Side, host: HTMLElement, file: string): Promise<void> {
  const res = await fetch(`../fixtures/excel/${file}`);
  await open(side, host, await res.arrayBuffer());
}

function setup(side: Side, name: 'left' | 'right', initial: string): void {
  const host = byId<HTMLElement>(name);
  const select = byId<HTMLSelectElement>(`sample-${name}`);
  const input = byId<HTMLInputElement>(`file-${name}`);
  for (const file of SAMPLES) {
    select.append(new Option(file, file, file === initial, file === initial));
  }
  select.addEventListener('change', () => {
    void openSample(side, host, select.value);
  });
  input.addEventListener('change', () => {
    const f = input.files?.[0];
    if (f) void f.arrayBuffer().then((data) => open(side, host, data));
  });
  void openSample(side, host, initial);
}

setup(left, 'left', 'showcase.xlsx');
setup(right, 'right', 'showcase.xlsx');
