/**
 * XSheet の例のページ（M3-2）：.xlsx を選ぶ・ドロップする・サンプルから選ぶと、タブ付きで表示する。
 * テーマ（OS に合わせる・明るい・ダーク）・文言（日本語・English）・図の有無を切り替えられる。
 * v0.5.0 の操作：組み込みの和暦（builtinDates）・見出しの有無・倍率（setZoom）・表示中のシートの検索（findText → setHighlights・focusHighlight）・
 * クリックした所のセル（cellAt）とヒット（highlightAt）。
 * 開発サーバー（`npm run examples`）で動かす。ライブラリは src から直接読む（ビルドの前でも動くように）。
 */
import {
  findText,
  openXlsx,
  type BuiltinDates,
  type TextMatch,
  type XlsxBook,
} from '../src/core/index';
import {
  MESSAGES_EN,
  MESSAGES_JA,
  renderWorkbook,
  type WorkbookView,
  type Zoom,
} from '../src/render/index';

/** サンプル（fixtures/excel の自作ファイル。照合用に Excel で作ったもの） */
const SAMPLES: readonly { file: string; label: string }[] = [
  { file: 'showcase.xlsx', label: '見本（README の画面）' },
  { file: 'ref-drawings.xlsx', label: '画像・図形' },
  { file: 'ref-shapes.xlsx', label: '図形の種類' },
  { file: 'ref-geoms.xlsx', label: '図形の形' },
  { file: 'ref-colors.xlsx', label: '色・塗り' },
  { file: 'ref-charts.xlsx', label: 'グラフ' },
  { file: 'ref-charts2.xlsx', label: 'グラフ（2）' },
  { file: 'ref-calibri.xlsx', label: '列幅（Calibri）' },
  { file: 'ref-yugothic.xlsx', label: '列幅（游ゴシック）' },
  { file: 'ref-malgun.xlsx', label: '列幅（Malgun Gothic）' },
  { file: 'ref-weekday.xlsx', label: '曜日の書式' },
  { file: 'ref-era.xlsx', label: '和暦の書式' },
];

const THEME_KEY = 'xsheet-example-theme';

type Theme = 'system' | 'light' | 'dark';

/** id で要素を引く（無ければページの作りの誤り） */
function byId<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (el === null) {
    throw new Error(`#${id} が無い`);
  }
  return el as T;
}

const fileInput = byId<HTMLInputElement>('file');
const sampleSelect = byId<HTMLSelectElement>('sample');
const themeSelect = byId<HTMLSelectElement>('theme');
const langSelect = byId<HTMLSelectElement>('lang');
const drawingsCheck = byId<HTMLInputElement>('drawings');
const tabPositionSelect = byId<HTMLSelectElement>('tab-position');
const headersCheck = byId<HTMLInputElement>('headers');
const builtinEraCheck = byId<HTMLInputElement>('builtin-era');
const zoomSelect = byId<HTMLSelectElement>('zoom');
const findForm = byId<HTMLFormElement>('find');
const findInput = byId<HTMLInputElement>('find-text');
const findPrev = byId<HTMLButtonElement>('find-prev');
const findCount = byId<HTMLOutputElement>('find-count');
const statusOut = byId<HTMLOutputElement>('status');
const viewer = byId<HTMLElement>('viewer');

/**
 * ページの CSP の nonce（M3-3）。style.css の <link> に付いた値を、書式の <style> にも付けてもらう。
 * nonce の属性値はブラウザが隠すことがあるため、getAttribute ではなく nonce プロパティで読む。
 */
const nonce =
  document.querySelector<HTMLLinkElement>('link[rel="stylesheet"][nonce]')?.nonce || undefined;

let view: WorkbookView | null = null;
/** 表示中のファイル（文言・図の切り替えで開き直す） */
let current: { name: string; data: ArrayBuffer } | null = null;
/** 開く処理の番号（遅れて終わった古い処理の結果を捨てる） */
let seq = 0;

/** 倍率の選択肢の値 → setZoom に渡す値 */
function zoomValue(): Zoom {
  return zoomSelect.value === 'fit' ? { fit: 'width' } : Number(zoomSelect.value);
}

/** 組み込みの和暦の指定 */
function builtinDatesValue(): BuiltinDates {
  return builtinEraCheck.checked ? 'japanese-era' : 'default';
}

/** 検索用のブック（renderWorkbook は開いたブックを外に出さないので、同じバイト列をもう 1 つ開く） */
let searchBook: { data: ArrayBuffer; builtinDates: BuiltinDates; book: XlsxBook } | null = null;
/** 今の検索の結果（どのシートで何を探したか・今のヒットの番号） */
let found: { sheet: number; query: string; matches: TextMatch[]; current: number } | null = null;
/** 実行中の検索の中断 */
let finding: AbortController | null = null;

/** 検索の結果を捨てる（表のハイライトはシートの切り替え・表示し直しで XSheet が消す） */
function resetFind(): void {
  finding?.abort();
  finding = null;
  found = null;
  findCount.value = '';
  findPrev.disabled = true;
}

/** 今のヒットを替えて件数を出す */
function focusHit(index: number, scroll = true): void {
  if (view === null || found === null || found.matches.length === 0) {
    return;
  }
  found.current = index;
  view.focusHighlight(index, { scroll });
  findCount.value = `${index + 1} / ${found.matches.length}`;
}

/** 表示中のシートで探す。同じ文字・同じシートなら次のヒットへ進む */
async function find(): Promise<void> {
  const query = findInput.value;
  const sheetIndex = view?.activeSheet ?? null;
  if (view === null || current === null || sheetIndex === null || query === '') {
    resetFind();
    view?.setHighlights(null);
    return;
  }
  if (found !== null && found.query === query && found.sheet === sheetIndex) {
    focusHit((found.current + 1) % Math.max(found.matches.length, 1));
    return;
  }
  resetFind();
  const ac = new AbortController();
  finding = ac;
  const target = view;
  findCount.value = '検索中…';
  const data = current.data;
  // 表示と同じ指定で開く（組み込みの和暦の有無で文字が変わるため）
  const builtinDates = builtinDatesValue();
  if (searchBook?.data !== data || searchBook.builtinDates !== builtinDates) {
    searchBook?.book.close();
    searchBook = { data, builtinDates, book: await openXlsx(data, { builtinDates }) };
  }
  const sheet = await searchBook.book.openSheet(sheetIndex);
  const started = performance.now();
  const result = await findText(sheet, query, { signal: ac.signal });
  if (finding !== ac || view !== target) {
    return; // 次の検索・表示し直しが始まった（中断した検索は AbortError で下の catch に行く）
  }
  finding = null;
  found = { sheet: sheetIndex, query, matches: result.matches, current: 0 };
  view.setHighlights(result.matches);
  const ms = Math.round(performance.now() - started);
  if (result.matches.length === 0) {
    findCount.value = `なし（${ms} ms）`;
    return;
  }
  findPrev.disabled = false;
  focusHit(0);
  statusOut.value = `${result.matches.length}${result.truncated ? ' 件以上' : ' 件'}（${ms} ms）`;
}

/** 列の番号（1 始まり）→ 列記号 */
function columnName(c: number): string {
  let name = '';
  for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}

/** テーマを当てる。system は属性を外して OS の設定（prefers-color-scheme）に任せる */
function applyTheme(theme: Theme): void {
  if (theme === 'system') {
    delete document.documentElement.dataset.theme;
  } else {
    document.documentElement.dataset.theme = theme;
  }
}

/** 保存したテーマ（保存できない環境では system） */
function loadTheme(): Theme {
  try {
    const saved = localStorage.getItem(THEME_KEY);
    if (saved === 'light' || saved === 'dark' || saved === 'system') {
      return saved;
    }
  } catch {
    // 保存先を使えない（プライベートウィンドウなど）
  }
  return 'system';
}

function saveTheme(theme: Theme): void {
  try {
    localStorage.setItem(THEME_KEY, theme);
  } catch {
    // 保存できなくても表示は変わらない
  }
}

/** 表示中のファイルを、今の指定（文言・図）で表示し直す */
async function show(): Promise<void> {
  if (current === null) {
    return;
  }
  const mySeq = ++seq;
  resetFind();
  view?.destroy();
  view = null;
  viewer.replaceChildren();
  const started = performance.now();
  statusOut.value = `${current.name} を開いています…`;
  const next = await renderWorkbook(viewer, current.data, {
    messages: langSelect.value === 'en' ? MESSAGES_EN : MESSAGES_JA,
    drawings: drawingsCheck.checked,
    tabPosition: tabPositionSelect.value === 'bottom' ? 'bottom' : 'top',
    nonce,
    zoom: zoomValue(),
    headers: headersCheck.checked,
    // 組み込みの日付の書式 27〜36・50〜58 を日本語版 Excel と同じに（ref-era.xlsx の builtin のシートで違いが見える）
    builtinDates: builtinDatesValue(),
    onSheetChange: resetFind, // シートが替わると XSheet がハイライトを消すので、検索の結果も捨てる
  });
  if (mySeq !== seq) {
    next.destroy(); // 開いている間に別のファイルが選ばれた
    return;
  }
  view = next;
  statusOut.value = `${current.name}（${Math.round(performance.now() - started)} ms）`;
}

/** 読み込みの依頼の番号（ファイルの選択・ドロップ・サンプル。遅れて読み終えた古い依頼を捨てる） */
let request = 0;

function report(e: unknown): void {
  statusOut.value = e instanceof Error ? `エラー：${e.message}` : 'エラー';
}

/** 内容を読んで表示する。読み終える前に次の依頼が来たら、この依頼の結果（エラーも）は捨てる */
async function load(name: string, read: () => Promise<ArrayBuffer>): Promise<void> {
  const myRequest = ++request;
  try {
    const data = await read();
    if (myRequest !== request) {
      return;
    }
    current = { name, data };
    await show();
  } catch (e) {
    if (myRequest === request) {
      report(e);
    }
  }
}

/** 今のファイルを表示し直す（文言・図の切り替え）。表示し直す前に次の表示が始まったら、エラーは出さない */
function reshow(): void {
  const mySeq = seq + 1;
  show().catch((e: unknown) => {
    if (mySeq === seq) {
      report(e);
    }
  });
}

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (file !== undefined) {
    sampleSelect.value = '';
    void load(file.name, () => file.arrayBuffer());
  }
});

for (const s of SAMPLES) {
  sampleSelect.append(new Option(s.label, s.file));
}
sampleSelect.addEventListener('change', () => {
  const file = sampleSelect.value;
  if (file === '') {
    return;
  }
  fileInput.value = '';
  void load(file, async () => {
    // ページからの相対（開発サーバーの /examples/ でも、Pages の <サイト>/examples/ でも同じ場所を指す）
    const res = await fetch(`../fixtures/excel/${file}`);
    if (!res.ok) {
      throw new Error(`${file} を読めない（${res.status}）`);
    }
    return res.arrayBuffer();
  });
});

themeSelect.value = loadTheme();
applyTheme(themeSelect.value as Theme);
themeSelect.addEventListener('change', () => {
  const theme = themeSelect.value as Theme;
  applyTheme(theme);
  saveTheme(theme);
});

langSelect.addEventListener('change', reshow);
drawingsCheck.addEventListener('change', reshow);
tabPositionSelect.addEventListener('change', reshow);
headersCheck.addEventListener('change', reshow);
builtinEraCheck.addEventListener('change', reshow);
zoomSelect.addEventListener('change', () => view?.setZoom(zoomValue()));

findForm.addEventListener('submit', (e) => {
  e.preventDefault();
  find().catch((err: unknown) => {
    if (!(err instanceof DOMException && err.name === 'AbortError')) {
      report(err);
    }
  });
});
findPrev.addEventListener('click', () => {
  if (found !== null && found.matches.length > 0) {
    focusHit((found.current - 1 + found.matches.length) % found.matches.length);
  }
});

// クリック：表の中ならセルの番地を出し、ヒットなら今のヒットにする
viewer.addEventListener('click', (e) => {
  if (view === null) {
    return;
  }
  const cell = view.cellAt(e.clientX, e.clientY);
  if (cell !== null) {
    statusOut.value = `${columnName(cell.c)}${cell.r}`;
  }
  const hit = view.highlightAt(e.clientX, e.clientY);
  if (hit >= 0) {
    focusHit(hit, false);
  }
});

// ?sample=<ファイル名> で開いたら、そのサンプルを最初から表示する（README の画面を撮るため）。一覧に無い名前は無視する
const initialSample = new URLSearchParams(location.search).get('sample');
if (initialSample !== null && SAMPLES.some((s) => s.file === initialSample)) {
  sampleSelect.value = initialSample;
  sampleSelect.dispatchEvent(new Event('change'));
}

// ドロップで開く
viewer.addEventListener('dragover', (e) => {
  e.preventDefault();
  viewer.classList.add('is-dragover');
});
viewer.addEventListener('dragleave', () => viewer.classList.remove('is-dragover'));
viewer.addEventListener('drop', (e) => {
  e.preventDefault();
  viewer.classList.remove('is-dragover');
  const file = e.dataTransfer?.files[0];
  if (file !== undefined) {
    sampleSelect.value = '';
    void load(file.name, () => file.arrayBuffer());
  }
});
