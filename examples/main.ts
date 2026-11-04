/**
 * XSheet の例のページ（M3-2）：.xlsx を選ぶ・ドロップする・サンプルから選ぶと、タブ付きで表示する。
 * テーマ（OS に合わせる・明るい・ダーク）・文言（日本語・English）・図の有無を切り替えられる。
 * 開発サーバー（`npm run examples`）で動かす。ライブラリは src から直接読む（ビルドの前でも動くように）。
 */
import { MESSAGES_EN, MESSAGES_JA, renderWorkbook, type WorkbookView } from '../src/render/index';

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
  view?.destroy();
  view = null;
  viewer.replaceChildren();
  const started = performance.now();
  statusOut.value = `${current.name} を開いています…`;
  const next = await renderWorkbook(viewer, current.data, {
    messages: langSelect.value === 'en' ? MESSAGES_EN : MESSAGES_JA,
    drawings: drawingsCheck.checked,
    nonce,
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
    const res = await fetch(`/fixtures/excel/${file}`);
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
