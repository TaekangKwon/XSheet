// ISSUE-110：make-ref-charts3.ps1 の出力 excel3.json を、照合テストが読む ref-charts3.expected.json に整える
// 使い方：node make-ref-charts3.extract.mjs <out dir>   <out dir>/excel3.json が入力。結果は <out dir>/ref-charts3.expected.json（PNG は入れない）
// 整え方は make-ref-charts2.extract.mjs と同じ（名前をキーにしたオブジェクト・null と空の入れ物を落とす・数値は 1e-4 に丸める・長さは pt のまま・色は RRGGBB のまま）。
//   charts＝ref-charts3.xlsx（excel3.json）。名前の重複はエラー。予約デバイス名（com3.json など）を避けて excel3.json
/* global process */
import fs from 'node:fs';
const [, , dir] = process.argv;
if (!dir) {
  console.error('usage: node make-ref-charts3.extract.mjs <out dir>');
  process.exit(2);
}

/** 数値を 1e-4 に丸め、null と空の入れ物（{}・[]）を落とす。落ちて空になったら undefined を返す */
function tidy(v) {
  if (v === null || v === undefined) return undefined;
  if (typeof v === 'number') return Number.isInteger(v) ? v : Math.round(v * 1e4) / 1e4;
  if (Array.isArray(v)) {
    const a = v.map((x) => (x === null ? null : tidy(x)));
    // 配列は位置が意味を持つ（点・系列・凡例の項目）ので、中身が無い要素も null で位置を保つ。全部無いなら落とす
    return a.every((x) => x === null || x === undefined) ? undefined : a.map((x) => x ?? null);
  }
  if (typeof v === 'object') {
    const o = {};
    for (const [k, x] of Object.entries(v)) {
      const t = tidy(x);
      if (t !== undefined) o[k] = t;
    }
    return Object.keys(o).length ? o : undefined;
  }
  return v;
}

/** グラフの配列 → 名前をキーにしたオブジェクト（名前の重複と欠落はエラー） */
function byName(list, label) {
  const out = {};
  for (const c of list) {
    if (!c || typeof c.Name !== 'string' || c.Name === '')
      throw new Error(`${label}: 名前の無いグラフがある`);
    if (c.Name in out) throw new Error(`${label}: 名前が重複している ${c.Name}`);
    const { Name, ...rest } = c;
    out[Name] = tidy(rest) ?? {};
  }
  return out;
}

/** 1 つの JSON を読む（無ければ null）。PowerShell が配列 1 個を値にしてしまった所は呼び出し側で配列にする */
function readJson(path) {
  if (!fs.existsSync(path)) return null;
  const text = fs.readFileSync(path, 'utf8');
  return JSON.parse(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text);
}

const com = readJson(`${dir}/excel3.json`);
if (!com)
  throw new Error(`${dir} に excel3.json が無い（make-ref-charts3.ps1 の create 段を先に）`);
const asList = (c) => (Array.isArray(c) ? c : c ? [c] : []);

const out = {
  excelVersion: com.excelVersion,
  unit: 'pt',
  note:
    com.note ??
    'COM の値は pt のまま（画面 125% では 0.6pt＝1px）。色は RRGGBB（負の数は Excel の色の定数）。換算は照合テストで行う',
  notes: asList(com.notes),
  charts: byName(asList(com.charts), 'charts'),
};
fs.writeFileSync(`${dir}/ref-charts3.expected.json`, JSON.stringify(out, null, 1));

// 概要（内訳）と、見るべき欠け（COM が読めなかった所）
const kinds = {};
for (const c of Object.values(out.charts)) kinds[c.Kind ?? '?'] = (kinds[c.Kind ?? '?'] ?? 0) + 1;
const noPlot = Object.entries(out.charts)
  .filter(([, c]) => !c.PlotArea?.Inside?.Width)
  .map(([n]) => n);
const noEntries = Object.entries(out.charts)
  .filter(([, c]) => c.HasLegend && !(c.Legend?.Entries?.length > 0))
  .map(([n]) => n);
console.log(
  `charts ${Object.keys(out.charts).length} (${Object.entries(kinds)
    .map(([k, n]) => `${k} ${n}`)
    .join(', ')}), notes ${out.notes.length}`,
);
if (noPlot.length)
  console.log(
    `警告：プロットの内側が読めていないグラフ ${noPlot.length} 件：${noPlot.slice(0, 5).join(', ')}…`,
  );
if (noEntries.length)
  console.log(
    `警告：凡例の項目が読めていないグラフ ${noEntries.length} 件：${noEntries.slice(0, 5).join(', ')}…`,
  );
