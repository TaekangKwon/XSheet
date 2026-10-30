// ISSUE-104：make-ref-charts.ps1 の出力（com.json と、あれば com.default.json）を、照合テストが読む ref-charts.expected.json に整える
// 使い方：node make-ref-charts.extract.mjs <out dir>   <out dir>/com.json（と com.default.json）が入力。結果は <out dir>/ref-charts.expected.json（PNG は入れない）
// 整え方：
//   - グラフの名前（ChartObject.Name）をキーにしたオブジェクトにする（charts＝ref-charts.xlsx、defaults＝ref-charts-default.xlsx）。名前の重複はエラー
//   - COM が読めなかった項目（null）・中身の無い入れ物は落とす（見つからない＝その部品が無い）
//   - 数値は 1e-4 に丸める（COM の Single の誤差 12.699999809… を除く）。長さは pt のまま（px への換算は照合テスト側）。色は RRGGBB のまま
//   - 組み立ての情報（Kind・Type・Data・Size・Check・Asked…）はそのまま残す
/* global process */
import fs from 'node:fs';
const [, , dir] = process.argv;
if (!dir) {
  console.error('usage: node make-ref-charts.extract.mjs <out dir>');
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

/** 1 つの JSON を読む（無ければ null） */
function readJson(path) {
  if (!fs.existsSync(path)) return null;
  const text = fs.readFileSync(path, 'utf8');
  return JSON.parse(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text);
}

const com = readJson(`${dir}/com.json`);
const def = readJson(`${dir}/com.default.json`);
if (!com && !def) throw new Error(`${dir} に com.json も com.default.json も無い`);

const out = {
  excelVersion: (com ?? def).excelVersion,
  unit: 'pt',
  note:
    com?.note ??
    'COM の値は pt のまま（画面 125% では 0.6pt＝1px）。色は RRGGBB（負の数は Excel の色の定数）。換算は照合テストで行う',
  notes: [...(com?.notes ?? []), ...(def?.notes ?? [])],
  charts: com ? byName(com.charts, 'charts') : {},
  defaults: def ? byName(def.charts, 'defaults') : {},
};
fs.writeFileSync(`${dir}/ref-charts.expected.json`, JSON.stringify(out, null, 1));

// 概要（内訳と、基準値の欠け）
const kinds = {};
let noScale = 0;
for (const c of Object.values(out.charts)) {
  kinds[c.Kind ?? '?'] = (kinds[c.Kind ?? '?'] ?? 0) + 1;
  const t = c.ValueAxis;
  const axisChart = ![5, -4120].includes(c.ChartType);
  if (
    axisChart &&
    !(
      t &&
      typeof t.MinimumScale === 'number' &&
      typeof t.MaximumScale === 'number' &&
      typeof t.MajorUnit === 'number'
    )
  )
    noScale++;
}
console.log(
  `charts ${Object.keys(out.charts).length} (${Object.entries(kinds)
    .map(([k, n]) => `${k} ${n}`)
    .join(', ')}), defaults ${Object.keys(out.defaults).length}, notes ${out.notes.length}`,
);
if (noScale)
  console.log(
    `警告：数値軸の最小・最大・間隔が読めていないグラフが ${noScale} 件ある（notes と com.json を確認）`,
  );
