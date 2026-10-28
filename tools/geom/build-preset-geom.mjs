// 正: InterHub docs/plan/adr/0028-xlsx-preset-shape-geometry.md（§決定 1）・docs/work/issues/90_xlsx-preset-geom-evaluator.md
// ============================================================================
// build-preset-geom.mjs — ECMA-376 の preset 定義 → 形の表（JSON）の変換器
//
//   presetShapeDefinitions.xml（Ecma International の著作物・リポジトリ直下の NOTICE 参照）を読み、
//   形ごとに av（調整値の既定）・gd（guide の式）・rect（文字領域）・paths（path 命令）だけを持つ
//   小さな JSON を出力する。ahLst（ハンドル）・cxnLst（接続点）は捨てる。依存なし（Node 標準のみ）。
//
//   使い方（既定の入出力でよい）:
//     npm run geom:build
//     node tools/geom/build-preset-geom.mjs --in <xml> --out <json> --allow rect,ellipse,...
//     （NOTICE はリポジトリ直下の 1 つだけ。package.json の files で配布物に入る）
//
//   - 出力は決定的（同じ入力・同じ許可リストなら同じバイト列）。形は原本の出現順。
//   - 線の 10 種（`line`・`straightConnector1`・`bentConnector2〜5`・`curvedConnector2〜5`）は表に入れない
//     （線は xlsx-shapes.ts の経路のまま）。`lineInv` は線ではないため入れる。
//   - `--allow` で許可リストを渡すと、その形だけを出す（容量・ライセンスで絞るための逃げ道）。
// ============================================================================

import { Buffer } from 'node:buffer';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import process from 'node:process';
import { fileURLToPath, pathToFileURL } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const DEFAULT_IN = resolve(HERE, 'presetShapeDefinitions.xml');
const DEFAULT_NOTICE = resolve(HERE, '../../NOTICE');
const DEFAULT_OUT = resolve(HERE, '../../src/geom/preset-geom.json');

/** 線として別経路で描く preset（core の `drawing-shape.ts` の `LINE_PRESETS` と同じ 10 種）。表に入れない */
export const EXCLUDED_PRESETS = Object.freeze([
  'line',
  'straightConnector1',
  'bentConnector2',
  'bentConnector3',
  'bentConnector4',
  'bentConnector5',
  'curvedConnector2',
  'curvedConnector3',
  'curvedConnector4',
  'curvedConnector5',
]);

/** 表の形式の版（形式を変えたら上げる） */
export const TABLE_VERSION = 1;

/** 出典の 1 行表記（生成物の `_source`） */
const SOURCE_LINE =
  'presetShapeDefinitions.xml from OfficeOpenXML-DrawingMLGeometries.zip (2015-05-27), ' +
  'included in the ECMA-376 Part 1 5th edition (2016-12) distribution ' +
  '(ECMA-376-1_5th_edition_december_2016.zip). (c) Ecma International. ' +
  'Converted by tools/geom/build-preset-geom.mjs. See _notice.';

/**
 * XML を最小の木（`{ name, attrs, children }`）にする。コメント・処理命令・テキストは捨てる。
 * presetShapeDefinitions.xml の規則的な構造だけを想定した簡易パーサ。
 *
 * @param {string} xml
 * @returns {{ name: string, attrs: Record<string, string>, children: any[] }} ルート要素
 */
export function parseXml(xml) {
  const tagRe = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<(\/?)([A-Za-z_][\w.:-]*)([^>]*?)(\/?)>/g;
  const attrRe = /([^\s=]+)\s*=\s*"([^"]*)"/g;
  const root = { name: '#root', attrs: Object.create(null), children: [] };
  const stack = [root];
  for (const m of xml.matchAll(tagRe)) {
    if (m[2] === undefined) continue; // コメント・処理命令
    const [, closing, name, rawAttrs, selfClose] = m;
    if (closing) {
      if (stack.length > 1) stack.pop();
      continue;
    }
    /** @type {Record<string, string>} */
    const attrs = Object.create(null);
    for (const a of rawAttrs.matchAll(attrRe)) attrs[a[1]] = a[2];
    const node = { name, attrs, children: [] };
    stack[stack.length - 1].children.push(node);
    if (!selfClose) stack.push(node);
  }
  return root.children[0] ?? root;
}

/** 引数・座標の値：整数なら数値、それ以外（組み込み変数・guide の名前）は文字列のまま */
function arg(s) {
  return /^-?\d+$/.test(s) ? Number(s) : s;
}

/** `<pt x y>` の子から [x, y]（値は {@link arg}） */
function pt(node) {
  return [arg(node.attrs.x), arg(node.attrs.y)];
}

/** `<path>` 1 つを `{ w?, h?, fill?, stroke?, c: 命令の列 }` にする（既定値は省く） */
function convertPath(node) {
  const out = {};
  if (node.attrs.w !== undefined) out.w = Number(node.attrs.w);
  if (node.attrs.h !== undefined) out.h = Number(node.attrs.h);
  if (node.attrs.fill !== undefined && node.attrs.fill !== 'norm') out.fill = node.attrs.fill;
  if (node.attrs.stroke === 'false' || node.attrs.stroke === '0') out.stroke = false;
  const c = [];
  for (const cmd of node.children) {
    const pts = cmd.children.filter((n) => n.name === 'pt').map(pt);
    switch (cmd.name) {
      case 'moveTo':
        c.push(['M', ...pts[0]]);
        break;
      case 'lnTo':
        c.push(['L', ...pts[0]]);
        break;
      case 'arcTo':
        c.push([
          'A',
          arg(cmd.attrs.wR),
          arg(cmd.attrs.hR),
          arg(cmd.attrs.stAng),
          arg(cmd.attrs.swAng),
        ]);
        break;
      case 'quadBezTo':
        c.push(['Q', ...pts[0], ...pts[1]]);
        break;
      case 'cubicBezTo':
        c.push(['C', ...pts[0], ...pts[1], ...pts[2]]);
        break;
      case 'close':
        c.push(['Z']);
        break;
      default:
        throw new Error(`未知の path 命令: ${cmd.name}`);
    }
  }
  out.c = c;
  return out;
}

/** 演算子ごとの引数の数（評価器 xlsx-geom.ts の `ARITY` と同じ 17 種） */
const ARITY = new Map([
  ['*/', 3],
  ['+-', 3],
  ['+/', 3],
  ['?:', 3],
  ['abs', 1],
  ['sqrt', 1],
  ['val', 1],
  ['max', 2],
  ['min', 2],
  ['mod', 3],
  ['pin', 3],
  ['sat2', 3],
  ['cat2', 3],
  ['sin', 2],
  ['cos', 2],
  ['tan', 2],
  ['at2', 2],
]);

/**
 * `<gd name fmla>` → `[名前, 演算子, 引数…]`（引数は {@link arg}）。
 * 原本には引数が 1 つ多い式がある（例 circularArrow の `+- xH 0 dxB 0`）。余分な引数は仕様上使われないため、
 * 評価器を厳密（引数の数の誤りは失敗）に保てるよう、変換時に演算子の引数の数へ切り詰める。足りない・未知の演算子は止める。
 */
function convertGuide(node) {
  const parts = node.attrs.fmla.trim().split(/\s+/);
  const op = parts[0];
  const n = ARITY.get(op);
  if (n === undefined || parts.length - 1 < n) {
    throw new Error(`式が不正: ${node.attrs.name} ${node.attrs.fmla}`);
  }
  return [node.attrs.name, op, ...parts.slice(1, 1 + n).map(arg)];
}

/** 形 1 つの定義（`<avLst>`・`<gdLst>`・`<rect>`・`<pathLst>` ほか）を表の項目にする */
function convertShape(node) {
  const out = {};
  for (const child of node.children) {
    if (child.name === 'avLst') {
      const av = child.children.map((g) => {
        const m = /^val\s+(-?\d+)$/.exec(g.attrs.fmla.trim());
        if (!m) throw new Error(`avLst の式が val ではない: ${node.name} ${g.attrs.name}`);
        return [g.attrs.name, Number(m[1])];
      });
      if (av.length > 0) out.av = av;
    } else if (child.name === 'gdLst') {
      const gd = child.children.map(convertGuide);
      if (gd.length > 0) out.gd = gd;
    } else if (child.name === 'rect') {
      out.rect = ['l', 't', 'r', 'b'].map((k) => arg(child.attrs[k]));
    } else if (child.name === 'pathLst') {
      out.paths = child.children.map(convertPath);
    }
    // ahLst・cxnLst・それ以外は描画に要らないので捨てる
  }
  if (!out.paths) out.paths = [];
  return out;
}

/**
 * 原本の XML 文字列と NOTICE の本文から、形の表の JSON 文字列を作る（決定的）。
 *
 * @param {string} xml presetShapeDefinitions.xml の内容
 * @param {string} notice NOTICE の内容（出力の `_notice` に入れる）
 * @param {{ allow?: string[] | null }} [options] `allow`：出す形の許可リスト（省略で、除外 10 種以外すべて）
 * @returns {string} 末尾改行つきの JSON
 */
export function convertPresetGeom(xml, notice, options = {}) {
  const text = xml.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
  const root = parseXml(text);
  const allow = options.allow ? new Set(options.allow) : null;
  const excluded = new Set(EXCLUDED_PRESETS);
  /** @type {Map<string, string>} 形名 → 1 行の JSON（出現順） */
  const lines = new Map();
  for (const shape of root.children) {
    const name = shape.name;
    if (excluded.has(name)) continue;
    if (allow && !allow.has(name)) continue;
    const json = JSON.stringify(convertShape(shape));
    const prev = lines.get(name);
    if (prev !== undefined) {
      // 原本には `upDownArrow` が同じ内容で 2 回ある。内容が違えば取り違えなので止める
      if (prev !== json) throw new Error(`同名で内容の違う形がある: ${name}`);
      continue;
    }
    lines.set(name, json);
  }
  const head = [
    `"_source": ${JSON.stringify(SOURCE_LINE)}`,
    `"_notice": ${JSON.stringify(notice.replace(/\r\n?/g, '\n').trimEnd())}`,
    `"version": ${TABLE_VERSION}`,
  ];
  const body = [...lines].map(([name, json]) => `${JSON.stringify(name)}: ${json}`);
  return `{\n${head.join(',\n')},\n"shapes": {\n${body.join(',\n')}\n}\n}\n`;
}

/** コマンドライン引数を読む */
function parseArgs(argv) {
  const out = {};
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--in') out.in = argv[++i];
    else if (a === '--out') out.out = argv[++i];
    else if (a === '--allow') out.allow = argv[++i].split(',').filter(Boolean);
    else throw new Error(`未知の引数: ${a}`);
  }
  return out;
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  const args = parseArgs(process.argv);
  if (args.allow && !args.out) {
    console.error(
      '[build-preset-geom] --allow を使うときは --out を指定してください（同梱の生成物を上書きしないため）',
    );
    process.exit(1);
  }
  const xml = readFileSync(args.in ?? DEFAULT_IN, 'utf8');
  const notice = readFileSync(DEFAULT_NOTICE, 'utf8');
  const json = convertPresetGeom(xml, notice, { allow: args.allow ?? null });
  if (args.allow) {
    const known = new Set(Object.keys(JSON.parse(convertPresetGeom(xml, notice)).shapes));
    const unknown = args.allow.filter((n) => !known.has(n));
    if (unknown.length > 0) {
      console.error(
        `[build-preset-geom] 表に無い名前（未知・または除外の線）: ${unknown.join(', ')}`,
      );
      process.exit(1);
    }
  }
  const outPath = args.out ?? DEFAULT_OUT;
  writeFileSync(outPath, json, 'utf8');
  console.log(`[build-preset-geom] ${outPath} (${Buffer.byteLength(json)} bytes)`);
}
