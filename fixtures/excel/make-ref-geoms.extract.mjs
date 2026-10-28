// ISSUE-91：解凍した ref-geoms.xlsx の drawing XML から図形ごとの形・枠線の値を取り出し、COM の値（com.json）と結合して ref-geoms.expected.json を作る
// 使い方：node make-ref-geoms.extract.mjs <dir>   <dir>/x に xlsx の解凍物、<dir>/out/com.json に make-ref-geoms.ps1 の出力。結果は <dir>/expected.json
/* global process */
import fs from 'node:fs';
const [, , dir] = process.argv;
const xml = fs.readFileSync(`${dir}/x/xl/drawings/drawing1.xml`, 'utf8');
const com = JSON.parse(fs.readFileSync(`${dir}/out/com.json`, 'utf8'));

// ---- 最小 XML パーサ ----
function parse(src) {
  const root = { tag: '#root', attrs: {}, children: [], text: '' };
  const stack = [root];
  const re =
    /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<\/([^\s>]+)\s*>|<([^\s/>]+)((?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let m;
  const dec = (s) =>
    s
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
      .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCodePoint(parseInt(n, 16)))
      .replace(/&amp;/g, '&');
  while ((m = re.exec(src))) {
    if (m[1]) {
      stack.pop();
      continue;
    }
    if (m[2]) {
      const attrs = {};
      for (const a of m[3].matchAll(/([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g))
        attrs[a[1]] = dec(a[2] ?? a[3]);
      const el = { tag: m[2], attrs, children: [], text: '' };
      stack[stack.length - 1].children.push(el);
      if (!m[4]) stack.push(el);
      continue;
    }
    if (m[5] !== undefined) stack[stack.length - 1].text += dec(m[5]);
  }
  return root;
}
const tree = parse(xml);
const ch = (el, tag) => el.children.filter((c) => c.tag === tag);
const c1 = (el, tag) => el && el.children.find((c) => c.tag === tag);
const num = (v) => (v === undefined ? undefined : Number(v));
const pick = (a, keys) => {
  const o = {};
  for (const k of keys) if (a[k] !== undefined) o[k] = a[k];
  return o;
};

function xfrmOf(spPr) {
  const x = c1(spPr, 'a:xfrm');
  if (!x) return undefined;
  const off = c1(x, 'a:off'),
    ext = c1(x, 'a:ext');
  return {
    off: off && { x: num(off.attrs.x), y: num(off.attrs.y) },
    ext: ext && { cx: num(ext.attrs.cx), cy: num(ext.attrs.cy) },
    rot: num(x.attrs.rot),
    flipH: x.attrs.flipH,
    flipV: x.attrs.flipV,
  };
}
function lnOf(spPr) {
  const ln = c1(spPr, 'a:ln');
  if (!ln) return undefined;
  const o = { w: num(ln.attrs.w) };
  o.noFill = !!c1(ln, 'a:noFill');
  o.solidFill = !!c1(ln, 'a:solidFill');
  const d = c1(ln, 'a:prstDash');
  if (d) o.prstDash = d.attrs.val;
  const h = c1(ln, 'a:headEnd');
  if (h) o.headEnd = pick(h.attrs, ['type', 'w', 'len']);
  const t = c1(ln, 'a:tailEnd');
  if (t) o.tailEnd = pick(t.attrs, ['type', 'w', 'len']);
  return o;
}
function fillOf(spPr) {
  for (const k of ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill'])
    if (c1(spPr, k)) return k.slice(2);
  return undefined;
}
/** prstGeom（prst と avLst の gd）または custGeom（avLst・gdLst の数、rect、path の w・h・fill・stroke と命令の列。点は文字列のまま） */
function geomOf(spPr) {
  const g = c1(spPr, 'a:prstGeom');
  if (g) {
    const av = c1(g, 'a:avLst');
    return {
      type: 'preset',
      prst: g.attrs.prst,
      avLst: av ? ch(av, 'a:gd').map((x) => ({ name: x.attrs.name, fmla: x.attrs.fmla })) : [],
    };
  }
  const cg = c1(spPr, 'a:custGeom');
  if (!cg) return undefined;
  const lst = (t) => {
    const e = c1(cg, t);
    return e ? ch(e, 'a:gd').map((x) => ({ name: x.attrs.name, fmla: x.attrs.fmla })) : [];
  };
  const r = c1(cg, 'a:rect');
  const pl = c1(cg, 'a:pathLst');
  const paths = (pl ? ch(pl, 'a:path') : []).map((p) => ({
    attrs: p.attrs,
    cmds: p.children.map((c) => ({
      op: c.tag.slice(2),
      attrs: c.attrs,
      pts: ch(c, 'a:pt').map((q) => [q.attrs.x, q.attrs.y]),
    })),
  }));
  return {
    type: 'custom',
    avLst: lst('a:avLst'),
    gdLst: lst('a:gdLst'),
    rect: r && r.attrs,
    childTags: cg.children.map((c) => c.tag),
    paths,
  };
}
function shapeOf(el) {
  const kind = el.tag.slice(4);
  const nv = el.children.find((c) => c.tag.startsWith('xdr:nv'));
  const cNvPr = c1(nv, 'xdr:cNvPr');
  const o = { kind, name: cNvPr.attrs.name };
  if (kind === 'sp') {
    const c = c1(nv, 'xdr:cNvSpPr');
    if (c && c.attrs.txBox) o.txBox = c.attrs.txBox;
  }
  const spPr = c1(el, 'xdr:spPr');
  if (spPr) {
    o.xfrm = xfrmOf(spPr);
    o.geom = geomOf(spPr);
    o.ln = lnOf(spPr);
    o.fill = fillOf(spPr);
  }
  const st = c1(el, 'xdr:style');
  if (st) {
    const g = (t) => {
      const e = c1(st, t);
      return e ? num(e.attrs.idx) : undefined;
    };
    o.style = { lnRef: g('a:lnRef'), fillRef: g('a:fillRef') };
  }
  const tx = c1(el, 'xdr:txBody');
  if (tx)
    o.text = ch(tx, 'a:p')
      .map((p) =>
        p.children
          .filter((c) => c.tag === 'a:r')
          .map((r) => (c1(r, 'a:t') || { text: '' }).text)
          .join(''),
      )
      .join('\n');
  return o;
}
const shapesXml = {};
for (const a of tree.children[0].children) {
  if (!a.tag.endsWith('Anchor')) continue;
  const el = a.children.find((c) => ['xdr:sp', 'xdr:cxnSp'].includes(c.tag));
  if (!el) continue;
  const s = shapeOf(el);
  shapesXml[s.name] = s;
}

// ---- 結合 ----
const out = {
  excelVersion: com.excelVersion,
  notes: com.notes,
  unitNote: 'COM の Left/Top/Width/Height は pt（1pt=12700 EMU）。xml.xfrm は EMU。',
  shapes: {},
};
for (const s of com.shapes)
  out.shapes[s.Name] = { com: s, input: com.inputs[s.Name], xml: shapesXml[s.Name] };
for (const name of Object.keys(shapesXml))
  if (!out.shapes[name])
    out.shapes[name] = { xml: shapesXml[name], note: 'XML のみ（COM の最上位に無い）' };

// ---- U6：自由図形が w・h と整数の点で書かれているか。xdr:sp の自由図形に矢じりが付くか ----
const u6 = [];
for (const [name, v] of Object.entries(out.shapes)) {
  const g = v.xml?.geom;
  if (!g || g.type !== 'custom') continue;
  const ptsAreInt = g.paths.every((p) =>
    p.cmds.every((c) => c.pts.every((q) => q.every((n) => /^-?\d+$/.test(n)))),
  );
  u6.push({
    name,
    paths: g.paths.length,
    hasWH: g.paths.every((p) => p.attrs.w !== undefined && p.attrs.h !== undefined),
    ptsAreInt,
    ops: [...new Set(g.paths.flatMap((p) => p.cmds.map((c) => c.op)))],
    avLst: g.avLst.length,
    gdLst: g.gdLst.length,
    rect: g.rect ?? null,
    head: v.xml.ln?.headEnd?.type ?? null,
    tail: v.xml.ln?.tailEnd?.type ?? null,
  });
}
out.u6 = u6;
fs.writeFileSync(`${dir}/expected.json`, JSON.stringify(out, null, 1));
console.log(
  'shapes',
  Object.keys(out.shapes).length,
  'custom',
  u6.length,
  'xml-only',
  Object.keys(shapesXml).filter((n) => !out.shapes[n].com),
);
