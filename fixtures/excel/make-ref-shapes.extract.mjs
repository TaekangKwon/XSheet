// drawing XML → 図形ごとの値、COM の値（com.json）と結合して expected.json を作る
/* global process */
import fs from 'node:fs';
const [, , dir] = process.argv;
const xml = fs.readFileSync(`${dir}/x/xl/drawings/drawing1.xml`, 'utf8');
const com = JSON.parse(fs.readFileSync(`${dir}/out/com.json`, 'utf8').replace(/^\uFEFF/, ''));

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
  const o = { w: num(ln.attrs.w), cap: ln.attrs.cap };
  const kinds = ln.children.map((c) => c.tag);
  o.children = kinds;
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
function geomOf(spPr) {
  const g = c1(spPr, 'a:prstGeom');
  if (!g) return undefined;
  const av = c1(g, 'a:avLst');
  const gd = av ? ch(av, 'a:gd').map((x) => ({ name: x.attrs.name, fmla: x.attrs.fmla })) : [];
  return { prst: g.attrs.prst, avLst: gd };
}
function txOf(tx) {
  if (!tx) return undefined;
  const bp = c1(tx, 'a:bodyPr');
  const o = {
    bodyPr: pick(bp.attrs, [
      'vert',
      'wrap',
      'lIns',
      'tIns',
      'rIns',
      'bIns',
      'anchor',
      'rot',
      'anchorCtr',
      'vertOverflow',
      'horzOverflow',
      'rtlCol',
      'numCol',
    ]),
  };
  o.bodyPrChildren = bp.children.map((c) => ({ tag: c.tag, ...c.attrs }));
  o.paragraphs = ch(tx, 'a:p').map((p) => {
    const pPr = c1(p, 'a:pPr');
    const pp = {};
    if (pPr) {
      pp.pPr = pick(pPr.attrs, ['algn', 'marL', 'indent', 'lvl']);
      for (const c of pPr.children) {
        if (c.tag === 'a:buChar') pp.buChar = c.attrs.char;
        else if (c.tag === 'a:buAutoNum') pp.buAutoNum = pick(c.attrs, ['type', 'startAt']);
        else if (c.tag === 'a:buNone') pp.buNone = true;
        else if (c.tag === 'a:buFont') pp.buFont = c.attrs.typeface;
      }
    }
    pp.items = [];
    for (const c of p.children) {
      if (c.tag === 'a:r' || c.tag === 'a:fld') {
        const rPr = c1(c, 'a:rPr');
        const t = c1(c, 'a:t');
        const it = {
          type: c.tag === 'a:r' ? 'r' : 'fld',
          text: t ? t.text : '',
          rPr: rPr
            ? pick(rPr.attrs, ['sz', 'b', 'i', 'u', 'strike', 'baseline', 'lang'])
            : undefined,
        };
        if (c.tag === 'a:fld') {
          it.fldType = c.attrs.type;
          it.fldId = c.attrs.id;
        }
        pp.items.push(it);
      } else if (c.tag === 'a:br') {
        const rPr = c1(c, 'a:rPr');
        pp.items.push({
          type: 'br',
          rPr: rPr ? pick(rPr.attrs, ['sz', 'b', 'i', 'u', 'strike']) : undefined,
        });
      } else if (c.tag === 'a:endParaRPr') pp.endParaRPr = pick(c.attrs, ['sz', 'b', 'i']);
    }
    return pp;
  });
  return o;
}
function styleOf(st) {
  if (!st) return undefined;
  const g = (t) => {
    const e = c1(st, t);
    return e ? num(e.attrs.idx) : undefined;
  };
  return { lnRef: g('a:lnRef'), fillRef: g('a:fillRef'), effectRef: g('a:effectRef') };
}
function shapeOf(el) {
  const kind = el.tag.slice(4);
  const nv = el.children.find((c) => c.tag.startsWith('xdr:nv'));
  const cNvPr = c1(nv, 'xdr:cNvPr');
  const o = { kind, id: num(cNvPr.attrs.id), name: cNvPr.attrs.name };
  if (cNvPr.attrs.hidden) o.hidden = cNvPr.attrs.hidden;
  if (el.attrs.textlink) o.textlink = el.attrs.textlink;
  if (el.attrs.macro !== undefined && el.attrs.macro !== '') o.macro = el.attrs.macro;
  if (kind === 'sp') {
    const c = c1(nv, 'xdr:cNvSpPr');
    if (c && c.attrs.txBox) o.txBox = c.attrs.txBox;
  }
  const cnx = c1(nv, 'xdr:cNvCxnSpPr');
  if (cnx) {
    const st = c1(cnx, 'a:stCxn'),
      en = c1(cnx, 'a:endCxn');
    if (st) o.stCxn = st.attrs;
    if (en) o.endCxn = en.attrs;
  }
  if (kind === 'grpSp') {
    const gp = c1(el, 'xdr:grpSpPr');
    const x = c1(gp, 'a:xfrm');
    const g = (t, a, b) => {
      const e = c1(x, t);
      return e && { [a]: num(e.attrs[a]), [b]: num(e.attrs[b]) };
    };
    o.grpSpPr = {
      off: g('a:off', 'x', 'y'),
      ext: g('a:ext', 'cx', 'cy'),
      chOff: g('a:chOff', 'x', 'y'),
      chExt: g('a:chExt', 'cx', 'cy'),
      rot: num(x.attrs.rot),
      flipH: x.attrs.flipH,
      flipV: x.attrs.flipV,
    };
    o.children = el.children
      .filter((c) =>
        ['xdr:sp', 'xdr:cxnSp', 'xdr:grpSp', 'xdr:pic', 'xdr:graphicFrame'].includes(c.tag),
      )
      .map(shapeOf);
    return o;
  }
  if (kind === 'graphicFrame') {
    const x = c1(el, 'xdr:xfrm');
    if (x) {
      const off = c1(x, 'a:off'),
        ext = c1(x, 'a:ext');
      o.xfrm = { off: off.attrs, ext: ext.attrs };
    }
    const g = c1(c1(el, 'a:graphic'), 'a:graphicData');
    if (g) o.graphicDataUri = g.attrs.uri;
    return o;
  }
  const spPr = c1(el, 'xdr:spPr');
  if (spPr) {
    o.xfrm = xfrmOf(spPr);
    o.geom = geomOf(spPr);
    o.ln = lnOf(spPr);
    o.fill = fillOf(spPr);
    o.spPrChildren = spPr.children.map((c) => c.tag);
  }
  o.style = styleOf(c1(el, 'xdr:style'));
  const tx = txOf(c1(el, 'xdr:txBody'));
  if (tx) o.txBody = tx;
  return o;
}
const shapesXml = {};
for (const a of tree.children[0].children) {
  if (!a.tag.endsWith('Anchor')) continue;
  const info = { anchor: a.tag.slice(4) };
  if (a.attrs.editAs) info.editAs = a.attrs.editAs;
  for (const k of ['from', 'to']) {
    const e = c1(a, 'xdr:' + k);
    if (e)
      info[k] = {
        col: num(c1(e, 'xdr:col').text),
        colOff: num(c1(e, 'xdr:colOff').text),
        row: num(c1(e, 'xdr:row').text),
        rowOff: num(c1(e, 'xdr:rowOff').text),
      };
  }
  let el = a.children.find((c) =>
    [
      'xdr:sp',
      'xdr:cxnSp',
      'xdr:grpSp',
      'xdr:pic',
      'xdr:graphicFrame',
      'mc:AlternateContent',
    ].includes(c.tag),
  );
  if (el.tag === 'mc:AlternateContent') {
    const choice = c1(el, 'mc:Choice'),
      fb = c1(el, 'mc:Fallback');
    info.alternateContent = {
      choiceRequires: choice.attrs.Requires,
      choiceChildren: choice.children.map((c) => c.tag),
      fallbackChildren: fb.children.map((c) => c.tag),
    };
    const cEl = choice.children[0];
    const s = shapeOf(cEl);
    s.alternateFallback = shapeOf(fb.children[0]);
    Object.assign(s, info);
    shapesXml[s.name] = s;
    continue;
  }
  const s = shapeOf(el);
  Object.assign(s, info);
  shapesXml[s.name] = s;
}

// ---- 結合 ----
const out = {
  excelVersion: com.excelVersion,
  notes: com.notes,
  unitNote:
    'COM の Left/Top/Width/Height は pt（1pt=12700 EMU）。xml.xfrm は EMU。グループの子の COM 値は絶対座標。',
  shapes: {},
};
const comByName = {};
for (const s of com.shapes) comByName[s.Name] = s;
for (const name of Object.keys({ ...comByName })) {
  out.shapes[name] = { com: comByName[name], input: com.inputs[name], xml: shapesXml[name] };
}
for (const name of Object.keys(shapesXml))
  if (!out.shapes[name])
    out.shapes[name] = { xml: shapesXml[name], note: 'XML のみ（COM の最上位に無い）' };

// ---- U1: rot/flip の一致 ----
const u1 = [];
for (const [name, v] of Object.entries(out.shapes)) {
  const x = v.xml?.xfrm || v.xml?.grpSpPr,
    c = v.com;
  if (!x || !c) continue;
  const xrot = (x.rot ?? 0) / 60000;
  const xh = x.flipH === '1',
    xv = x.flipV === '1';
  const comRot = c.Rotation,
    ch = !!c.HorizontalFlip,
    cv = !!c.VerticalFlip;
  u1.push({
    name,
    xmlRotDeg: xrot,
    xmlFlipH: xh,
    xmlFlipV: xv,
    comRot,
    comFlipH: ch,
    comFlipV: cv,
    match: Math.abs(xrot - comRot) < 0.01 && xh === ch && xv === cv,
  });
}
out.u1 = u1;
fs.writeFileSync(`${dir}/expected.json`, JSON.stringify(out, null, 1));
const mism = u1.filter((r) => !r.match);
console.log(
  'shapes',
  Object.keys(out.shapes).length,
  'xml-only',
  Object.keys(shapesXml).filter((n) => !comByName[n]),
  'u1 rows',
  u1.length,
  'mismatch',
  mism.length,
);
console.log(JSON.stringify(mism, null, 0));
