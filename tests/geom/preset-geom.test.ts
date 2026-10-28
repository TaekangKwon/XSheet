/// <reference types="node" />
// 形の表（生成物）と変換器のテスト（InterHub ISSUE-90 から写した）。評価器そのものは evaluate.test.ts
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
  EXCLUDED_PRESETS,
  convertPresetGeom,
  parseXml,
} from '../../tools/geom/build-preset-geom.mjs';
import { LINE_PRESETS } from '../../src/core/drawing-shape';
import { getPresetGeom, presetNames } from '../../src/geom/evaluate';

const read = (rel: string): string => readFileSync(new URL(rel, import.meta.url), 'utf8');
const XML = read('../../tools/geom/presetShapeDefinitions.xml');
const NOTICE = read('../../NOTICE');
const COMMITTED = read('../../src/geom/preset-geom.json');
const PACKAGE = JSON.parse(read('../../package.json')) as { files?: string[] };
const lf = (s: string): string => s.replace(/\r\n?/g, '\n');

const MINI = `<?xml version="1.0" encoding="utf-8"?>
<presetShapeDefinitons>
  <line>
    <rect l="l" t="t" r="r" b="b" />
    <pathLst><path fill="none"><moveTo><pt x="l" y="t" /></moveTo><lnTo><pt x="r" y="b" /></lnTo></path></pathLst>
  </line>
  <lineInv>
    <pathLst><path fill="none"><moveTo><pt x="l" y="b" /></moveTo><lnTo><pt x="r" y="t" /></lnTo></path></pathLst>
  </lineInv>
  <foo>
    <avLst><gd name="adj" fmla="val 25000" /></avLst>
    <gdLst><gd name="a" fmla="pin 0 adj 50000" /><gd name="b" fmla="abs a" /></gdLst>
    <ahLst><ahXY gdRefX="adj" minX="0" maxX="1"><pos x="a" y="t" /></ahXY></ahLst>
    <cxnLst><cxn ang="0"><pos x="r" y="vc" /></cxn></cxnLst>
    <rect l="l" t="t" r="b" b="b" />
    <pathLst>
      <path w="10" h="20" stroke="false" fill="darken" extrusionOk="false">
        <moveTo><pt x="l" y="0" /></moveTo>
        <arcTo wR="wd2" hR="hd2" stAng="cd2" swAng="-5400000" />
        <quadBezTo><pt x="1" y="2" /><pt x="3" y="4" /></quadBezTo>
        <cubicBezTo><pt x="1" y="2" /><pt x="3" y="4" /><pt x="5" y="6" /></cubicBezTo>
        <close />
      </path>
    </pathLst>
  </foo>
  <bar><pathLst><path><moveTo><pt x="0" y="0" /></moveTo></path></pathLst></bar>
</presetShapeDefinitons>`;

describe('変換器（build-preset-geom.mjs）', () => {
  it('小さな XML：ahLst・cxnLst を捨て、線の preset を除き、lineInv は残す', () => {
    const out = JSON.parse(convertPresetGeom(MINI, 'N'));
    expect(Object.keys(out.shapes)).toEqual(['lineInv', 'foo', 'bar']);
    expect(out.shapes.foo).toEqual({
      av: [['adj', 25000]],
      gd: [
        ['a', 'pin', 0, 'adj', 50000],
        ['b', 'abs', 'a'],
      ],
      rect: ['l', 't', 'b', 'b'],
      paths: [
        {
          w: 10,
          h: 20,
          fill: 'darken',
          stroke: false,
          c: [
            ['M', 'l', 0],
            ['A', 'wd2', 'hd2', 'cd2', -5400000],
            ['Q', 1, 2, 3, 4],
            ['C', 1, 2, 3, 4, 5, 6],
            ['Z'],
          ],
        },
      ],
    });
    expect(out.shapes.lineInv).toEqual({
      paths: [
        {
          fill: 'none',
          c: [
            ['M', 'l', 'b'],
            ['L', 'r', 't'],
          ],
        },
      ],
    });
    expect(out._notice).toBe('N');
    expect(out.version).toBe(1);
  });

  it('許可リストで形を絞れる', () => {
    const out = JSON.parse(convertPresetGeom(MINI, 'N', { allow: ['foo'] }));
    expect(Object.keys(out.shapes)).toEqual(['foo']);
  });

  it('同名で内容が違う形は止める・同じ内容の重複は 1 つにする', () => {
    const dup = (b: string): string =>
      `<r><x><pathLst><path><moveTo><pt x="0" y="0" /></moveTo></path></pathLst></x>` +
      `<x><pathLst><path><moveTo><pt x="${b}" y="0" /></moveTo></path></pathLst></x></r>`;
    expect(Object.keys(JSON.parse(convertPresetGeom(dup('0'), 'N')).shapes)).toEqual(['x']);
    expect(() => convertPresetGeom(dup('1'), 'N')).toThrow();
  });

  it('XML の簡易パーサ：BOM・属性・自己閉じ・入れ子', () => {
    const root = parseXml('<a><b k="v" /><c><d x="1"></d></c></a>');
    expect(root.name).toBe('a');
    expect(root.children.map((n) => n.name)).toEqual(['b', 'c']);
    expect(root.children[0].attrs.k).toBe('v');
    expect(root.children[1].children[0].attrs.x).toBe('1');
  });

  it('出力は決定的で、同梱した JSON は原本の XML からの再生成と一致する（改行コードを除く）', () => {
    const a = convertPresetGeom(XML, NOTICE);
    const b = convertPresetGeom(XML, NOTICE);
    expect(a).toBe(b);
    expect(lf(COMMITTED)).toBe(lf(a));
  });
});

describe('形の表（preset-geom.json）', () => {
  const table = JSON.parse(COMMITTED) as {
    _source: string;
    _notice: string;
    shapes: Record<string, Record<string, unknown>>;
  };

  it('Ecma の出典と著作権表示を含む', () => {
    expect(table._source).toContain('Ecma International');
    expect(table._notice).toContain('COPYRIGHT NOTICE');
    expect(table._notice).toContain('Ecma International');
    expect(table._notice).toContain('Copyright License and Disclaimer');
  });

  it('著作権表示の NOTICE は配布物に入る（package.json の files）', () => {
    expect(PACKAGE.files).toContain('NOTICE');
    expect(NOTICE).toContain('COPYRIGHT NOTICE');
    expect(lf(table._notice)).toBe(lf(NOTICE).trimEnd());
  });

  it('除外の 10 種は core の LINE_PRESETS と一致する', () => {
    // 変換器の除外 10 種は、線として別経路で描く LINE_PRESETS（core の drawing-shape.ts）と一致する
    expect(new Set(EXCLUDED_PRESETS)).toEqual(new Set(LINE_PRESETS));
  });

  it('ahLst・cxnLst と線の 10 種を含まず、lineInv を含む', () => {
    expect(EXCLUDED_PRESETS).toHaveLength(10);
    for (const p of EXCLUDED_PRESETS) expect(p in table.shapes).toBe(false);
    expect('lineInv' in table.shapes).toBe(true);
    expect(COMMITTED).not.toContain('ahLst');
    expect(COMMITTED).not.toContain('cxnLst');
    for (const s of Object.values(table.shapes)) {
      expect(Object.keys(s).every((k) => ['av', 'gd', 'rect', 'paths'].includes(k))).toBe(true);
    }
  });

  it('getPresetGeom は Map で引く：既知の形は返り、プロトタイプ由来の名前は返らない', () => {
    expect(getPresetGeom('rightArrow')?.paths.length).toBeGreaterThan(0);
    expect(presetNames()).toHaveLength(Object.keys(table.shapes).length);
    for (const bad of ['__proto__', 'constructor', 'toString', 'hasOwnProperty', 'line', '']) {
      expect(getPresetGeom(bad)).toBeUndefined();
    }
  });
});
