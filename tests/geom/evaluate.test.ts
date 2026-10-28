// 図形の形の評価器（guide の式・path 命令 → 経路・文字領域）のテスト（InterHub ISSUE-90 から写した）
import { describe, expect, it } from 'vitest';
import {
  MAX_GEOM_COMMANDS,
  MAX_GEOM_GUIDES,
  evaluateGeom,
  evaluateGuides,
  getPresetGeom,
  presetNames,
  type GeomDef,
  type GeomGuide,
  type GeomPath,
  type GeomResult,
} from '../../src/geom/evaluate';

const EMU = 9525;
const DEG = 60000;

/** guide だけの定義 */
const defOf = (gd: GeomGuide[], av?: [string, number][]): GeomDef => ({ av, gd, paths: [] });

/** guide 1 つを評価して値（EMU）を返す。失敗なら例外 */
function one(guide: GeomGuide, w = 100, h = 100, adjust?: Map<string, number>): number {
  const r = evaluateGuides(defOf([guide]), adjust, w, h);
  if (!r.ok) throw new Error(`評価に失敗: ${r.reason}`);
  return r.values.get(guide[0]) as number;
}

/** 経路だけの定義を評価する */
function paths(w: number, h: number, ...pathDefs: GeomDef['paths']): GeomResult {
  return evaluateGeom({ paths: pathDefs }, undefined, w, h);
}

function okPaths(r: GeomResult): GeomPath[] {
  if (!r.ok) throw new Error(`評価に失敗: ${r.reason}`);
  return r.paths;
}

describe('演算子 17 種', () => {
  it('*/ +- +/ ?:', () => {
    expect(one(['a', '*/', 3, 4, 2])).toBe(6);
    expect(one(['a', '*/', 3, 4, 0])).toBe(0); // ゼロ除算
    expect(one(['a', '+-', 10, 3, 2])).toBe(11);
    expect(one(['a', '+/', 10, 2, 4])).toBe(3);
    expect(one(['a', '+/', 10, 2, 0])).toBe(0);
    expect(one(['a', '?:', 1, 5, 9])).toBe(5);
    expect(one(['a', '?:', 0, 5, 9])).toBe(9); // 0 は偽
    expect(one(['a', '?:', -1, 5, 9])).toBe(9);
  });

  it('abs sqrt val max min mod pin', () => {
    expect(one(['a', 'abs', -5])).toBe(5);
    expect(one(['a', 'sqrt', 16])).toBe(4);
    expect(one(['a', 'sqrt', -4])).toBe(0); // 負の平方根は 0
    expect(one(['a', 'val', 7])).toBe(7);
    expect(one(['a', 'max', 3, 9])).toBe(9);
    expect(one(['a', 'min', 3, 9])).toBe(3);
    expect(one(['a', 'mod', 3, 4, 12])).toBe(13);
    expect(one(['a', 'pin', 0, 5, 10])).toBe(5);
    expect(one(['a', 'pin', 0, -5, 10])).toBe(0);
    expect(one(['a', 'pin', 0, 15, 10])).toBe(10);
  });

  it('sat2 cat2 sin cos tan at2（角度は 1/60000 度）', () => {
    expect(one(['a', 'sat2', 100, 1, 1])).toBeCloseTo(100 * Math.SQRT1_2, 9);
    expect(one(['a', 'cat2', 100, 1, 1])).toBeCloseTo(100 * Math.SQRT1_2, 9);
    expect(one(['a', 'sat2', 100, 0, 0])).toBe(0);
    expect(one(['a', 'sin', 100, 90 * DEG])).toBeCloseTo(100, 9);
    expect(one(['a', 'cos', 100, 0])).toBeCloseTo(100, 9);
    expect(one(['a', 'cos', 100, 90 * DEG])).toBeCloseTo(0, 9);
    expect(one(['a', 'tan', 100, 45 * DEG])).toBeCloseTo(100, 9);
    expect(one(['a', 'at2', 1, 1])).toBeCloseTo(45 * DEG, 6);
    expect(one(['a', 'at2', -1, 0])).toBeCloseTo(180 * DEG, 6);
    expect(one(['a', 'at2', 0, 0])).toBe(0);
  });

  it('sat2・cat2・at2 は象限をまたいでも符号が正しい（at2 の値域は (-180°, 180°]）', () => {
    expect(one(['a', 'cat2', 100, -1, 0])).toBeCloseTo(-100, 9);
    expect(one(['a', 'sat2', 100, -1, 0])).toBeCloseTo(0, 9);
    expect(one(['a', 'sat2', 100, 0, -1])).toBeCloseTo(-100, 9);
    expect(one(['a', 'cat2', 100, -1, -1])).toBeCloseTo(-100 * Math.SQRT1_2, 9);
    expect(one(['a', 'at2', -1, -1])).toBeCloseTo(-135 * DEG, 6);
    expect(one(['a', 'at2', 1, -1])).toBeCloseTo(-45 * DEG, 6);
    expect(one(['a', 'at2', -1, 1])).toBeCloseTo(135 * DEG, 6);
  });
});

describe('値の扱い', () => {
  it('有限でない結果は 0、各 guide は ±10^30 に丸める', () => {
    expect(one(['a', 'val', Infinity])).toBe(0);
    expect(one(['a', 'val', NaN])).toBe(0);
    expect(one(['a', '*/', 1e30, 1e30, 1])).toBe(1e30);
    expect(one(['a', '*/', -1e30, 1e30, 1])).toBe(-1e30);
    expect(one(['a', 'tan', 100, 90 * DEG])).toBeGreaterThan(1e17); // tan 90° は巨大（丸めには届かない）
  });

  it('EMU の 4 乗（約 10^20）の中間値も丸めずに保つ（円と直線の交点の判別式 r²d² − D² が 0 に潰れない。ISSUE-93）', () => {
    const r = evaluateGuides(
      defOf([
        ['q3', '*/', 'w', 'w', 1],
        ['q4', '*/', 'h', 'h', 1],
        ['q5', '*/', 'q3', 'q4', 1],
        ['q6', '*/', 'q3', 'q3', 1],
        ['q7', '+-', 'q5', 0, 'q6'],
      ]),
      undefined,
      100,
      50,
    );
    if (!r.ok) throw new Error(r.reason);
    expect(r.values.get('q5')).toBeGreaterThan(1e15);
    expect(r.values.get('q6')).toBeGreaterThan(1e15);
    expect(Math.abs(r.values.get('q7') as number)).toBeGreaterThan(1e15); // q5 ≠ q6（上限に張り付くと差が 0 になる）
  });

  it('path の座標・文字領域は ±10^15 に丸める（guide の巨大な値を描画側へ渡さない）', () => {
    const def: GeomDef = {
      gd: [
        ['a', '*/', 'w', 'w', 1],
        ['b', '*/', 'a', 'a', 1],
        ['huge', '*/', 'b', 'b', 1],
      ],
      paths: [
        {
          w: 100,
          h: 100,
          c: [
            ['M', 0, 0],
            ['L', 'huge', 'huge'],
          ],
        },
      ],
      rect: [0, 0, 'huge', 'huge'],
    };
    const r = evaluateGeom(def, undefined, 100, 100);
    if (!r.ok) throw new Error(r.reason);
    const seg = r.paths[0]!.subpaths[0]!.segs[0]!;
    expect(seg.t).toBe('L');
    if (seg.t === 'L') expect(seg.x).toBe(1e15);
    expect(r.textRect.r).toBeCloseTo(1e15 / EMU, 0);
  });

  it('単位：式は EMU で評価する（w px → w × 9525）', () => {
    expect(one(['a', 'val', 'w'], 200, 100)).toBe(200 * EMU);
    expect(one(['a', 'val', 'h'], 200, 100)).toBe(100 * EMU);
  });
});

describe('組み込み変数', () => {
  const vars = (w: number, h: number): Map<string, number> => {
    const names = [
      'w',
      'h',
      'l',
      't',
      'r',
      'b',
      'hc',
      'vc',
      'ss',
      'ls',
      'wd2',
      'wd3',
      'wd4',
      'wd5',
      'wd6',
      'wd8',
      'wd10',
      'wd12',
      'wd32',
      'hd2',
      'hd3',
      'hd4',
      'hd5',
      'hd6',
      'hd8',
      'ssd2',
      'ssd4',
      'ssd6',
      'ssd8',
      'ssd16',
      'ssd32',
      'cd2',
      'cd4',
      'cd8',
      '3cd4',
      '3cd8',
      '5cd8',
      '7cd8',
    ];
    const r = evaluateGuides(
      defOf(names.map((n): GeomGuide => ['g_' + n, 'val', n])),
      undefined,
      w,
      h,
    );
    if (!r.ok) throw new Error(r.reason);
    return new Map(names.map((n) => [n, r.values.get('g_' + n) as number]));
  };

  it('横長（200×100）', () => {
    const v = vars(200, 100);
    const W = 200 * EMU;
    const H = 100 * EMU;
    expect([v.get('l'), v.get('t'), v.get('r'), v.get('b')]).toEqual([0, 0, W, H]);
    expect(v.get('hc')).toBe(W / 2);
    expect(v.get('vc')).toBe(H / 2);
    expect(v.get('ss')).toBe(H);
    expect(v.get('ls')).toBe(W);
    expect(v.get('wd2')).toBe(W / 2);
    expect(v.get('wd32')).toBe(W / 32);
    expect(v.get('hd6')).toBe(H / 6);
    expect(v.get('ssd8')).toBe(H / 8);
  });

  it('縦長（100×200）・正方形（80×80）の ss・ls', () => {
    const t = vars(100, 200);
    expect(t.get('ss')).toBe(100 * EMU);
    expect(t.get('ls')).toBe(200 * EMU);
    expect(t.get('ssd16')).toBe((100 * EMU) / 16);
    const s = vars(80, 80);
    expect(s.get('ss')).toBe(80 * EMU);
    expect(s.get('ls')).toBe(80 * EMU);
  });

  it('角度の定数（1/60000 度）', () => {
    const v = vars(10, 10);
    expect(v.get('cd2')).toBe(180 * DEG);
    expect(v.get('cd4')).toBe(90 * DEG);
    expect(v.get('cd8')).toBe(45 * DEG);
    expect(v.get('3cd4')).toBe(270 * DEG);
    expect(v.get('3cd8')).toBe(135 * DEG);
    expect(v.get('5cd8')).toBe(225 * DEG);
    expect(v.get('7cd8')).toBe(315 * DEG);
  });
});

describe('評価の順と失敗', () => {
  it('avLst：既定を、渡された同じ名前の値で上書きする（av に無い名前は無視）', () => {
    const def = defOf([['x', '+-', 'adj', 1, 0]], [['adj', 100]]);
    const get = (adj?: Map<string, number>): number => {
      const r = evaluateGuides(def, adj, 10, 10);
      if (!r.ok) throw new Error(r.reason);
      return r.values.get('x') as number;
    };
    expect(get()).toBe(101);
    expect(get(new Map([['adj', 5]]))).toBe(6);
    expect(
      get(
        new Map([
          ['x', 5],
          ['w', 1],
        ]),
      ),
    ).toBe(101);
    // 有限でない調整値は既定値に戻す
    expect(get(new Map([['adj', NaN]]))).toBe(101);
    expect(get(new Map([['adj', Infinity]]))).toBe(101);
    expect(get(new Map([['adj', 0]]))).toBe(1);
  });

  it('gdLst は上から 1 回ずつ：前の guide を引ける', () => {
    const r = evaluateGuides(
      defOf([
        ['a', 'val', 3],
        ['b', '*/', 'a', 'a', 1],
      ]),
      undefined,
      10,
      10,
    );
    expect(r.ok && r.values.get('b')).toBe(9);
  });

  it('前方参照・自己参照・未知の名前は例外にせず失敗を返す', () => {
    const fail = (gd: GeomGuide[]): string => {
      const r = evaluateGuides(defOf(gd), undefined, 10, 10);
      return r.ok ? 'ok' : r.reason;
    };
    expect(
      fail([
        ['p', 'val', 'q'],
        ['q', 'val', 1],
      ]),
    ).toBe('name');
    expect(fail([['p', '+-', 'p', 1, 0]])).toBe('name');
    expect(fail([['a', 'val', 'nope']])).toBe('name');
  });

  it('未知の演算子・引数の数の誤りは失敗を返す', () => {
    const fail = (gd: GeomGuide[]): string => {
      const r = evaluateGuides(defOf(gd), undefined, 10, 10);
      return r.ok ? 'ok' : r.reason;
    };
    expect(fail([['a', 'frob', 1, 2, 3]])).toBe('op');
    expect(fail([['a', '*/', 1, 2]])).toBe('args');
    expect(fail([['a', 'abs', 1, 2]])).toBe('args');
    expect(fail([['a', 'val']])).toBe('args');
    expect(fail([['a', 'sin', 1, 2, 3]])).toBe('args');
  });

  it('guide・調整値の名前に __proto__・constructor があっても壊れない', () => {
    const r = evaluateGuides(
      defOf(
        [
          ['__proto__', 'val', 5],
          ['constructor', '+-', '__proto__', 1, 0],
        ],
        [['toString', 7]],
      ),
      new Map([['toString', 9]]),
      10,
      10,
    );
    expect(r.ok && r.values.get('__proto__')).toBe(5);
    expect(r.ok && r.values.get('constructor')).toBe(6);
    expect(r.ok && r.values.get('toString')).toBe(9);
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  });
});

describe('上限（guide 512・path 命令 4,096）', () => {
  const guides = (n: number): GeomGuide[] =>
    Array.from({ length: n }, (_, i): GeomGuide => ['g' + i, 'val', i]);
  const cmds = (n: number): GeomDef['paths'] => [
    { c: [['M', 0, 0], ...Array.from({ length: n - 1 }, () => ['L', 1, 1] as const)] },
  ];

  it('上限ちょうどは通り、1 つ超えると失敗（例外にならない）', () => {
    expect(MAX_GEOM_GUIDES).toBe(512);
    expect(MAX_GEOM_COMMANDS).toBe(4096);
    expect(evaluateGeom({ gd: guides(512), paths: [] }, undefined, 10, 10).ok).toBe(true);
    const over = evaluateGeom({ gd: guides(513), paths: [] }, undefined, 10, 10);
    expect(over).toEqual({ ok: false, reason: 'limit' });
    // 調整値（av）も guide の数に数える
    const av = evaluateGeom({ av: [['a', 1]], gd: guides(512), paths: [] }, undefined, 10, 10);
    expect(av).toEqual({ ok: false, reason: 'limit' });
    expect(evaluateGeom({ paths: cmds(4096) }, undefined, 10, 10).ok).toBe(true);
    expect(evaluateGeom({ paths: cmds(4097) }, undefined, 10, 10)).toEqual({
      ok: false,
      reason: 'limit',
    });
    // path が複数でも合計で数える
    expect(evaluateGeom({ paths: [...cmds(2048), ...cmds(2049)] }, undefined, 10, 10).ok).toBe(
      false,
    );
  });
});

describe('path 命令 → 経路', () => {
  it('moveTo・lnTo・close：座標系が無ければ EMU を px に戻す', () => {
    const [p] = okPaths(
      paths(100, 100, {
        c: [['M', EMU, 2 * EMU], ['L', 3 * EMU, 2 * EMU], ['L', 3 * EMU, 4 * EMU], ['Z']],
      }),
    );
    expect(p.fill).toBe('norm');
    expect(p.stroke).toBe(true);
    expect(p.subpaths).toEqual([
      {
        start: { x: 1, y: 2 },
        segs: [
          { t: 'L', x: 3, y: 2 },
          { t: 'L', x: 3, y: 4 },
        ],
        closed: true,
      },
    ]);
  });

  it('path の w・h があれば、その座標系から矩形の大きさへ線形に写す（0 は無指定）', () => {
    const [p] = okPaths(
      paths(200, 50, {
        w: 100,
        h: 100,
        c: [
          ['M', 10, 10],
          ['L', 50, 50],
        ],
      }),
    );
    expect(p.subpaths[0].start).toEqual({ x: 20, y: 5 });
    expect(p.subpaths[0].segs[0]).toEqual({ t: 'L', x: 100, y: 25 });
    const [q] = okPaths(paths(200, 50, { w: 0, h: 0, c: [['M', EMU, EMU]] }));
    expect(q.subpaths).toEqual([]); // 命令が M だけ → 区間なしで捨てる
    const [z] = okPaths(
      paths(200, 50, {
        w: 0,
        h: 0,
        c: [
          ['M', EMU, EMU],
          ['L', 2 * EMU, EMU],
        ],
      }),
    );
    expect(z.subpaths[0].start).toEqual({ x: 1, y: 1 });
  });

  it('cubicBezTo はそのまま、quadBezTo は誤差なく 3 次ベジェへ変換する', () => {
    const [c] = okPaths(
      paths(100, 100, {
        w: 100,
        h: 100,
        c: [
          ['M', 0, 0],
          ['C', 10, 20, 30, 40, 50, 60],
        ],
      }),
    );
    expect(c.subpaths[0].segs[0]).toEqual({ t: 'C', x1: 10, y1: 20, x2: 30, y2: 40, x: 50, y: 60 });

    const [q] = okPaths(
      paths(200, 50, {
        w: 100,
        h: 100,
        c: [
          ['M', 0, 0],
          ['Q', 50, 100, 100, 0],
        ],
      }),
    );
    const s = q.subpaths[0].segs[0];
    if (s.t !== 'C') throw new Error('C のはず');
    // 2 次: B(t) = (1-t)²P0 + 2(1-t)t Q + t²P1（拡大 sx=2・sy=0.5 を掛けた座標）
    for (const t of [0.1, 0.25, 0.5, 0.9]) {
      const u = 1 - t;
      const qx = (2 * u * t * 50 + t * t * 100) * 2;
      const qy = 2 * u * t * 100 * 0.5;
      const cx = u ** 3 * 0 + 3 * u * u * t * s.x1 + 3 * u * t * t * s.x2 + t ** 3 * s.x;
      const cy = u ** 3 * 0 + 3 * u * u * t * s.y1 + 3 * u * t * t * s.y2 + t ** 3 * s.y;
      expect(cx).toBeCloseTo(qx, 9);
      expect(cy).toBeCloseTo(qy, 9);
    }
  });

  it('close の後の lnTo は、閉じた部分経路の始点から新しい部分経路を始める。M だけの部分経路は捨てる', () => {
    const [p] = okPaths(
      paths(100, 100, {
        w: 100,
        h: 100,
        c: [['M', 0, 0], ['L', 10, 0], ['L', 10, 10], ['Z'], ['L', 5, 5], ['M', 7, 7]],
      }),
    );
    expect(p.subpaths).toHaveLength(2);
    expect(p.subpaths[0].closed).toBe(true);
    expect(p.subpaths[1]).toEqual({
      start: { x: 0, y: 0 },
      segs: [{ t: 'L', x: 5, y: 5 }],
      closed: false,
    });
  });

  it('path ごとの fill・stroke を持つ', () => {
    const r = okPaths(
      paths(
        10,
        10,
        {
          fill: 'lighten',
          stroke: false,
          c: [
            ['M', 0, 0],
            ['L', 1, 1],
          ],
        },
        {
          fill: 'none',
          c: [
            ['M', 0, 0],
            ['L', 1, 1],
          ],
        },
        {
          c: [
            ['M', 0, 0],
            ['L', 1, 1],
          ],
        },
      ),
    );
    expect(r.map((p) => [p.fill, p.stroke])).toEqual([
      ['lighten', false],
      ['none', true],
      ['norm', true],
    ]);
  });

  it('座標に組み込み変数・guide の名前が使える。未知の名前・未知の命令・引数不足は失敗', () => {
    const def: GeomDef = {
      gd: [['x1', '*/', 'w', 1, 4]],
      paths: [
        {
          c: [
            ['M', 'l', 'vc'],
            ['L', 'x1', 'b'],
          ],
        },
      ],
    };
    const r = evaluateGeom(def, undefined, 200, 100);
    expect(r.ok && r.paths[0].subpaths[0]).toEqual({
      start: { x: 0, y: 50 },
      segs: [{ t: 'L', x: 50, y: 100 }],
      closed: false,
    });
    const bad = (c: GeomDef['paths'][number]['c']): string => {
      const e = paths(10, 10, { c });
      return e.ok ? 'ok' : e.reason;
    };
    expect(bad([['M', 'nope', 0]])).toBe('name');
    expect(bad([['X', 0, 0]])).toBe('cmd');
    expect(bad([['M', 0]])).toBe('args');
    expect(bad([['C', 0, 0, 0, 0, 0]])).toBe('args');
  });
});

describe('arcTo', () => {
  /** 区間の列を細かく標本化した点（始点を含む） */
  function sample(p: GeomPath['subpaths'][number]): { x: number; y: number }[] {
    const pts = [p.start];
    let cur = p.start;
    for (const s of p.segs) {
      if (s.t === 'L') {
        pts.push({ x: s.x, y: s.y });
      } else {
        for (let i = 1; i <= 20; i++) {
          const t = i / 20;
          const u = 1 - t;
          pts.push({
            x: u ** 3 * cur.x + 3 * u * u * t * s.x1 + 3 * u * t * t * s.x2 + t ** 3 * s.x,
            y: u ** 3 * cur.y + 3 * u * u * t * s.y1 + 3 * u * t * t * s.y2 + t ** 3 * s.y,
          });
        }
      }
      cur = { x: s.x, y: s.y };
    }
    return pts;
  }

  /** 楕円（中心 cx,cy・半径 rx,ry）からの相対誤差の最大（正規化した半径 − 1） */
  const maxErr = (
    pts: { x: number; y: number }[],
    cx: number,
    cy: number,
    rx: number,
    ry: number,
  ) => Math.max(...pts.map((q) => Math.abs(Math.hypot((q.x - cx) / rx, (q.y - cy) / ry) - 1)));

  /** 視角 a（度）の楕円上の点 */
  const onEllipse = (cx: number, cy: number, rx: number, ry: number, aDeg: number) => {
    const a = (aDeg * Math.PI) / 180;
    const th = Math.atan2(rx * Math.sin(a), ry * Math.cos(a));
    return { x: cx + rx * Math.cos(th), y: cy + ry * Math.sin(th) };
  };

  const arc = (rx: number, ry: number, st: number, sw: number, cx = 100, cy = 100) => {
    const s = onEllipse(cx, cy, rx, ry, st);
    const r = paths(200, 200, {
      w: 200,
      h: 200,
      c: [
        ['M', s.x, s.y],
        ['A', rx, ry, st * DEG, sw * DEG],
      ],
    });
    return okPaths(r)[0].subpaths[0];
  };

  it('真円：360°・90°・負の振り角で半径の 0.03% 未満', () => {
    for (const [st, sw] of [
      [0, 360],
      [30, 90],
      [200, -270],
      [0, -360],
      [45, 180],
    ]) {
      const p = arc(50, 50, st, sw);
      expect(maxErr(sample(p), 100, 100, 50, 50)).toBeLessThan(3e-4);
    }
  });

  it('楕円：途中の角度（視角）・振り角 ±360° でも 0.03% 未満', () => {
    for (const [st, sw] of [
      [0, 360],
      [45, 270],
      [-30, -300],
      [100, 360],
      [10, 45],
      [300, -360],
    ]) {
      const p = arc(80, 30, st, sw);
      expect(maxErr(sample(p), 100, 100, 80, 30)).toBeLessThan(3e-4);
    }
  });

  it('終点は視角（中心から見た方向）で stAng + swAng の位置', () => {
    for (const [st, sw] of [
      [45, 270],
      [10, 60],
      [200, -100],
    ]) {
      const p = arc(80, 30, st, sw);
      const e = p.segs[p.segs.length - 1];
      const a = Math.atan2(e.y - 100, e.x - 100);
      const want = ((st + sw) * Math.PI) / 180;
      expect(Math.cos(a - want)).toBeCloseTo(1, 9);
    }
  });

  it('振り角 360° は始点に戻る', () => {
    const p = arc(80, 30, 45, 360);
    const e = p.segs[p.segs.length - 1];
    expect(e.x).toBeCloseTo(p.start.x, 9);
    expect(e.y).toBeCloseTo(p.start.y, 9);
    expect(p.segs).toHaveLength(4);
  });

  it('半径が 0 以下・振り角 0 は区間を作らない。巨大な振り角は 360° に丸める', () => {
    const none = (c: GeomDef['paths'][number]['c']): number =>
      okPaths(paths(200, 200, { w: 200, h: 200, c }))[0].subpaths.length;
    expect(
      none([
        ['M', 150, 100],
        ['A', 0, 50, 0, 90 * DEG],
      ]),
    ).toBe(0);
    expect(
      none([
        ['M', 150, 100],
        ['A', 50, -1, 0, 90 * DEG],
      ]),
    ).toBe(0);
    expect(
      none([
        ['M', 150, 100],
        ['A', 50, 50, 0, 0],
      ]),
    ).toBe(0);
    const big = okPaths(
      paths(200, 200, {
        w: 200,
        h: 200,
        c: [
          ['M', 150, 100],
          ['A', 50, 50, 0, 1e15],
        ],
      }),
    )[0];
    expect(big.subpaths[0].segs).toHaveLength(4);
  });

  it('座標系の拡大（path の w・h）をかけても弧は楕円の弧のまま', () => {
    const r = okPaths(
      paths(300, 100, {
        w: 100,
        h: 100,
        c: [
          ['M', 100, 50],
          ['A', 50, 50, 0, 360 * DEG],
        ],
      }),
    );
    // 中心 (150,50)・半径 (150,50) の楕円になる
    expect(maxErr(sample(r[0].subpaths[0]), 150, 50, 150, 50)).toBeLessThan(3e-4);
  });
});

describe('文字領域', () => {
  it('rect の l・t・r・b を評価して px で返す', () => {
    const r = evaluateGeom(
      { gd: [['x', '*/', 'w', 1, 4]], rect: ['x', 't', 'r', 'vc'], paths: [] },
      undefined,
      200,
      100,
    );
    expect(r.ok && r.textRect).toEqual({ l: 50, t: 0, r: 200, b: 50 });
  });

  it('rect が無い・幅か高さが 0 以下なら、回転前の矩形全体', () => {
    const full = { l: 0, t: 0, r: 200, b: 100 };
    expect(evaluateGeom({ paths: [] }, undefined, 200, 100)).toMatchObject({ textRect: full });
    expect(
      evaluateGeom({ rect: ['hc', 't', 'hc', 'b'], paths: [] }, undefined, 200, 100),
    ).toMatchObject({ textRect: full });
    expect(
      evaluateGeom({ rect: ['l', 'vc', 'r', 'vc'], paths: [] }, undefined, 200, 100),
    ).toMatchObject({ textRect: full });
    expect(
      evaluateGeom({ rect: ['r', 't', 'l', 'b'], paths: [] }, undefined, 200, 100),
    ).toMatchObject({ textRect: full });
  });

  it('rect の名前が未知なら失敗', () => {
    expect(evaluateGeom({ rect: ['l', 't', 'nope', 'b'], paths: [] }, undefined, 10, 10)).toEqual({
      ok: false,
      reason: 'name',
    });
  });
});

describe('preset の表（全形）', () => {
  const finite = (r: GeomResult): boolean => {
    if (!r.ok) return false;
    const nums: number[] = [r.textRect.l, r.textRect.t, r.textRect.r, r.textRect.b];
    for (const p of r.paths) {
      for (const s of p.subpaths) {
        nums.push(s.start.x, s.start.y);
        for (const g of s.segs) {
          if (g.t === 'L') nums.push(g.x, g.y);
          else nums.push(g.x1, g.y1, g.x2, g.y2, g.x, g.y);
        }
      }
    }
    return nums.every(Number.isFinite);
  };

  it('線の 10 種を除く全 preset が、既定の調整値と複数の縦横比で、例外なく有限の経路を返す', () => {
    const names = presetNames();
    expect(names.length).toBe(176);
    for (const name of names) {
      const def = getPresetGeom(name);
      expect(def, name).toBeDefined();
      for (const [w, h] of [
        [200, 200],
        [400, 100],
        [100, 400],
        [1, 1],
        [0, 0],
        [30000, 7],
      ]) {
        const r = evaluateGeom(def as GeomDef, undefined, w, h);
        expect(r.ok, `${name} ${w}x${h}: ${r.ok ? '' : r.reason}`).toBe(true);
        expect(finite(r), `${name} ${w}x${h}`).toBe(true);
        if (w >= 1 && r.ok) {
          expect(
            r.paths.some((p) => p.subpaths.length > 0),
            `${name} 経路なし`,
          ).toBe(true);
        }
      }
    }
  });

  it('全 preset が、極端な調整値（負・巨大）でも例外なく有限', () => {
    for (const name of presetNames()) {
      const def = getPresetGeom(name) as GeomDef;
      for (const v of [-1e12, -50000, 0, 1e12, 1e18]) {
        const adj = new Map((def.av ?? []).map(([k]) => [k, v] as const));
        const r = evaluateGeom(def, adj, 150, 90);
        expect(r.ok, name).toBe(true);
        expect(finite(r), `${name} adj=${v}`).toBe(true);
      }
    }
  });

  it('rightArrow（200×100）：矢じりの長さは ss と調整値で決まる', () => {
    const r = evaluateGeom(getPresetGeom('rightArrow') as GeomDef, undefined, 200, 100);
    if (!r.ok) throw new Error(r.reason);
    // a1=a2=50%：軸の太さ h/2、矢じりの長さ ss/2 → x1=150・y1=25・y2=75
    expect(r.paths).toHaveLength(1);
    expect(r.paths[0].subpaths[0]).toEqual({
      start: { x: 0, y: 25 },
      segs: [
        { t: 'L', x: 150, y: 25 },
        { t: 'L', x: 150, y: 0 },
        { t: 'L', x: 200, y: 50 },
        { t: 'L', x: 150, y: 100 },
        { t: 'L', x: 150, y: 75 },
        { t: 'L', x: 0, y: 75 },
      ],
      closed: true,
    });
    expect(r.textRect).toEqual({ l: 0, t: 25, r: 175, b: 75 });
    // 調整値の上書き：adj1=25000（軸が細くなる）・adj2=25000（矢じりが短くなる）
    const r2 = evaluateGeom(
      getPresetGeom('rightArrow') as GeomDef,
      new Map([
        ['adj1', 25000],
        ['adj2', 25000],
      ]),
      200,
      100,
    );
    if (!r2.ok) throw new Error(r2.reason);
    const s = r2.paths[0].subpaths[0];
    expect(s.start).toEqual({ x: 0, y: 37.5 });
    expect(s.segs[0]).toEqual({ t: 'L', x: 175, y: 37.5 });
  });

  it('roundRect（200×100）：角の丸みは ss の 16.667%', () => {
    const r = evaluateGeom(getPresetGeom('roundRect') as GeomDef, undefined, 200, 100);
    if (!r.ok) throw new Error(r.reason);
    const x1 = (100 * 16667) / 100000;
    const s = r.paths[0].subpaths[0];
    expect(s.closed).toBe(true);
    expect(s.start.x).toBeCloseTo(0, 9);
    expect(s.start.y).toBeCloseTo(x1, 9);
    expect(s.segs).toHaveLength(7); // 4 つの角（各 90° = 1 ベジェ）＋ 3 辺（4 辺目は close）
    const ends = s.segs.map((g) => [g.x, g.y]);
    expect(ends[0][0]).toBeCloseTo(x1, 9);
    expect(ends[0][1]).toBeCloseTo(0, 9);
    expect(ends[1][0]).toBeCloseTo(200 - x1, 9);
    expect(ends[1][1]).toBe(0);
    expect(ends[2][0]).toBeCloseTo(200, 9);
    expect(ends[2][1]).toBeCloseTo(x1, 9);
    expect(ends[5][0]).toBeCloseTo(x1, 9);
    expect(ends[5][1]).toBeCloseTo(100, 9);
    // 文字領域：角の丸みの分だけ内側（il = x1 × 0.29289）
    const il = (x1 * 29289) / 100000;
    expect(r.textRect.l).toBeCloseTo(il, 9);
    expect(r.textRect.r).toBeCloseTo(200 - il, 9);
    // 90° の弧の制御点：k = 4/3·tan(π/8) = 0.5523
    const c = s.segs[0];
    if (c.t !== 'C') throw new Error('C のはず');
    expect(c.x1).toBeCloseTo(0, 9);
    expect(c.y1).toBeCloseTo(x1 * (1 - 0.5522847498), 6);
  });

  it('ellipse（200×100）：4 つの弧・文字領域は内接矩形', () => {
    const r = evaluateGeom(getPresetGeom('ellipse') as GeomDef, undefined, 200, 100);
    if (!r.ok) throw new Error(r.reason);
    const s = r.paths[0].subpaths[0];
    expect(s.start).toEqual({ x: 0, y: 50 });
    expect(s.segs).toHaveLength(4);
    const ends = s.segs.map((g) => [g.x, g.y]);
    const want = [
      [100, 0],
      [200, 50],
      [100, 100],
      [0, 50],
    ];
    ends.forEach((e, i) => {
      expect(e[0]).toBeCloseTo(want[i][0], 9);
      expect(e[1]).toBeCloseTo(want[i][1], 9);
    });
    expect(r.textRect.l).toBeCloseTo(100 - 100 * Math.SQRT1_2, 9);
    expect(r.textRect.t).toBeCloseTo(50 - 50 * Math.SQRT1_2, 9);
  });

  it('diamond（200×100）：4 頂点・文字領域は (w/4,h/4)〜(3w/4,3h/4)', () => {
    const r = evaluateGeom(getPresetGeom('diamond') as GeomDef, undefined, 200, 100);
    if (!r.ok) throw new Error(r.reason);
    const s = r.paths[0].subpaths[0];
    expect([s.start, ...s.segs.map((g) => ({ x: g.x, y: g.y }))]).toEqual([
      { x: 0, y: 50 },
      { x: 100, y: 0 },
      { x: 200, y: 50 },
      { x: 100, y: 100 },
    ]);
    expect(r.textRect).toEqual({ l: 50, t: 25, r: 150, b: 75 });
  });

  it('can（100×200）：3 つの path（本体・上面の明るい面・上面の輪郭）', () => {
    const r = evaluateGeom(getPresetGeom('can') as GeomDef, undefined, 100, 200);
    if (!r.ok) throw new Error(r.reason);
    // adj=25% → y1 = ss·a/200000 = 12.5、y2 = 25、y3 = 187.5
    expect(r.paths.map((p) => [p.fill, p.stroke])).toEqual([
      ['norm', false],
      ['lighten', false],
      ['none', true],
    ]);
    expect(r.paths[0].subpaths[0].start).toEqual({ x: 0, y: 12.5 });
    expect(r.textRect).toEqual({ l: 0, t: 25, r: 100, b: 187.5 });
    // 本体の左辺から右辺の下の弧まで：下辺は y=200（=b）まで届く
    const ys = r.paths[0].subpaths[0].segs.map((g) => g.y);
    expect(Math.max(...ys)).toBeCloseTo(200, 9);
  });

  it('lineInv は図形として評価できる（左下→右上）', () => {
    const r = evaluateGeom(getPresetGeom('lineInv') as GeomDef, undefined, 80, 40);
    if (!r.ok) throw new Error(r.reason);
    expect(r.paths[0].subpaths[0].start).toEqual({ x: 0, y: 40 });
    expect(r.paths[0].subpaths[0].segs[0]).toEqual({ t: 'L', x: 80, y: 0 });
  });
});

describe('upArrow（原本に定義が無い。downArrow の上下反転として扱う）', () => {
  const pts = (r: GeomResult): { x: number; y: number }[] => {
    if (!r.ok) throw new Error(r.reason);
    const sp = r.paths[0].subpaths[0];
    return [sp.start, ...sp.segs.map((s) => ({ x: s.x, y: s.y }))];
  };

  it('原本の表には無く、getPresetGeom だけが返す（presetNames には含めない）。downArrow と同じ調整値・guide', () => {
    expect(presetNames()).not.toContain('upArrow');
    const up = getPresetGeom('upArrow') as GeomDef;
    const down = getPresetGeom('downArrow') as GeomDef;
    expect(up.flipY).toBe(true);
    expect(down.flipY).toBeUndefined();
    expect(up.av).toBe(down.av);
    expect(up.gd).toBe(down.gd);
    expect(getPresetGeom('upArrow')).toBe(up);
  });

  it('既定の調整値（100×200）：矢先が上（頂点 (50,0)）、羽根は y=50、軸の根元は y=200。ECMA-376 の upArrow の形', () => {
    const r = evaluateGeom(getPresetGeom('upArrow') as GeomDef, undefined, 100, 200);
    const near = (p: { x: number; y: number }[]): number[][] =>
      p.map((q) => [q.x, q.y].map((v) => Math.round(v * 1e6) / 1e6));
    expect(near(pts(r))).toEqual([
      [0, 50],
      [25, 50],
      [25, 200],
      [75, 200],
      [75, 50],
      [100, 50],
      [50, 0],
    ]);
    expect(r.ok && r.paths[0].subpaths[0].closed).toBe(true);
    // 文字領域は軸（x1〜x2）の、羽根の付け根（y2=25）から根元まで
    if (!r.ok) throw new Error(r.reason);
    expect(r.textRect.l).toBeCloseTo(25, 9);
    expect(r.textRect.t).toBeCloseTo(25, 9);
    expect(r.textRect.r).toBeCloseTo(75, 9);
    expect(r.textRect.b).toBeCloseTo(200, 9);
  });

  it('downArrow の経路を y → h − y に写したものと一致する（調整値を変えても。同じ avLst を渡す）', () => {
    for (const adj of [
      undefined,
      new Map([
        ['adj1', 20000],
        ['adj2', 80000],
      ]),
      new Map([
        ['adj1', 100000],
        ['adj2', 0],
      ]),
    ]) {
      const down = pts(evaluateGeom(getPresetGeom('downArrow') as GeomDef, adj, 120, 90));
      const up = pts(evaluateGeom(getPresetGeom('upArrow') as GeomDef, adj, 120, 90));
      expect(up).toHaveLength(down.length);
      down.forEach((p, i) => {
        expect(up[i].x).toBeCloseTo(p.x, 9);
        expect(up[i].y).toBeCloseTo(90 - p.y, 9);
      });
    }
  });

  it('有限で例外なし（極端な調整値・大きさでも）。幅・高さ 0 でも失敗しない', () => {
    const def = getPresetGeom('upArrow') as GeomDef;
    for (const v of [-1e12, 0, 50000, 1e18]) {
      const r = evaluateGeom(
        def,
        new Map([
          ['adj1', v],
          ['adj2', v],
        ]),
        150,
        90,
      );
      expect(r.ok).toBe(true);
      expect(pts(r).every((p) => Number.isFinite(p.x) && Number.isFinite(p.y))).toBe(true);
    }
    expect(evaluateGeom(def, undefined, 0, 0).ok).toBe(true);
  });

  it('flipY は経路のベジェの制御点も反転する（定義に flipY を付けた曲線）', () => {
    const def: GeomDef = {
      flipY: true,
      paths: [
        {
          w: 10,
          h: 10,
          c: [
            ['M', 0, 0],
            ['C', 2, 0, 4, 10, 10, 10],
          ],
        },
      ],
    };
    const r = evaluateGeom(def, undefined, 20, 40);
    if (!r.ok) throw new Error(r.reason);
    expect(r.paths[0].subpaths[0].start).toEqual({ x: 0, y: 40 });
    expect(r.paths[0].subpaths[0].segs[0]).toEqual({
      t: 'C',
      x1: 4,
      y1: 40,
      x2: 8,
      y2: 0,
      x: 20,
      y: 0,
    });
  });
});

describe('円弧矢印（circularArrow・leftCircularArrow・leftRightCircularArrow。ISSUE-93）', () => {
  /** 円と直線の交点の式は EMU の 4 乗の中間値を作る。丸め（±10^15）で判別式が 0 になると、矢じりの付け根の交点（F・C）が壊れる */
  const names = ['circularArrow', 'leftCircularArrow', 'leftRightCircularArrow'] as const;
  const sizes: [number, number][] = [
    [66.67, 53.33],
    [120, 80],
    [60, 100],
    [100, 100],
  ];

  /** guide の値（px） */
  function px(prst: string, w: number, h: number): (n: string) => number {
    const g = evaluateGuides(getPresetGeom(prst) as GeomDef, undefined, w, h);
    if (!g.ok) throw new Error(g.reason);
    return (n) => (g.values.get(n) as number) / EMU;
  }

  it('矢じりの頂点 A が弧（中心線の楕円 rw3×rh3）の上にあり、先頭の弧の終点 F が外側の楕円の上で、矢じりの付け根の線 B–G の上にある', () => {
    for (const prst of names) {
      for (const [w, h] of sizes) {
        const v = px(prst, w, h);
        const hc = w / 2;
        const vc = h / 2;
        const onEllipse = (x: number, y: number, rx: number, ry: number): number =>
          ((x - hc) / rx) ** 2 + ((y - vc) / ry) ** 2;
        const tag = `${prst} ${w}x${h}`;
        // 頂点 A は中心線の楕円の上（半径 rw3・rh3）
        expect(onEllipse(v('xA'), v('yA'), v('rw3'), v('rh3')), tag + ' A').toBeCloseTo(1, 6);
        // 外側の弧の終点 F は外側の楕円（rw1・rh1）の上
        expect(onEllipse(v('xF'), v('yF'), v('rw1'), v('rh1')), tag + ' F').toBeCloseTo(1, 6);
        // F は B–G の直線の上（外積が 0）
        const cross =
          (v('xG') - v('xB')) * (v('yF') - v('yB')) - (v('yG') - v('yB')) * (v('xF') - v('xB'));
        expect(Math.abs(cross), tag + ' B-G-F').toBeLessThan(1e-6 * Math.max(w, h) ** 2);
        // 内側の弧の始点 C は内側の楕円（rw2・rh2）の上
        expect(onEllipse(v('xC'), v('yC'), v('rw2'), v('rh2')), tag + ' C').toBeCloseTo(1, 6);
      }
    }
  });

  it('経路は矩形の中に収まり、閉じた部分経路 1 つ。先頭の弧は始点 E から F まで、矢じりは G′→A→B′ を通る', () => {
    for (const prst of names) {
      const [w, h] = [120, 80] as const;
      const r = evaluateGeom(getPresetGeom(prst) as GeomDef, undefined, w, h);
      if (!r.ok) throw new Error(r.reason);
      expect(r.paths).toHaveLength(1);
      const sp = r.paths[0]!.subpaths;
      expect(sp).toHaveLength(1);
      expect(sp[0]!.closed).toBe(true);
      const all = [
        sp[0]!.start,
        ...sp[0]!.segs.flatMap((s) =>
          s.t === 'L'
            ? [{ x: s.x, y: s.y }]
            : [
                { x: s.x1, y: s.y1 },
                { x: s.x2, y: s.y2 },
                { x: s.x, y: s.y },
              ],
        ),
      ];
      for (const p of all) {
        expect(p.x).toBeGreaterThanOrEqual(-1);
        expect(p.x).toBeLessThanOrEqual(w + 1);
        expect(p.y).toBeGreaterThanOrEqual(-1);
        expect(p.y).toBeLessThanOrEqual(h + 1);
      }
    }
  });

  it('circularArrow の既定：弧は左（180°）から上を回って右へ。矢じりの頂点は右端の中心線の上（ptAng 360°）、先頭の弧は付け根の線の手前（約 163°）で終わる', () => {
    const v = px('circularArrow', 66.67, 53.33);
    expect((v('stAng') * EMU) / DEG).toBeCloseTo(180, 6);
    expect((v('ptAng') * EMU) / DEG).toBeCloseTo(360, 4);
    // 外側の弧の振り角（先頭の弧）は 90° を大きく超える（壊れた評価では F が真上に来て 90° で終わっていた）
    expect((v('swAng') * EMU) / DEG).toBeGreaterThan(150);
    expect((v('swAng') * EMU) / DEG).toBeLessThan(170);
    // 矢じりの頂点は中心線の右端（hc + rw3, vc）
    expect(v('xA')).toBeCloseTo(66.67 / 2 + v('rw3'), 4);
    expect(v('yA')).toBeCloseTo(53.33 / 2, 2);
    // 付け根の線の端 G・B は頂点の少し上に水平に並ぶ（半幅 thh）
    expect(v('xG') - v('xB')).toBeCloseTo(2 * v('thh'), 4);
    expect(v('yG')).toBeCloseTo(v('yB'), 6);
    expect(v('yG')).toBeLessThan(v('yA'));
  });
});
