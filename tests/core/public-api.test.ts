import { readFileSync } from 'node:fs';
import { describe, expect, expectTypeOf, it } from 'vitest';
import * as core from '../../src/core/index';
import * as internal from '../../src/core/internal';
import type {
  ChartModel,
  ChartResult,
  DrawingColor,
  DrawingGroup,
  DrawingLine,
  DrawingNotice,
  DrawingText,
  Geometry,
  ImageResult,
  OpenXlsxOptions,
  Rgba,
  ShapeGeom,
  ShapeOutline,
  SheetDrawing,
  SheetDrawings,
  SheetLayout,
  SheetRow,
  XlsxBook,
  XlsxErrorKind,
  XlsxLimits,
  XlsxSheet,
} from '../../src/core/index';

// M1-3：公開の範囲（docs/02_architecture.md §2）。入口から出すものが増減したら、このテストと文書を合わせて直す
describe('公開の入口（xsheet）', () => {
  it('値の公開はこれだけ', () => {
    expect(Object.keys(core).sort()).toEqual(
      [
        'DEFAULT_LIMITS',
        'XlsxError',
        'colName',
        'colWidthToPx',
        'createGeometry',
        'defaultColWidthToPx',
        'findText',
        'fontDefaultRowPx',
        'isPaperFill',
        'lookupMdw',
        'openXlsx',
        'ptToPx',
      ].sort(),
    );
  });

  it('types.ts の公開の型はすべて入口から出す（図の型も openDrawings の戻り値から辿れる）', () => {
    const read = (p: string) => readFileSync(new URL(p, import.meta.url), 'utf8');
    const declared = [
      ...read('../../src/core/types.ts').matchAll(/^export (?:interface|type) (\w+)/gm),
    ].map((m) => m[1]);
    const index = read('../../src/core/index.ts');
    const block = /export type \{([^}]*)\} from '\.\/types';/.exec(index)?.[1] ?? '';
    const exported = block
      .split(',')
      .map((s) => s.trim())
      .filter((s) => s !== '');
    expect(declared.length).toBeGreaterThan(40);
    expect(exported.sort()).toEqual(declared.sort());
  });

  it('関数と型のつながり（型の検査は npm run typecheck で走る）', () => {
    expectTypeOf(core.openXlsx).parameters.toEqualTypeOf<
      [data: ArrayBuffer | Uint8Array, options?: OpenXlsxOptions]
    >();
    expectTypeOf(core.openXlsx).returns.resolves.toEqualTypeOf<XlsxBook>();
    expectTypeOf<XlsxBook['openSheet']>().returns.resolves.toEqualTypeOf<XlsxSheet>();
    expectTypeOf<XlsxBook['openDrawings']>().returns.resolves.toEqualTypeOf<SheetDrawings>();
    expectTypeOf<XlsxBook['readImage']>().returns.resolves.toEqualTypeOf<ImageResult>();
    expectTypeOf<XlsxBook['readChart']>().returns.resolves.toEqualTypeOf<ChartResult>();
    expectTypeOf<Extract<ChartResult, { ok: true }>['chart']>().toEqualTypeOf<ChartModel>();
    expectTypeOf<XlsxSheet['getRows']>().returns.toEqualTypeOf<SheetRow[]>();
    expectTypeOf<XlsxSheet['layout']>().toEqualTypeOf<SheetLayout>();
    expectTypeOf<SheetDrawings['notices']>().toEqualTypeOf<readonly DrawingNotice[]>();
    expectTypeOf<SheetDrawing['text']>().toEqualTypeOf<DrawingText | undefined>();
    expectTypeOf<SheetDrawing['line']>().toEqualTypeOf<DrawingLine | undefined>();
    expectTypeOf<SheetDrawing['geom']>().toEqualTypeOf<ShapeGeom | undefined>();
    expectTypeOf<SheetDrawing['outline']>().toEqualTypeOf<ShapeOutline | undefined>();
    expectTypeOf<SheetDrawing['fillColor']>().toEqualTypeOf<DrawingColor | undefined>();
    expectTypeOf<SheetDrawing['group']>().toEqualTypeOf<DrawingGroup | undefined>();
    expectTypeOf(core.createGeometry).returns.toEqualTypeOf<Geometry>();
    expectTypeOf(core.DEFAULT_LIMITS).toEqualTypeOf<Readonly<XlsxLimits>>();
    expectTypeOf<core.XlsxError['kind']>().toEqualTypeOf<XlsxErrorKind>();
    expectTypeOf<Rgba>().toEqualTypeOf<{ rgb: number; alpha: number }>();
  });

  it('上限の既定値は凍結されていて、使う側が書き換えられない', () => {
    expect(Object.isFrozen(core.DEFAULT_LIMITS)).toBe(true);
  });
});

describe('半公開の入口（xsheet/internal）', () => {
  it('render が使う部品だけを出す（版の互換は約束しない）', () => {
    expect(Object.keys(internal).sort()).toEqual([
      'DEFAULT_MDW',
      'DEFAULT_THEME',
      'applyFillMode',
      'applyTransforms',
      'emptyLayout',
      'formatNumberByCode',
    ]);
  });

  it('公開の入口と名前が重ならない', () => {
    for (const name of Object.keys(internal)) expect(core).not.toHaveProperty(name);
  });
});
