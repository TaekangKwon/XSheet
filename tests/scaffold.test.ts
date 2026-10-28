import { describe, expect, it } from 'vitest';

// M0-1：雛形の確認。各層の入口が読み込めることと、core のテストが node 環境で走ること。
describe('雛形', () => {
  it('core・geom・render の入口を読み込める', async () => {
    await expect(import('../src/core/index')).resolves.toBeDefined();
    await expect(import('../src/geom/index')).resolves.toBeDefined();
    await expect(import('../src/render/index')).resolves.toBeDefined();
  });

  it('core のテストは DOM の無い環境で走る', () => {
    expect(typeof document).toBe('undefined');
  });
});
