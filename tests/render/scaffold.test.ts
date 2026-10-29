// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';

// M0-1：雛形の確認。render のテストは jsdom 環境で走る。
describe('雛形（render）', () => {
  it('jsdom の document がある', () => {
    const div = document.createElement('div');
    div.className = 'xs-test';
    document.body.append(div);
    expect(document.querySelector('.xs-test')).toBe(div);
  });
});
