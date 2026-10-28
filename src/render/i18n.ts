/**
 * M2-1 の仮置き：InterHub の文言関数 `t()` の代わり。キーをそのまま返す（InterHub の試験と同じ、文言の表が無い状態）。
 * M2-2 で `RenderOptions.messages` に置き換えて、このファイルを消す。
 */
export function t(key: string, _params?: Record<string, string | number>): string {
  return key;
}
