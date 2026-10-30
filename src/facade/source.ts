// 正: docs/adr/0003-universal-distribution.md §2
/**
 * `view`・`read` が受け取る xlsx の入れ物をバイト列にする（M6-2）。
 * URL の文字列は受けない（ライブラリは通信しない）。使う側が `fetch(url)` を渡せば、受け取った `Response` の中身を読むだけ。
 * 別の realm（iframe・jsdom）の物でも動くよう、`instanceof` ではなく形で見分ける。
 */

/** `view`・`read` が受け取る xlsx の入れ物 */
export type XSheetSource =
  | Blob
  | ArrayBuffer
  | ArrayBufferView
  | Response
  | PromiseLike<Blob | ArrayBuffer | ArrayBufferView | Response>;

/** `arrayBuffer()` を持つもの（Blob・File・Response） */
interface BufferLike {
  arrayBuffer(): Promise<ArrayBuffer>;
  ok?: unknown;
  status?: unknown;
}

const isBufferLike = (v: unknown): v is BufferLike =>
  typeof v === 'object' && v !== null && typeof (v as BufferLike).arrayBuffer === 'function';

const isArrayBuffer = (v: unknown): v is ArrayBuffer =>
  Object.prototype.toString.call(v) === '[object ArrayBuffer]';

/**
 * 入れ物を Uint8Array にする。
 * @throws TypeError 受けない形（URL の文字列など）・失敗した `Response`（`ok` が false）
 */
export async function toSourceBytes(source: XSheetSource): Promise<Uint8Array> {
  const value: unknown = await source;
  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  }
  if (isArrayBuffer(value)) return new Uint8Array(value);
  if (isBufferLike(value)) {
    // Response：失敗した応答（404 の HTML など）を xlsx として読まない
    if (value.ok === false) {
      throw new TypeError(`XSheet: the response failed (HTTP ${String(value.status)})`);
    }
    return new Uint8Array(await value.arrayBuffer());
  }
  if (typeof value === 'string') {
    throw new TypeError(
      'XSheet: a URL string is not accepted (XSheet never fetches). Pass fetch(url) or a File/Blob/ArrayBuffer instead',
    );
  }
  throw new TypeError(
    'XSheet: the source must be a File, Blob, ArrayBuffer, typed array or Response',
  );
}
