/**
 * `view`・`read` が受け取る xlsx の入れ物をバイト列にする（M6-2）。
 * URL の文字列は受けない（ライブラリは通信しない）。使う側が `fetch(url)` を渡せば、受け取った `Response` の中身を読むだけ。
 * 別の realm（iframe・jsdom）の物でも動くよう、`instanceof` ではなく形で見分ける。
 */
/** `view`・`read` が受け取る xlsx の入れ物 */
export type XSheetSource = Blob | ArrayBuffer | ArrayBufferView | Response | PromiseLike<Blob | ArrayBuffer | ArrayBufferView | Response>;
/**
 * 入れ物を Uint8Array にする。
 * @throws TypeError 受けない形（URL の文字列など）・失敗した `Response`（`ok` が false）
 */
export declare function toSourceBytes(source: XSheetSource): Promise<Uint8Array>;
