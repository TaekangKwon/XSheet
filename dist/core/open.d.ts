import type { XlsxLimits } from './limits';
/** 開いた zip（部品の読み出し口） */
export interface XlsxPackage {
    /** zip の中のファイルのパス（先頭に / を付けない） */
    has(path: string): boolean;
    /** 部品を文字（UTF-8）で読む。無ければ null。展開後の大きさが maxBytes を超えたら too-large */
    readText(path: string, maxBytes: number): Promise<string | null>;
    /** 部品をバイト列で読む（画像など）。無ければ null。展開しながら数え、maxBytes を超えたらそこで止めて too-large */
    readBytes(path: string, maxBytes: number): Promise<Uint8Array | null>;
    /** 読み込んだ zip を放す */
    dispose(): void;
}
/** Uint8Array へそろえる */
export declare function toBytes(data: ArrayBuffer | Uint8Array): Uint8Array;
/**
 * 先頭の印を確かめる。
 * - `PK\x03\x04` 以外は zip でない。
 * - OLE（`D0 CF 11 E0`）は、パスワード付きの xlsx（暗号化）か古い形式。どちらにしても読めないので encrypted。
 */
export declare function checkSignature(bytes: Uint8Array): void;
/** zip を開く。壊れていれば broken、ファイル数が多すぎれば too-large */
export declare function openPackage(bytes: Uint8Array, limits: XlsxLimits): Promise<XlsxPackage>;
