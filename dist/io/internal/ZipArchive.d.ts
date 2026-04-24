export interface ZipEntry {
    name: string;
    data: Uint8Array;
}
export declare function writeZip(entries: ZipEntry[]): Promise<Uint8Array>;
export declare function readZip(buffer: Uint8Array): Promise<Map<string, Uint8Array>>;
export declare function encodeText(value: string): Uint8Array;
export declare function decodeText(value: Uint8Array): string;
//# sourceMappingURL=ZipArchive.d.ts.map