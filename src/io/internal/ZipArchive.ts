import JSZip from "jszip";

const _textEncoder = new TextEncoder();
const _textDecoder = new TextDecoder();

export interface ZipEntry {
  name: string;
  data: Uint8Array;
}

export async function writeZip(entries: ZipEntry[]): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const entry of entries) {
    zip.file(entry.name, entry.data);
  }
  const data = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 }
  });
  return data;
}

export async function readZip(buffer: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(buffer);
  const out = new Map<string, Uint8Array>();
  const paths = Object.keys(zip.files);
  for (const path of paths) {
    const file = zip.files[path];
    if (!file || file.dir) {
      continue;
    }
    out.set(path, await file.async("uint8array"));
  }
  return out;
}

export function encodeText(value: string): Uint8Array {
  return _textEncoder.encode(value);
}

export function decodeText(value: Uint8Array): string {
  return _textDecoder.decode(value);
}
