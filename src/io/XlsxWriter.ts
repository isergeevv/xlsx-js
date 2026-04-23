import { Workbook } from "../models/Workbook";
import type { SaveWorkbookOptions } from "../types";

export class XlsxWriter {
  public async write(workbook: Workbook, _options: SaveWorkbookOptions = {}): Promise<Uint8Array> {
    void workbook;
    throw new Error("XlsxWriter.write is not implemented yet");
  }
}
