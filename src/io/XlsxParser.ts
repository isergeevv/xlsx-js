import { Workbook } from "../models/Workbook";
import type { LoadWorkbookOptions } from "../types";

export class XlsxParser {
  public async parse(buffer: Uint8Array, _options: LoadWorkbookOptions = {}): Promise<Workbook> {
    void buffer;
    throw new Error("XlsxParser.parse is not implemented yet");
  }
}
