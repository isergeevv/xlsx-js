import { Workbook } from "../models/Workbook";
import { XlsxParser } from "./XlsxParser";
import { XlsxWriter } from "./XlsxWriter";
import type { LoadWorkbookOptions, SaveWorkbookOptions, WorkbookInput } from "../types";

export class XlsxDocument {
  private readonly _parser: XlsxParser;
  private readonly _writer: XlsxWriter;

  constructor(parser = new XlsxParser(), writer = new XlsxWriter()) {
    this._parser = parser;
    this._writer = writer;
  }

  public createWorkbook(): Workbook {
    return new Workbook();
  }

  public async load(input: WorkbookInput, options: LoadWorkbookOptions = {}): Promise<Workbook> {
    return this._parser.parse(input, options);
  }

  public async save(workbook: Workbook, options: SaveWorkbookOptions = {}): Promise<Uint8Array> {
    return this._writer.write(workbook, options);
  }

  public async saveToPath(path: string, workbook: Workbook, options: SaveWorkbookOptions = {}): Promise<void> {
    return this._writer.writeToPath(path, workbook, options);
  }
}
