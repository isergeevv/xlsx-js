import { Workbook } from "../models/Workbook";
import { XlsxParser } from "./XlsxParser";
import { XlsxWriter } from "./XlsxWriter";
import type { LoadWorkbookOptions, SaveWorkbookOptions, WorkbookInput } from "../types";
export declare class XlsxDocument {
    private readonly _parser;
    private readonly _writer;
    constructor(parser?: XlsxParser, writer?: XlsxWriter);
    createWorkbook(): Workbook;
    load(input: WorkbookInput, options?: LoadWorkbookOptions): Promise<Workbook>;
    serialize(workbook: Workbook, options?: SaveWorkbookOptions): Promise<Uint8Array>;
    writeToPath(path: string, workbook: Workbook, options?: SaveWorkbookOptions): Promise<void>;
}
//# sourceMappingURL=XlsxDocument.d.ts.map