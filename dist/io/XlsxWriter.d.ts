import { Workbook } from "../models/Workbook";
import type { SaveWorkbookOptions } from "../types";
export declare class XlsxWriter {
    write(workbook: Workbook, options?: SaveWorkbookOptions): Promise<Uint8Array>;
    writeToPath(path: string, workbook: Workbook, options?: SaveWorkbookOptions): Promise<void>;
    private _worksheetXml;
    private _cellXml;
    private _serializePrimitive;
    private _metadataJson;
    private _sanitizeStyle;
    private _contentTypesXml;
    private _rootRelsXml;
    private _workbookXml;
    private _workbookRelsXml;
    private _columnName;
    private _resolveSheetPaths;
    private _upsertDrawingTag;
    private _buildChartAssets;
    private _sheetRelsPath;
    private _sheetRelsXml;
    private _drawingXml;
    private _drawingRelsXml;
    private _chartAnchorXml;
    private _chartXml;
}
//# sourceMappingURL=XlsxWriter.d.ts.map