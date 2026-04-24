import { Workbook } from "../models/Workbook";
import type { LoadWorkbookOptions, WorkbookInput } from "../types";
export declare class XlsxParser {
    parse(input: WorkbookInput, options?: LoadWorkbookOptions): Promise<Workbook>;
    private _chartOptionsFromSnapshot;
    private _chartPositionFromSnapshot;
    private _getTextEntry;
    private _parseSheetEntries;
    private _parseWorkbookRelationshipTargets;
    private _parseWorksheetXml;
    private _parsePrimitive;
    private _fromA1;
    private _parseMetadata;
}
//# sourceMappingURL=XlsxParser.d.ts.map