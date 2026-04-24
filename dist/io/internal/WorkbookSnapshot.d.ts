import { Workbook } from "../../models/Workbook";
export interface WorkbookSnapshot {
    entries: Map<string, Uint8Array>;
    sheetPathByName: Map<string, string>;
    sheetXmlByName: Map<string, string>;
}
export declare function setWorkbookSnapshot(workbook: Workbook, snapshot: WorkbookSnapshot): void;
export declare function getWorkbookSnapshot(workbook: Workbook): WorkbookSnapshot | undefined;
//# sourceMappingURL=WorkbookSnapshot.d.ts.map