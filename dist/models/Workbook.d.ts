import { Worksheet } from "./Worksheet";
import type { WorkbookMetadata } from "../types";
export declare class Workbook {
    private readonly _metadata;
    private readonly _sheets;
    constructor(metadata?: WorkbookMetadata);
    get metadata(): WorkbookMetadata;
    addWorksheet(name: string): Worksheet;
    getWorksheet(name: string): Worksheet | undefined;
    removeWorksheet(name: string): boolean;
    listWorksheets(): Worksheet[];
    renameWorksheet(from: string, to: string): this;
}
//# sourceMappingURL=Workbook.d.ts.map