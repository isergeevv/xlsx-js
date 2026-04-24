import type { AddRowOptions, TableOptions } from "../types";
export declare class Table {
    private _name;
    private _range;
    private _headerRow;
    private _totalsRow;
    constructor(options: TableOptions);
    get name(): string;
    get range(): string;
    get headerRow(): boolean;
    get totalsRow(): boolean;
    rename(nextName: string): this;
    setRange(nextRange: string): this;
    setHeaderRow(enabled: boolean): this;
    setTotalsRow(enabled: boolean): this;
    /**
     * Grows the table range by one row at the bottom, or updates the range as if a sheet row were inserted before `at`
     * (aligned with `Worksheet.addRow` table bookkeeping). Does not move cell contents; use `Worksheet.addRow` when
     * inserting rows in populated sheets.
     */
    addRow(options?: AddRowOptions): this;
    private _assertName;
}
//# sourceMappingURL=Table.d.ts.map