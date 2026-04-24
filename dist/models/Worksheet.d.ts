import { Cell } from "./Cell";
import { Chart } from "./Chart";
import { Table } from "./Table";
import type { AddRowOptions, AddTableRowOptions, CellPrimitive, ChartOptions, TableOptions, WorksheetOptions } from "../types";
export interface WorksheetCellEntry {
    row: number;
    col: number;
    cell: Cell;
}
export declare class Worksheet {
    private readonly _id;
    private _name;
    private readonly _cells;
    private readonly _tables;
    private readonly _charts;
    private _dirty;
    constructor(options: WorksheetOptions);
    get id(): string;
    get name(): string;
    get isDirty(): boolean;
    markClean(): this;
    rename(nextName: string): this;
    getCell(a1: string): Cell;
    setCellValue(a1: string, value: CellPrimitive): this;
    deleteCell(a1: string): boolean;
    /**
     * Adds a logical row: either appends after the last used row (no cell moves), or inserts before the row
     * given by `options.at` (A1; column ignored) and shifts existing cells at that row and below down by one. When
     * inserting, table ranges, chart anchors/series strings, and formula text on this sheet are adjusted for the
     * new row (best-effort A1 / row-range rewriting, not a full formula parse).
     * @returns The 0-based row index of the new empty row.
     */
    addRow(options?: AddRowOptions): number;
    /**
     * Appends a row to the table’s range (or inserts before `at` via `addRow`, which shifts the sheet), then optionally
     * writes `values` across the table’s columns on that new row.
     * @returns 0-based sheet row index of the new table row.
     */
    addTableRow(tableName: string, options?: AddTableRowOptions): number;
    addTable(options: TableOptions): Table;
    getTable(name: string): Table | undefined;
    removeTable(name: string): boolean;
    listTables(): Table[];
    addChart(options: ChartOptions): Chart;
    getChart(id: string): Chart | undefined;
    removeChart(id: string): boolean;
    listCharts(): Chart[];
    listCells(): WorksheetCellEntry[];
    private _getOrCreateAt;
    private _writeTableRowValues;
    private static _key;
    private static _assertAddressInGrid;
}
//# sourceMappingURL=Worksheet.d.ts.map