type CellPrimitive = string | number | boolean | Date | null;
interface CellAddress {
    row: number;
    col: number;
}
interface CellStyle {
    numFmt?: string;
    fontName?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
}
interface CellFormula {
    expression: string;
    result?: CellPrimitive;
}
interface WorksheetOptions {
    id?: string;
    name: string;
}
interface TableOptions {
    name: string;
    range: string;
    headerRow?: boolean;
    totalsRow?: boolean;
}
/** Optional row placement for `Worksheet.addRow` and `Table.addRow`. */
interface AddRowOptions {
    /**
     * Insert a new row **before** the **row** of this A1 cell (column is ignored). The cell must be a valid
     * single-cell A1 address (e.g. `"A4"` or `"B4"`; both use row 3 as 0-based). Omit to append after the
     * last row that currently has cell data (no shifts; returns the next free row index).
     */
    at?: string;
}
/** Cell values for one table row: left-to-right array, or 0-based column offset within the table → value. */
type TableRowValues = CellPrimitive[] | Record<number, CellPrimitive>;
/** Options for `Worksheet.addTableRow`. */
interface AddTableRowOptions {
    /**
     * Same as `Worksheet.addRow`’s `at`: a single A1 address; the **row** (only) selects where to insert.
     * Must fall within the table’s vertical span or immediately below it.
     */
    at?: string;
    /** Values for the new row within the table’s column span (see `TableRowValues`). */
    values?: TableRowValues;
}
type ChartType = "line" | "pie";
interface ChartSeriesOptions {
    values: string;
    categories?: string;
    name?: string;
}
interface ChartPosition {
    /** Top-left of the chart anchor, e.g. `"E1"`. */
    from: string;
    /** Bottom-right of the chart anchor, e.g. `"M21"`. */
    to: string;
}
interface ChartOptions {
    id?: string;
    type: ChartType;
    title?: string;
    series: ChartSeriesOptions[];
    position?: ChartPosition;
}
interface WorkbookMetadata {
    createdBy?: string;
    modifiedBy?: string;
    createdAt?: Date;
    modifiedAt?: Date;
}
interface LoadWorkbookOptions {
    preserveStyles?: boolean;
}
interface SaveWorkbookOptions {
    includeStyles?: boolean;
}
type WorkbookInput = Uint8Array | string;

declare class Cell {
    private _value;
    private _formula?;
    private _style?;
    private readonly _onChange?;
    constructor(value?: CellPrimitive, onChange?: () => void);
    get value(): CellPrimitive;
    get formula(): CellFormula | undefined;
    get style(): CellStyle | undefined;
    setValue(value: CellPrimitive): this;
    setFormula(formulaExpression: string, result?: CellPrimitive): this;
    setStyle(style: CellStyle): this;
}

declare class Chart {
    private readonly _id;
    private _type;
    private _title?;
    private _series;
    private _from;
    private _to;
    private readonly _onChange?;
    constructor(options: ChartOptions, onChange?: () => void);
    get id(): string;
    get type(): ChartType;
    get title(): string | undefined;
    get series(): ChartSeriesOptions[];
    get position(): ChartPosition;
    setTitle(title: string | undefined): this;
    setSeries(series: ChartSeriesOptions[]): this;
    setPosition(position: ChartPosition): this;
    /**
     * Shifts anchor rows and series range strings when a row is inserted on `worksheetName` before
     * the **row** of A1 `beforeA1` (the column in `beforeA1` is ignored).
     */
    applyRowInsertBefore(beforeA1: string, worksheetName: string): void;
}

declare class Table {
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

interface WorksheetCellEntry {
    row: number;
    col: number;
    cell: Cell;
}
declare class Worksheet {
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

declare class Workbook {
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

declare class XlsxParser {
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

declare class XlsxWriter {
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

declare class XlsxDocument {
    private readonly _parser;
    private readonly _writer;
    constructor(parser?: XlsxParser, writer?: XlsxWriter);
    createWorkbook(): Workbook;
    load(input: WorkbookInput, options?: LoadWorkbookOptions): Promise<Workbook>;
    serialize(workbook: Workbook, options?: SaveWorkbookOptions): Promise<Uint8Array>;
    writeToPath(path: string, workbook: Workbook, options?: SaveWorkbookOptions): Promise<void>;
}

declare class CellRange {
    private readonly _start;
    private readonly _end;
    static fromA1(range: string): CellRange;
    /** One cell, e.g. `B4` (1-based row/column in Excel, stored as 0-based in {@link CellAddress}). */
    static addressFromA1(a1: string): CellAddress;
    static addressToA1(address: CellAddress): string;
    constructor(start: CellAddress, end: CellAddress);
    get start(): CellAddress;
    get end(): CellAddress;
    toA1(): string;
    private static _parseAddress;
    private static _columnToIndex;
    private static _addressToA1;
    private static _indexToColumn;
}

/** Excel .xlsx grid size (OOXML / Excel 2007+). */
declare const EXCEL_MAX_ROW_1BASED = 1048576;
declare const EXCEL_MAX_ROW_0BASED: number;
declare const EXCEL_MAX_COL_1BASED = 16384;
declare const EXCEL_MAX_COL_0BASED: number;

export { Cell, CellRange, Chart, EXCEL_MAX_COL_0BASED, EXCEL_MAX_COL_1BASED, EXCEL_MAX_ROW_0BASED, EXCEL_MAX_ROW_1BASED, Table, Workbook, Worksheet, XlsxDocument, XlsxParser, XlsxWriter };
export type { AddRowOptions, AddTableRowOptions, CellAddress, CellFormula, CellPrimitive, CellStyle, ChartOptions, ChartPosition, ChartSeriesOptions, ChartType, LoadWorkbookOptions, SaveWorkbookOptions, TableOptions, TableRowValues, WorkbookInput, WorkbookMetadata, WorksheetOptions };
