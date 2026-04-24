export type CellPrimitive = string | number | boolean | Date | null;

export interface CellAddress {
  row: number;
  col: number;
}

export interface CellStyle {
  numFmt?: string;
  fontName?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

export interface CellFormula {
  expression: string;
  result?: CellPrimitive;
}

export interface WorksheetOptions {
  id?: string;
  name: string;
}

export interface TableOptions {
  name: string;
  range: string;
  headerRow?: boolean;
  totalsRow?: boolean;
}

/** Optional row placement for `Worksheet.addRow` and `Table.addRow`. */
export interface AddRowOptions {
  /**
   * Insert a new row **before** the **row** of this A1 cell (column is ignored). The cell must be a valid
   * single-cell A1 address (e.g. `"A4"` or `"B4"`; both use row 3 as 0-based). Omit to append after the
   * last row that currently has cell data (no shifts; returns the next free row index).
   */
  at?: string;
}

/** Cell values for one table row: left-to-right array, or 0-based column offset within the table → value. */
export type TableRowValues = CellPrimitive[] | Record<number, CellPrimitive>;

/** Options for `Worksheet.addTableRow`. */
export interface AddTableRowOptions {
  /**
   * Same as `Worksheet.addRow`’s `at`: a single A1 address; the **row** (only) selects where to insert.
   * Must fall within the table’s vertical span or immediately below it.
   */
  at?: string;
  /** Values for the new row within the table’s column span (see `TableRowValues`). */
  values?: TableRowValues;
}

export type ChartType = "line" | "pie";

export interface ChartSeriesOptions {
  values: string;
  categories?: string;
  name?: string;
}

export interface ChartPosition {
  /** Top-left of the chart anchor, e.g. `"E1"`. */
  from: string;
  /** Bottom-right of the chart anchor, e.g. `"M21"`. */
  to: string;
}

export interface ChartOptions {
  id?: string;
  type: ChartType;
  title?: string;
  series: ChartSeriesOptions[];
  position?: ChartPosition;
}

export interface WorkbookMetadata {
  createdBy?: string;
  modifiedBy?: string;
  createdAt?: Date;
  modifiedAt?: Date;
}

export interface LoadWorkbookOptions {
  preserveStyles?: boolean;
}

export interface SaveWorkbookOptions {
  includeStyles?: boolean;
}

export type WorkbookInput = Uint8Array | string;
