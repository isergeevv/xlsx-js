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
   * 0-based sheet row index: insert a new row before this index (existing cells at this row and below shift down).
   * Omit to append after the last row that currently has cell data (no shifts; returns the next free row index).
   */
  at?: number;
}

/** Cell values for one table row: left-to-right array, or 0-based column offset within the table → value. */
export type TableRowValues = CellPrimitive[] | Record<number, CellPrimitive>;

/** Options for `Worksheet.addTableRow`. */
export interface AddTableRowOptions {
  /**
   * 0-based sheet row: insert before this row (same as `Worksheet.addRow`), then write `values` on that row.
   * Must fall within the table’s vertical span or immediately below it (`start.row <= at <= end.row + 1`).
   */
  at?: number;
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
  from: CellAddress;
  to: CellAddress;
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
