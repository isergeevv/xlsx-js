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
