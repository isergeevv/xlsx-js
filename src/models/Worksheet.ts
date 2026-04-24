import { Cell } from "./Cell";
import { Chart } from "./Chart";
import { Table } from "./Table";
import { CellRange } from "./CellRange";
import { EXCEL_MAX_COL_0BASED, EXCEL_MAX_ROW_0BASED } from "../excelLimits";
import { shiftRefsInStringForRowInsert } from "../refs/shiftRowInsert";
import type {
  AddRowOptions,
  AddTableRowOptions,
  CellAddress,
  CellPrimitive,
  ChartOptions,
  TableOptions,
  TableRowValues,
  WorksheetOptions
} from "../types";

export interface WorksheetCellEntry {
  row: number;
  col: number;
  cell: Cell;
}

export class Worksheet {
  private readonly _id: string;
  private _name: string;
  private readonly _cells = new Map<string, Cell>();
  private readonly _tables = new Map<string, Table>();
  private readonly _charts = new Map<string, Chart>();
  private _dirty: boolean;

  constructor(options: WorksheetOptions) {
    this._id = options.id ?? `ws_${Date.now()}_${Math.floor(Math.random() * 1_000_000)}`;
    this._name = options.name;
    this._dirty = false;
  }

  public get id(): string {
    return this._id;
  }

  public get name(): string {
    return this._name;
  }

  public get isDirty(): boolean {
    return this._dirty;
  }

  public markClean(): this {
    this._dirty = false;
    return this;
  }

  public rename(nextName: string): this {
    this._name = nextName;
    this._dirty = true;
    return this;
  }

  public getCell(a1: string): Cell {
    const { row, col } = CellRange.addressFromA1(a1);
    return this._getOrCreateAt(row, col);
  }

  public setCellValue(a1: string, value: CellPrimitive): this {
    const { row, col } = CellRange.addressFromA1(a1);
    this._getOrCreateAt(row, col).setValue(value);
    return this;
  }

  public deleteCell(a1: string): boolean {
    const { row, col } = CellRange.addressFromA1(a1);
    Worksheet._assertAddressInGrid(row, col);
    const deleted = this._cells.delete(Worksheet._key({ row, col }));
    if (deleted) {
      this._dirty = true;
    }
    return deleted;
  }

  /**
   * Adds a logical row: either appends after the last used row (no cell moves), or inserts before the row
   * given by `options.at` (A1; column ignored) and shifts existing cells at that row and below down by one. When
   * inserting, table ranges, chart anchors/series strings, and formula text on this sheet are adjusted for the
   * new row (best-effort A1 / row-range rewriting, not a full formula parse).
   * @returns The 0-based row index of the new empty row.
   */
  public addRow(options?: AddRowOptions): number {
    if (options?.at !== undefined) {
      const at = CellRange.addressFromA1(options.at).row;
      if (!Number.isInteger(at) || at < 0) {
        throw new Error("Row index must be a non-negative integer");
      }
      if (at > EXCEL_MAX_ROW_0BASED) {
        throw new Error(`Row index cannot exceed ${EXCEL_MAX_ROW_0BASED} (Excel max row, 0-based)`);
      }

      for (const { row } of this.listCells()) {
        if (row >= at && row >= EXCEL_MAX_ROW_0BASED) {
          throw new Error(
            `Cannot insert row at ${at}: cell at row ${row} would move past Excel limit (${EXCEL_MAX_ROW_0BASED + 1} rows)`
          );
        }
      }

      const entries = this.listCells().filter((e) => e.row >= at);
      entries.sort((a, b) => (b.row - a.row) || (b.col - a.col));
      for (const { row, col, cell } of entries) {
        this._cells.delete(Worksheet._key({ row, col }));
        this._cells.set(Worksheet._key({ row: row + 1, col }), cell);
      }

      for (const table of this._tables.values()) {
        table.addRow({ at: options.at });
      }

      for (const chart of this._charts.values()) {
        chart.applyRowInsertBefore(options.at, this._name);
      }

      for (const { cell } of this.listCells()) {
        const f = cell.formula;
        if (f?.expression) {
          const nextExpr = shiftRefsInStringForRowInsert(f.expression, this._name, at);
          if (nextExpr !== f.expression) {
            cell.setFormula(nextExpr, f.result);
          }
        }
      }

      this._dirty = true;
      return at;
    }

    let maxRow = -1;
    for (const key of this._cells.keys()) {
      const row = Number(key.split(":")[0]);
      if (row > maxRow) {
        maxRow = row;
      }
    }
    const next = maxRow + 1;
    if (next > EXCEL_MAX_ROW_0BASED) {
      throw new Error(`Next row index ${next} exceeds Excel maximum (${EXCEL_MAX_ROW_0BASED + 1} rows)`);
    }
    return next;
  }

  /**
   * Appends a row to the table’s range (or inserts before `at` via `addRow`, which shifts the sheet), then optionally
   * writes `values` across the table’s columns on that new row.
   * @returns 0-based sheet row index of the new table row.
   */
  public addTableRow(tableName: string, options?: AddTableRowOptions): number {
    const table = this._tables.get(tableName);
    if (!table) {
      throw new Error(`Table "${tableName}" does not exist in worksheet "${this.name}"`);
    }

    const parsed = CellRange.fromA1(table.range);
    const sr = parsed.start.row;
    const sc = parsed.start.col;
    const er = parsed.end.row;
    const ec = parsed.end.col;
    const colCount = ec - sc + 1;

    if (options?.at !== undefined) {
      const at = CellRange.addressFromA1(options.at).row;
      if (at < sr || at > er + 1) {
        throw new Error(`Table row insert "at" (${at}) must satisfy ${sr} <= at <= ${er + 1} for this table`);
      }
      this.addRow({ at: options.at });
      this._writeTableRowValues(at, sc, colCount, options.values);
      return at;
    }

    const newRow = er + 1;
    if (newRow > EXCEL_MAX_ROW_0BASED) {
      throw new Error(`Cannot append table row past Excel maximum (${EXCEL_MAX_ROW_0BASED + 1} rows)`);
    }
    table.addRow();
    this._writeTableRowValues(newRow, sc, colCount, options?.values);
    return newRow;
  }

  public addTable(options: TableOptions): Table {
    if (this._tables.has(options.name)) {
      throw new Error(`Table "${options.name}" already exists in worksheet "${this.name}"`);
    }
    const t = new Table(options);
    this._tables.set(t.name, t);
    this._dirty = true;
    return t;
  }

  public getTable(name: string): Table | undefined {
    return this._tables.get(name);
  }

  public removeTable(name: string): boolean {
    const removed = this._tables.delete(name);
    if (removed) {
      this._dirty = true;
    }
    return removed;
  }

  public listTables(): Table[] {
    return [...this._tables.values()];
  }

  public addChart(options: ChartOptions): Chart {
    const chart = new Chart(options, () => {
      this._dirty = true;
    });
    if (this._charts.has(chart.id)) {
      throw new Error(`Chart "${chart.id}" already exists in worksheet "${this.name}"`);
    }
    this._charts.set(chart.id, chart);
    this._dirty = true;
    return chart;
  }

  public getChart(id: string): Chart | undefined {
    return this._charts.get(id);
  }

  public removeChart(id: string): boolean {
    const removed = this._charts.delete(id);
    if (removed) {
      this._dirty = true;
    }
    return removed;
  }

  public listCharts(): Chart[] {
    return [...this._charts.values()];
  }

  public listCells(): WorksheetCellEntry[] {
    return [...this._cells.entries()].map(([key, cell]) => {
      const [rowText, colText] = key.split(":");
      return {
        row: Number(rowText),
        col: Number(colText),
        cell
      };
    });
  }

  private _getOrCreateAt(row: number, col: number): Cell {
    Worksheet._assertAddressInGrid(row, col);
    const key = Worksheet._key({ row, col });
    const existing = this._cells.get(key);
    if (existing) {
      return existing;
    }

    const created = new Cell(null, () => {
      this._dirty = true;
    });
    this._cells.set(key, created);
    return created;
  }

  private _writeTableRowValues(
    row: number,
    startCol: number,
    colCount: number,
    values: TableRowValues | undefined
  ): void {
    if (values === undefined) {
      return;
    }
    if (Array.isArray(values)) {
      const n = Math.min(colCount, values.length);
      for (let i = 0; i < n; i += 1) {
        this.setCellValue(CellRange.addressToA1({ row, col: startCol + i }), values[i]);
      }
      return;
    }
    for (const [key, v] of Object.entries(values)) {
      const offset = Number(key);
      if (!Number.isInteger(offset) || offset < 0 || offset >= colCount) {
        throw new Error(`Table row value key "${key}" must be an integer column offset in [0, ${colCount - 1}]`);
      }
      this.setCellValue(CellRange.addressToA1({ row, col: startCol + offset }), v);
    }
  }

  private static _key(address: CellAddress): string {
    return `${address.row}:${address.col}`;
  }

  private static _assertAddressInGrid(row: number, col: number): void {
    if (!Number.isInteger(row) || row < 0 || row > EXCEL_MAX_ROW_0BASED) {
      throw new Error(
        `Row index must be an integer in [0, ${EXCEL_MAX_ROW_0BASED}] (${EXCEL_MAX_ROW_0BASED + 1} rows max)`
      );
    }
    if (!Number.isInteger(col) || col < 0 || col > EXCEL_MAX_COL_0BASED) {
      throw new Error(
        `Column index must be an integer in [0, ${EXCEL_MAX_COL_0BASED}] (${EXCEL_MAX_COL_0BASED + 1} columns max)`
      );
    }
  }
}
