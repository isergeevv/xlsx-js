import { EXCEL_MAX_ROW_0BASED } from "../excelLimits";
import type { AddRowOptions, TableOptions } from "../types";
import { CellRange } from "./CellRange";

export class Table {
  private _name: string;
  private _range: string;
  private _headerRow: boolean;
  private _totalsRow: boolean;

  constructor(options: TableOptions) {
    this._name = options.name;
    this._range = options.range;
    this._headerRow = options.headerRow ?? true;
    this._totalsRow = options.totalsRow ?? false;
  }

  public get name(): string {
    return this._name;
  }

  public get range(): string {
    return this._range;
  }

  public get headerRow(): boolean {
    return this._headerRow;
  }

  public get totalsRow(): boolean {
    return this._totalsRow;
  }

  public rename(nextName: string): this {
    this._assertName(nextName);
    this._name = nextName;
    return this;
  }

  public setRange(nextRange: string): this {
    this._range = nextRange;
    return this;
  }

  public setHeaderRow(enabled: boolean): this {
    this._headerRow = enabled;
    return this;
  }

  public setTotalsRow(enabled: boolean): this {
    this._totalsRow = enabled;
    return this;
  }

  /**
   * Grows the table range by one row at the bottom, or updates the range as if a sheet row were inserted before `at`
   * (aligned with `Worksheet.addRow` table bookkeeping). Does not move cell contents; use `Worksheet.addRow` when
   * inserting rows in populated sheets.
   */
  public addRow(options?: AddRowOptions): this {
    const parsed = CellRange.fromA1(this._range);
    const sr = parsed.start.row;
    const sc = parsed.start.col;
    const er = parsed.end.row;
    const ec = parsed.end.col;

    if (options?.at === undefined) {
      if (er + 1 > EXCEL_MAX_ROW_0BASED) {
        throw new Error(`Table range cannot extend past row ${EXCEL_MAX_ROW_0BASED + 1}`);
      }
      this._range = new CellRange({ row: sr, col: sc }, { row: er + 1, col: ec }).toA1();
      return this;
    }

    const at = CellRange.addressFromA1(options.at).row;
    if (!Number.isInteger(at) || at < 0) {
      throw new Error("Row index must be a non-negative integer");
    }

    if (at < sr) {
      if (er + 1 > EXCEL_MAX_ROW_0BASED) {
        throw new Error(`Table range cannot extend past row ${EXCEL_MAX_ROW_0BASED + 1}`);
      }
      this._range = new CellRange({ row: sr + 1, col: sc }, { row: er + 1, col: ec }).toA1();
    } else if (at <= er + 1) {
      if (er + 1 > EXCEL_MAX_ROW_0BASED) {
        throw new Error(`Table range cannot extend past row ${EXCEL_MAX_ROW_0BASED + 1}`);
      }
      this._range = new CellRange({ row: sr, col: sc }, { row: er + 1, col: ec }).toA1();
    }

    return this;
  }

  private _assertName(name: string): void {
    if (!name.trim()) {
      throw new Error("Table name cannot be empty");
    }
  }
}
