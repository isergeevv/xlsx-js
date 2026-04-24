import { CellRange } from "./CellRange";
import { EXCEL_MAX_ROW_0BASED } from "../excelLimits";
import { shiftRefsInStringForRowInsert } from "../refs/shiftRowInsert";
import type { CellAddress, ChartOptions, ChartPosition, ChartSeriesOptions, ChartType } from "../types";

const DEFAULT_POSITION_ANCHOR: { from: CellAddress; to: CellAddress } = (() => {
  const from = CellRange.addressFromA1("E1");
  const to = CellRange.addressFromA1("M21");
  return { from, to };
})();

export class Chart {
  private readonly _id: string;
  private _type: ChartType;
  private _title?: string;
  private _series: ChartSeriesOptions[];
  private _from: CellAddress;
  private _to: CellAddress;
  private readonly _onChange?: () => void;

  constructor(options: ChartOptions, onChange?: () => void) {
    this._id = options.id ?? `chart_${Date.now()}_${Math.floor(Math.random() * 1_000_000)}`;
    this._type = options.type;
    this._title = options.title;
    this._series = options.series.map((series) => ({ ...series }));
    if (options.position) {
      this._from = { ...CellRange.addressFromA1(options.position.from) };
      this._to = { ...CellRange.addressFromA1(options.position.to) };
    } else {
      this._from = { ...DEFAULT_POSITION_ANCHOR.from };
      this._to = { ...DEFAULT_POSITION_ANCHOR.to };
    }
    this._onChange = onChange;
  }

  public get id(): string {
    return this._id;
  }

  public get type(): ChartType {
    return this._type;
  }

  public get title(): string | undefined {
    return this._title;
  }

  public get series(): ChartSeriesOptions[] {
    return this._series.map((entry) => ({ ...entry }));
  }

  public get position(): ChartPosition {
    return {
      from: CellRange.addressToA1(this._from),
      to: CellRange.addressToA1(this._to)
    };
  }

  public setTitle(title: string | undefined): this {
    this._title = title;
    this._onChange?.();
    return this;
  }

  public setSeries(series: ChartSeriesOptions[]): this {
    this._series = series.map((entry) => ({ ...entry }));
    this._onChange?.();
    return this;
  }

  public setPosition(position: ChartPosition): this {
    this._from = { ...CellRange.addressFromA1(position.from) };
    this._to = { ...CellRange.addressFromA1(position.to) };
    this._onChange?.();
    return this;
  }

  /**
   * Shifts anchor rows and series range strings when a row is inserted on `worksheetName` before
   * the **row** of A1 `beforeA1` (the column in `beforeA1` is ignored).
   */
  public applyRowInsertBefore(beforeA1: string, worksheetName: string): void {
    const insertBefore = CellRange.addressFromA1(beforeA1).row;
    let changed = false;
    const bumpRow = (row: number): number => (row >= insertBefore ? row + 1 : row);
    const nextFromRow = bumpRow(this._from.row);
    const nextToRow = bumpRow(this._to.row);
    if (nextFromRow > EXCEL_MAX_ROW_0BASED || nextToRow > EXCEL_MAX_ROW_0BASED) {
      throw new Error(
        `Chart anchor row cannot exceed Excel maximum (${EXCEL_MAX_ROW_0BASED} zero-based, ${EXCEL_MAX_ROW_0BASED + 1} rows)`
      );
    }
    if (nextFromRow !== this._from.row || nextToRow !== this._to.row) {
      this._from = { row: nextFromRow, col: this._from.col };
      this._to = { row: nextToRow, col: this._to.col };
      changed = true;
    }

    const nextSeries = this._series.map((entry) => {
      const values = shiftRefsInStringForRowInsert(entry.values, worksheetName, insertBefore);
      const categories = entry.categories
        ? shiftRefsInStringForRowInsert(entry.categories, worksheetName, insertBefore)
        : undefined;
      const name = entry.name ? shiftRefsInStringForRowInsert(entry.name, worksheetName, insertBefore) : undefined;
      if (values !== entry.values || categories !== entry.categories || name !== entry.name) {
        changed = true;
      }
      return { ...entry, values, categories, name };
    });
    this._series = nextSeries;

    if (changed) {
      this._onChange?.();
    }
  }
}
