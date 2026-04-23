import { EXCEL_MAX_ROW_0BASED } from "../excelLimits";
import type { ChartOptions, ChartPosition, ChartSeriesOptions, ChartType } from "../types";
import { shiftRefsInStringForRowInsert } from "../refs/shiftRowInsert";

const DEFAULT_POSITION: ChartPosition = {
  from: { row: 0, col: 4 },
  to: { row: 20, col: 12 }
};

export class Chart {
  private readonly _id: string;
  private _type: ChartType;
  private _title?: string;
  private _series: ChartSeriesOptions[];
  private _position: ChartPosition;
  private readonly _onChange?: () => void;

  constructor(options: ChartOptions, onChange?: () => void) {
    this._id = options.id ?? `chart_${Date.now()}_${Math.floor(Math.random() * 1_000_000)}`;
    this._type = options.type;
    this._title = options.title;
    this._series = options.series.map((series) => ({ ...series }));
    this._position = options.position
      ? {
          from: { ...options.position.from },
          to: { ...options.position.to }
        }
      : {
          from: { ...DEFAULT_POSITION.from },
          to: { ...DEFAULT_POSITION.to }
        };
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
      from: { ...this._position.from },
      to: { ...this._position.to }
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
    this._position = {
      from: { ...position.from },
      to: { ...position.to }
    };
    this._onChange?.();
    return this;
  }

  /** Shifts anchor rows and series range strings when a row is inserted on `worksheetName` before `insertBefore`. */
  public applyRowInsertBefore(insertBefore: number, worksheetName: string): void {
    let changed = false;
    const bumpRow = (row: number): number => (row >= insertBefore ? row + 1 : row);
    const nextFromRow = bumpRow(this._position.from.row);
    const nextToRow = bumpRow(this._position.to.row);
    if (nextFromRow > EXCEL_MAX_ROW_0BASED || nextToRow > EXCEL_MAX_ROW_0BASED) {
      throw new Error(
        `Chart anchor row cannot exceed Excel maximum (${EXCEL_MAX_ROW_0BASED} zero-based, ${EXCEL_MAX_ROW_0BASED + 1} rows)`
      );
    }
    if (nextFromRow !== this._position.from.row || nextToRow !== this._position.to.row) {
      this._position = {
        from: { row: nextFromRow, col: this._position.from.col },
        to: { row: nextToRow, col: this._position.to.col }
      };
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
