import type { ChartOptions, ChartPosition, ChartSeriesOptions, ChartType } from "../types";

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
}
