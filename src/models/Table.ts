import type { TableOptions } from "../types";

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

  private _assertName(name: string): void {
    if (!name.trim()) {
      throw new Error("Table name cannot be empty");
    }
  }
}
