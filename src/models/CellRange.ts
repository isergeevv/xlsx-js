import type { CellAddress } from "../types";

export class CellRange {
  private readonly _start: CellAddress;
  private readonly _end: CellAddress;

  public static fromA1(range: string): CellRange {
    const [left, right] = range.split(":");
    if (!left || !right) {
      throw new Error(`Invalid A1 range "${range}"`);
    }
    return new CellRange(CellRange._parseAddress(left), CellRange._parseAddress(right));
  }

  private static _parseAddress(address: string): CellAddress {
    const match = /^([A-Z]+)(\d+)$/i.exec(address.trim());
    if (!match) {
      throw new Error(`Invalid A1 address "${address}"`);
    }

    const [, colText, rowText] = match;
    return {
      row: Number(rowText) - 1,
      col: CellRange._columnToIndex(colText.toUpperCase())
    };
  }

  private static _columnToIndex(column: string): number {
    let value = 0;
    for (let i = 0; i < column.length; i += 1) {
      value = value * 26 + (column.charCodeAt(i) - 64);
    }
    return value - 1;
  }

  private static _addressToA1(address: CellAddress): string {
    return `${CellRange._indexToColumn(address.col)}${address.row + 1}`;
  }

  private static _indexToColumn(colIndex: number): string {
    let n = colIndex + 1;
    let out = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  }

  constructor(start: CellAddress, end: CellAddress) {
    this._start = start;
    this._end = end;
  }

  public get start(): CellAddress {
    return { ...this._start };
  }

  public get end(): CellAddress {
    return { ...this._end };
  }

  public toA1(): string {
    return `${CellRange._addressToA1(this._start)}:${CellRange._addressToA1(this._end)}`;
  }
}
