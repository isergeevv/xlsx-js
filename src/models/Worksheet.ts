import { Cell } from "./Cell";
import { Table } from "./Table";
import type { CellAddress, CellPrimitive, TableOptions, WorksheetOptions } from "../types";

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

  public getCell(row: number, col: number): Cell {
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

  public setCellValue(row: number, col: number, value: CellPrimitive): this {
    this.getCell(row, col).setValue(value);
    return this;
  }

  public deleteCell(row: number, col: number): boolean {
    const deleted = this._cells.delete(Worksheet._key({ row, col }));
    if (deleted) {
      this._dirty = true;
    }
    return deleted;
  }

  public addTable(options: TableOptions): Table {
    if (this._tables.has(options.name)) {
      throw new Error(`Table "${options.name}" already exists in worksheet "${this.name}"`);
    }
    const table = new Table(options);
    this._tables.set(table.name, table);
    this._dirty = true;
    return table;
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

  private static _key(address: CellAddress): string {
    return `${address.row}:${address.col}`;
  }
}
