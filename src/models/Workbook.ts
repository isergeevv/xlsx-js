import { Worksheet } from "./Worksheet";
import type { WorkbookMetadata } from "../types";

export class Workbook {
  private readonly _metadata: WorkbookMetadata;
  private readonly _sheets = new Map<string, Worksheet>();

  constructor(metadata: WorkbookMetadata = {}) {
    this._metadata = { ...metadata };
  }

  public get metadata(): WorkbookMetadata {
    return { ...this._metadata };
  }

  public addWorksheet(name: string): Worksheet {
    if (this._sheets.has(name)) {
      throw new Error(`Worksheet "${name}" already exists`);
    }
    const sheet = new Worksheet({ name });
    this._sheets.set(name, sheet);
    return sheet;
  }

  public getWorksheet(name: string): Worksheet | undefined {
    return this._sheets.get(name);
  }

  public removeWorksheet(name: string): boolean {
    return this._sheets.delete(name);
  }

  public listWorksheets(): Worksheet[] {
    return [...this._sheets.values()];
  }

  public renameWorksheet(from: string, to: string): this {
    const sheet = this._sheets.get(from);
    if (!sheet) {
      throw new Error(`Worksheet "${from}" does not exist`);
    }
    if (this._sheets.has(to)) {
      throw new Error(`Worksheet "${to}" already exists`);
    }

    this._sheets.delete(from);
    sheet.rename(to);
    this._sheets.set(to, sheet);
    return this;
  }
}
