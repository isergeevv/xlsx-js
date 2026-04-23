import { Workbook } from "../models/Workbook";
import { Worksheet } from "../models/Worksheet";
import { readFile } from "node:fs/promises";
import type { CellPrimitive, CellStyle, LoadWorkbookOptions, WorkbookInput } from "../types";
import { decodeText, readZip } from "./internal/ZipArchive";
import { getAttribute, readTagText, xmlUnescape } from "./internal/Xml";

export class XlsxParser {
  public async parse(input: WorkbookInput, options: LoadWorkbookOptions = {}): Promise<Workbook> {
    const bytes = typeof input === "string" ? new Uint8Array(await readFile(input)) : input;
    const zipEntries = await readZip(bytes);
    const workbookXml = this._getTextEntry(zipEntries, "xl/workbook.xml");
    const workbookMeta = this._parseMetadata(zipEntries.get("xl/xlsxjs.json"));
    const workbook = new Workbook(workbookMeta.workbook);

    const sheets = this._parseSheetEntries(workbookXml);
    for (const sheet of sheets) {
      const worksheet = workbook.addWorksheet(sheet.name);
      const xml = this._getTextEntry(zipEntries, sheet.path);
      this._parseWorksheetXml(xml, worksheet);

      const sheetMeta = workbookMeta.sheets.get(sheet.name);
      for (const table of sheetMeta?.tables ?? []) {
        worksheet.addTable(table);
      }
      if (options.preserveStyles) {
        for (const styleEntry of sheetMeta?.styles ?? []) {
          worksheet.getCell(styleEntry.row, styleEntry.col).setStyle(styleEntry.style);
        }
      }
    }

    return workbook;
  }

  private _getTextEntry(entries: Map<string, Uint8Array>, name: string): string {
    const raw = entries.get(name);
    if (!raw) {
      throw new Error(`Missing XLSX entry: ${name}`);
    }
    return decodeText(raw);
  }

  private _parseSheetEntries(workbookXml: string): Array<{ name: string; path: string }> {
    const sheets: Array<{ name: string; path: string }> = [];
    const sheetRegex = /<sheet\b[^>]*>/g;
    let match = sheetRegex.exec(workbookXml);
    let index = 1;
    while (match) {
      const name = getAttribute(match[0], "name");
      if (name) {
        sheets.push({ name, path: `xl/worksheets/sheet${index}.xml` });
        index += 1;
      }
      match = sheetRegex.exec(workbookXml);
    }
    return sheets;
  }

  private _parseWorksheetXml(xml: string, worksheet: Worksheet): void {
    const cellRegex = /<c\b[^>]*>([\s\S]*?)<\/c>/g;
    let match = cellRegex.exec(xml);
    while (match) {
      const wholeTag = match[0];
      const inner = match[1];
      const reference = getAttribute(wholeTag, "r");
      if (!reference) {
        match = cellRegex.exec(xml);
        continue;
      }
      const { row, col } = this._fromA1(reference);
      const type = getAttribute(wholeTag, "t");
      const formula = readTagText(inner, "f");
      const valueText = type === "inlineStr" ? readTagText(inner, "t") : readTagText(inner, "v");
      const primitive = this._parsePrimitive(valueText, type);

      const cell = worksheet.getCell(row, col);
      if (formula) {
        cell.setFormula(formula, primitive);
      } else {
        cell.setValue(primitive);
      }
      match = cellRegex.exec(xml);
    }
  }

  private _parsePrimitive(raw: string | undefined, type: string | undefined): CellPrimitive {
    if (raw === undefined) {
      return null;
    }
    if (raw === "__xlsxjs:null") {
      return null;
    }
    if (raw.startsWith("__xlsxjs:date:")) {
      return new Date(raw.slice("__xlsxjs:date:".length));
    }
    if (type === "b") {
      return raw === "1";
    }
    if (type === "n") {
      return Number(raw);
    }
    return xmlUnescape(raw);
  }

  private _fromA1(reference: string): { row: number; col: number } {
    const match = /^([A-Z]+)(\d+)$/i.exec(reference);
    if (!match) {
      throw new Error(`Invalid cell reference ${reference}`);
    }
    const row = Number(match[2]) - 1;
    let col = 0;
    const colText = match[1].toUpperCase();
    for (let i = 0; i < colText.length; i += 1) {
      col = col * 26 + (colText.charCodeAt(i) - 64);
    }
    return { row, col: col - 1 };
  }

  private _parseMetadata(raw: Uint8Array | undefined): {
    workbook: { createdBy?: string; modifiedBy?: string; createdAt?: Date; modifiedAt?: Date };
    sheets: Map<string, { styles: Array<{ row: number; col: number; style: CellStyle }>; tables: Array<{ name: string; range: string; headerRow?: boolean; totalsRow?: boolean }> }>;
  } {
    if (!raw) {
      return { workbook: {}, sheets: new Map() };
    }

    const parsed = JSON.parse(decodeText(raw)) as {
      workbook?: { createdBy?: string | null; modifiedBy?: string | null; createdAt?: string | null; modifiedAt?: string | null };
      sheets?: Array<{
        name: string;
        styles?: Array<{ row: number; col: number; style: CellStyle }>;
        tables?: Array<{ name: string; range: string; headerRow?: boolean; totalsRow?: boolean }>;
      }>;
    };

    const sheets = new Map<string, { styles: Array<{ row: number; col: number; style: CellStyle }>; tables: Array<{ name: string; range: string; headerRow?: boolean; totalsRow?: boolean }> }>();
    for (const sheet of parsed.sheets ?? []) {
      sheets.set(sheet.name, {
        styles: sheet.styles ?? [],
        tables: sheet.tables ?? []
      });
    }

    return {
      workbook: {
        createdBy: parsed.workbook?.createdBy ?? undefined,
        modifiedBy: parsed.workbook?.modifiedBy ?? undefined,
        createdAt: parsed.workbook?.createdAt ? new Date(parsed.workbook.createdAt) : undefined,
        modifiedAt: parsed.workbook?.modifiedAt ? new Date(parsed.workbook.modifiedAt) : undefined
      },
      sheets
    };
  }
}
