import { Workbook } from "../models/Workbook";
import { Worksheet } from "../models/Worksheet";
import { writeFile } from "node:fs/promises";
import type { CellPrimitive, CellStyle, SaveWorkbookOptions } from "../types";
import { getWorkbookSnapshot } from "./internal/WorkbookSnapshot";
import { encodeText, writeZip } from "./internal/ZipArchive";
import { xmlEscape } from "./internal/Xml";

export class XlsxWriter {
  public async write(workbook: Workbook, options: SaveWorkbookOptions = {}): Promise<Uint8Array> {
    const worksheets = workbook.listWorksheets();
    const snapshot = getWorkbookSnapshot(workbook);
    const sheetPathByName = snapshot?.sheetPathByName ?? new Map<string, string>();
    const sheetXmlByName = snapshot?.sheetXmlByName ?? new Map<string, string>();
    const baseEntries = new Map<string, Uint8Array>(snapshot?.entries ?? []);

    const resolvedSheetPaths = this._resolveSheetPaths(worksheets, sheetPathByName);
    for (const worksheet of worksheets) {
      const path = resolvedSheetPaths.get(worksheet.name);
      if (!path) {
        continue;
      }
      const baseXml = sheetXmlByName.get(worksheet.name);
      const xml = this._worksheetXml(worksheet, baseXml);
      baseEntries.set(path, encodeText(xml));
    }

    if (!snapshot) {
      baseEntries.set("[Content_Types].xml", encodeText(this._contentTypesXml(worksheets.length, resolvedSheetPaths)));
      baseEntries.set("_rels/.rels", encodeText(this._rootRelsXml()));
      baseEntries.set("xl/workbook.xml", encodeText(this._workbookXml(worksheets)));
      baseEntries.set("xl/_rels/workbook.xml.rels", encodeText(this._workbookRelsXml(worksheets, resolvedSheetPaths)));
    }
    baseEntries.set("xl/xlsxjs.json", encodeText(this._metadataJson(workbook, options)));

    return writeZip([...baseEntries.entries()].map(([name, data]) => ({ name, data })));
  }

  public async writeToPath(path: string, workbook: Workbook, options: SaveWorkbookOptions = {}): Promise<void> {
    const buffer = await this.write(workbook, options);
    await writeFile(path, buffer);
  }

  private _worksheetXml(sheet: Worksheet, baseXml?: string): string {
    const rows = new Map<number, Array<{ col: number; xml: string }>>();
    for (const cellEntry of sheet.listCells()) {
      const rowEntries = rows.get(cellEntry.row) ?? [];
      rowEntries.push({ col: cellEntry.col, xml: this._cellXml(cellEntry.row, cellEntry.col, cellEntry.cell.value, cellEntry.cell.formula) });
      rows.set(cellEntry.row, rowEntries);
    }

    const sortedRows = [...rows.entries()].sort((a, b) => a[0] - b[0]);
    const rowXml = sortedRows
      .map(([row, cells]) => {
        const sortedCells = cells.sort((a, b) => a.col - b.col).map((entry) => entry.xml).join("");
        return `<row r="${row + 1}">${sortedCells}</row>`;
      })
      .join("");

    const sheetDataXml = `<sheetData>${rowXml}</sheetData>`;
    if (baseXml && /<sheetData[\s\S]*<\/sheetData>/.test(baseXml)) {
      return baseXml.replace(/<sheetData[\s\S]*<\/sheetData>/, sheetDataXml);
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  ${sheetDataXml}
</worksheet>`;
  }

  private _cellXml(
    row: number,
    col: number,
    value: CellPrimitive,
    formula: { expression: string; result?: CellPrimitive } | undefined
  ): string {
    const ref = `${this._columnName(col)}${row + 1}`;
    if (formula) {
      const serialized = this._serializePrimitive(formula.result ?? null);
      const typeAttr = serialized.type ? ` t="${serialized.type}"` : "";
      return `<c r="${ref}"${typeAttr}><f>${xmlEscape(formula.expression)}</f><v>${xmlEscape(serialized.value)}</v></c>`;
    }

    const serialized = this._serializePrimitive(value);
    if (serialized.type === "inlineStr") {
      return `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(serialized.value)}</t></is></c>`;
    }
    const typeAttr = serialized.type ? ` t="${serialized.type}"` : "";
    return `<c r="${ref}"${typeAttr}><v>${xmlEscape(serialized.value)}</v></c>`;
  }

  private _serializePrimitive(value: CellPrimitive): { type?: "n" | "b" | "str" | "inlineStr"; value: string } {
    if (value === null) {
      return { type: "str", value: "__xlsxjs:null" };
    }
    if (typeof value === "number") {
      return { type: "n", value: String(value) };
    }
    if (typeof value === "boolean") {
      return { type: "b", value: value ? "1" : "0" };
    }
    if (value instanceof Date) {
      return { type: "str", value: `__xlsxjs:date:${value.toISOString()}` };
    }
    return { type: "inlineStr", value };
  }

  private _metadataJson(workbook: Workbook, options: SaveWorkbookOptions): string {
    const metadata = workbook.metadata;
    return JSON.stringify({
      workbook: {
        createdBy: metadata.createdBy ?? null,
        modifiedBy: metadata.modifiedBy ?? null,
        createdAt: metadata.createdAt?.toISOString() ?? null,
        modifiedAt: metadata.modifiedAt?.toISOString() ?? null
      },
      sheets: workbook.listWorksheets().map((sheet) => ({
        name: sheet.name,
        id: sheet.id,
        tables: sheet.listTables().map((table) => ({
          name: table.name,
          range: table.range,
          headerRow: table.headerRow,
          totalsRow: table.totalsRow
        })),
        styles: options.includeStyles
          ? sheet
              .listCells()
              .filter((entry) => entry.cell.style)
              .map((entry) => ({
                row: entry.row,
                col: entry.col,
                style: this._sanitizeStyle(entry.cell.style)
              }))
          : []
      }))
    });
  }

  private _sanitizeStyle(style: CellStyle | undefined): CellStyle | null {
    if (!style) {
      return null;
    }
    return { ...style };
  }

  private _contentTypesXml(sheetCount: number, sheetPathByName: Map<string, string>): string {
    let overrides = `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`;
    const worksheetPaths = new Set<string>([...sheetPathByName.values()]);
    if (worksheetPaths.size === 0) {
      for (let i = 1; i <= sheetCount; i += 1) {
        worksheetPaths.add(`xl/worksheets/sheet${i}.xml`);
      }
    }
    for (const path of worksheetPaths) {
      overrides += `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${overrides}
</Types>`;
  }

  private _rootRelsXml(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  }

  private _workbookXml(worksheets: Worksheet[]): string {
    const sheetsXml = worksheets
      .map((sheet, index) => `<sheet name="${xmlEscape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
      .join("");

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheetsXml}</sheets>
</workbook>`;
  }

  private _workbookRelsXml(worksheets: Worksheet[], sheetPathByName: Map<string, string>): string {
    let relationships = "";
    for (let i = 0; i < worksheets.length; i += 1) {
      const sheet = worksheets[i];
      const fullPath = sheetPathByName.get(sheet.name) ?? `xl/worksheets/sheet${i + 1}.xml`;
      const target = fullPath.replace(/^xl\//, "");
      relationships += `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${target}"/>`;
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${relationships}
</Relationships>`;
  }

  private _columnName(col: number): string {
    let n = col + 1;
    let out = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  }

  private _resolveSheetPaths(worksheets: Worksheet[], existing: Map<string, string>): Map<string, string> {
    const used = new Set(existing.values());
    const resolved = new Map<string, string>();
    let nextIndex = 1;

    for (const sheet of worksheets) {
      const current = existing.get(sheet.name);
      if (current) {
        resolved.set(sheet.name, current);
        continue;
      }

      while (used.has(`xl/worksheets/sheet${nextIndex}.xml`)) {
        nextIndex += 1;
      }
      const generated = `xl/worksheets/sheet${nextIndex}.xml`;
      used.add(generated);
      resolved.set(sheet.name, generated);
      nextIndex += 1;
    }

    return resolved;
  }
}
