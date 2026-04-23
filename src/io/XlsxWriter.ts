import { Workbook } from "../models/Workbook";
import { Chart } from "../models/Chart";
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
    const chartAssets = this._buildChartAssets(worksheets, resolvedSheetPaths);
    for (const worksheet of worksheets) {
      const path = resolvedSheetPaths.get(worksheet.name);
      if (!path) {
        continue;
      }
      const sheetChartAssets = chartAssets.bySheetPath.get(path);
      if (snapshot && !worksheet.isDirty && !sheetChartAssets) {
        continue;
      }
      const baseXml = sheetXmlByName.get(worksheet.name);
      const xml = this._worksheetXml(worksheet, baseXml, sheetChartAssets?.drawingRelId);
      baseEntries.set(path, encodeText(xml));
      if (sheetChartAssets) {
        baseEntries.set(sheetChartAssets.sheetRelsPath, encodeText(sheetChartAssets.sheetRelsXml));
      }
    }

    for (const file of chartAssets.files) {
      baseEntries.set(file.path, encodeText(file.xml));
    }

    if (!snapshot) {
      baseEntries.set(
        "[Content_Types].xml",
        encodeText(this._contentTypesXml(worksheets.length, resolvedSheetPaths, chartAssets.contentTypeOverrides))
      );
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

  private _worksheetXml(sheet: Worksheet, baseXml?: string, drawingRelId?: string): string {
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
    if (baseXml) {
      let nextXml = baseXml;
      if (/<sheetData[\s\S]*<\/sheetData>/.test(baseXml)) {
        nextXml = baseXml.replace(/<sheetData[\s\S]*<\/sheetData>/, sheetDataXml);
      } else if (/<sheetData\s*\/>/.test(baseXml)) {
        nextXml = baseXml.replace(/<sheetData\s*\/>/, sheetDataXml);
      }

      if (drawingRelId) {
        nextXml = this._upsertDrawingTag(nextXml, drawingRelId);
      }
      return nextXml;
    }

    const drawingXml = drawingRelId ? `<drawing r:id="${drawingRelId}"/>` : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${sheetDataXml}
  ${drawingXml}
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
        charts: sheet.listCharts().map((chart) => ({
          id: chart.id,
          type: chart.type,
          title: chart.title,
          series: chart.series,
          position: chart.position
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

  private _contentTypesXml(
    sheetCount: number,
    sheetPathByName: Map<string, string>,
    extraOverrides: Array<{ path: string; contentType: string }>
  ): string {
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
    for (const override of extraOverrides) {
      overrides += `<Override PartName="/${override.path}" ContentType="${override.contentType}"/>`;
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

  private _upsertDrawingTag(xml: string, drawingRelId: string): string {
    const drawingTag = `<drawing r:id="${drawingRelId}"/>`;
    if (/<drawing\b[^>]*\/>/.test(xml)) {
      return xml.replace(/<drawing\b[^>]*\/>/, drawingTag);
    }
    return xml.replace(/<\/worksheet>/, `  ${drawingTag}\n</worksheet>`);
  }

  private _buildChartAssets(
    worksheets: Worksheet[],
    sheetPathByName: Map<string, string>
  ): {
    bySheetPath: Map<
      string,
      {
        drawingRelId: string;
        sheetRelsPath: string;
        sheetRelsXml: string;
      }
    >;
    files: Array<{ path: string; xml: string }>;
    contentTypeOverrides: Array<{ path: string; contentType: string }>;
  } {
    const bySheetPath = new Map<
      string,
      {
        drawingRelId: string;
        sheetRelsPath: string;
        sheetRelsXml: string;
      }
    >();
    const files: Array<{ path: string; xml: string }> = [];
    const overrides: Array<{ path: string; contentType: string }> = [];
    let drawingIndex = 1;
    let chartIndex = 1;

    for (const worksheet of worksheets) {
      if (!worksheet.isDirty) {
        continue;
      }
      const charts = worksheet.listCharts();
      if (charts.length === 0) {
        continue;
      }
      const sheetPath = sheetPathByName.get(worksheet.name);
      if (!sheetPath) {
        continue;
      }

      const drawingPath = `xl/drawings/drawing${drawingIndex}.xml`;
      const drawingRelsPath = `xl/drawings/_rels/drawing${drawingIndex}.xml.rels`;
      const sheetRelsPath = this._sheetRelsPath(sheetPath);
      const drawingRelId = "rIdXlsxJsDrawing1";

      const chartEntries: Array<{
        chartRelId: string;
        chartPath: string;
        chartXml: string;
        anchorXml: string;
      }> = [];
      for (let i = 0; i < charts.length; i += 1) {
        const chart = charts[i];
        const chartPath = `xl/charts/chart${chartIndex}.xml`;
        const chartRelId = `rIdChart${i + 1}`;
        chartEntries.push({
          chartRelId,
          chartPath,
          chartXml: this._chartXml(chart, i),
          anchorXml: this._chartAnchorXml(chart.position, chartRelId, i + 1)
        });
        chartIndex += 1;
      }

      const drawingXml = this._drawingXml(chartEntries.map((entry) => entry.anchorXml));
      const drawingRelsXml = this._drawingRelsXml(
        chartEntries.map((entry) => ({ id: entry.chartRelId, targetPath: entry.chartPath }))
      );

      files.push({ path: drawingPath, xml: drawingXml });
      files.push({ path: drawingRelsPath, xml: drawingRelsXml });
      for (const entry of chartEntries) {
        files.push({ path: entry.chartPath, xml: entry.chartXml });
      }

      overrides.push({
        path: drawingPath,
        contentType: "application/vnd.openxmlformats-officedocument.drawing+xml"
      });
      for (const entry of chartEntries) {
        overrides.push({
          path: entry.chartPath,
          contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
        });
      }

      bySheetPath.set(sheetPath, {
        drawingRelId,
        sheetRelsPath,
        sheetRelsXml: this._sheetRelsXml(drawingRelId, drawingPath)
      });

      drawingIndex += 1;
    }

    return { bySheetPath, files, contentTypeOverrides: overrides };
  }

  private _sheetRelsPath(sheetPath: string): string {
    const segments = sheetPath.split("/");
    const fileName = segments.length > 0 ? segments[segments.length - 1] : "sheet1.xml";
    return `xl/worksheets/_rels/${fileName}.rels`;
  }

  private _sheetRelsXml(drawingRelId: string, drawingPath: string): string {
    const target = drawingPath.replace(/^xl\/drawings\//, "../drawings/");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${drawingRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="${target}"/>
</Relationships>`;
  }

  private _drawingXml(anchors: string[]): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${anchors.join("")}
</xdr:wsDr>`;
  }

  private _drawingRelsXml(relations: Array<{ id: string; targetPath: string }>): string {
    const xml = relations
      .map((relation) => {
        const segments = relation.targetPath.split("/");
        const fileName = segments.length > 0 ? segments[segments.length - 1] : relation.targetPath;
        return `<Relationship Id="${relation.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/${fileName}"/>`;
      })
      .join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${xml}
</Relationships>`;
  }

  private _chartAnchorXml(
    position: { from: { row: number; col: number }; to: { row: number; col: number } },
    chartRelId: string,
    chartIndex: number
  ): string {
    return `<xdr:twoCellAnchor>
    <xdr:from><xdr:col>${position.from.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${position.from.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>${position.to.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${position.to.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="${chartIndex}" name="Chart ${chartIndex}"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm/>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="${chartRelId}"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`;
  }

  private _chartXml(chart: Chart, chartIndex: number): string {
    const seriesXml = chart.series
      .map((series, index) => {
        const nameXml = series.name ? `<c:tx><c:v>${xmlEscape(series.name)}</c:v></c:tx>` : "";
        const catXml = series.categories
          ? `<c:cat><c:strRef><c:f>${xmlEscape(series.categories)}</c:f></c:strRef></c:cat>`
          : "";
        return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${nameXml}${catXml}<c:val><c:numRef><c:f>${xmlEscape(
          series.values
        )}</c:f></c:numRef></c:val></c:ser>`;
      })
      .join("");

    const titleXml = chart.title
      ? `<c:title><c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>${xmlEscape(
          chart.title
        )}</a:t></a:r></a:p></c:rich></c:tx></c:title>`
      : "";

    const chartTypeXml =
      chart.type === "pie"
        ? `<c:pieChart><c:varyColors val="1"/>${seriesXml}</c:pieChart>`
        : `<c:lineChart><c:grouping val="standard"/>${seriesXml}<c:axId val="${500000 + chartIndex}"/><c:axId val="${
            600000 + chartIndex
          }"/></c:lineChart>
<c:catAx><c:axId val="${500000 + chartIndex}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="${
            600000 + chartIndex
          }"/><c:tickLblPos val="nextTo"/></c:catAx>
<c:valAx><c:axId val="${600000 + chartIndex}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="${
            500000 + chartIndex
          }"/><c:tickLblPos val="nextTo"/></c:valAx>`;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    ${titleXml}
    <c:plotArea>${chartTypeXml}</c:plotArea>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
  }
}
