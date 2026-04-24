import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, readFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import {
  Cell,
  CellRange,
  Table,
  Workbook,
  Worksheet,
  XlsxDocument
} from "../dist/index.js";

test("Cell supports value, formula, and style updates", () => {
  const cell = new Cell("initial");
  assert.equal(cell.value, "initial");
  assert.equal(cell.formula, undefined);

  cell.setFormula("A1+B1", 42);
  assert.deepEqual(cell.formula, { expression: "A1+B1", result: 42 });

  cell.setStyle({ bold: true, fontName: "Arial" });
  assert.deepEqual(cell.style, { bold: true, fontName: "Arial" });

  cell.setValue(5);
  assert.equal(cell.value, 5);
  assert.equal(cell.formula, undefined);
});

test("Table manages name and display options", () => {
  const table = new Table({ name: "Sales", range: "A1:C10" });
  assert.equal(table.name, "Sales");
  assert.equal(table.range, "A1:C10");
  assert.equal(table.headerRow, true);
  assert.equal(table.totalsRow, false);

  table.rename("Revenue").setRange("A1:D12").setTotalsRow(true);
  assert.equal(table.name, "Revenue");
  assert.equal(table.range, "A1:D12");
  assert.equal(table.totalsRow, true);
});

test("Worksheet manages cells and tables", () => {
  const worksheet = new Worksheet({ name: "Data" });
  assert.equal(worksheet.name, "Data");

  worksheet.setCellValue(0, 0, "hello");
  assert.equal(worksheet.getCell("A1").value, "hello");
  assert.equal(worksheet.deleteCell(0, 0), true);
  assert.equal(worksheet.deleteCell(0, 0), false);

  const table = worksheet.addTable({ name: "T1", range: "A1:B2" });
  assert.equal(table.name, "T1");
  assert.equal(worksheet.getTable("T1"), table);
  assert.equal(worksheet.listTables().length, 1);
  assert.equal(worksheet.removeTable("T1"), true);
  assert.equal(worksheet.removeTable("T1"), false);
});

test("Workbook manages worksheets", () => {
  const workbook = new Workbook({ createdBy: "test-user" });
  const sheet = workbook.addWorksheet("Sheet1");
  assert.equal(sheet.name, "Sheet1");
  assert.equal(workbook.getWorksheet("Sheet1"), sheet);
  assert.equal(workbook.listWorksheets().length, 1);

  workbook.renameWorksheet("Sheet1", "Renamed");
  assert.equal(workbook.getWorksheet("Sheet1"), undefined);
  assert.equal(workbook.getWorksheet("Renamed")?.name, "Renamed");
  assert.equal(workbook.removeWorksheet("Renamed"), true);
});

test("CellRange converts between A1 strings and addresses", () => {
  const range = CellRange.fromA1("A1:C3");
  assert.deepEqual(range.start, { row: 0, col: 0 });
  assert.deepEqual(range.end, { row: 2, col: 2 });
  assert.equal(range.toA1(), "A1:C3");
});

test("XlsxDocument delegates load/serialize operations", async () => {
  const workbook = new Workbook();
  workbook.addWorksheet("Sheet1");

  const parser = {
    async parse(buffer, options) {
      assert.equal(buffer.length, 2);
      assert.deepEqual(options, { preserveStyles: true });
      return workbook;
    }
  };

  const writer = {
    async write(nextWorkbook, options) {
      assert.equal(nextWorkbook, workbook);
      assert.deepEqual(options, { includeStyles: true });
      return new Uint8Array([1, 2, 3]);
    }
  };

  const document = new XlsxDocument(parser, writer);
  const loaded = await document.load(new Uint8Array([10, 20]), { preserveStyles: true });
  assert.equal(loaded, workbook);

  const serialized = await document.serialize(workbook, { includeStyles: true });
  assert.deepEqual([...serialized], [1, 2, 3]);
});

test("XlsxDocument saves and loads workbook from buffer", async () => {
  const document = new XlsxDocument();
  const workbook = document.createWorkbook();
  const sheet = workbook.addWorksheet("Roundtrip");
  sheet.setCellValue(0, 0, "Hello");
  sheet.setCellValue(1, 0, 123);
  sheet.getCell("A3").setFormula("A2*2", 246);
  sheet.getCell("B1").setStyle({ bold: true, fontName: "Calibri" });

  const bytes = await document.serialize(workbook, { includeStyles: true });
  assert.equal(bytes instanceof Uint8Array, true);
  assert.equal(bytes.length > 0, true);

  const loaded = await document.load(bytes, { preserveStyles: true });
  const loadedSheet = loaded.getWorksheet("Roundtrip");
  assert.ok(loadedSheet);
  assert.equal(loadedSheet?.getCell("A1").value, "Hello");
  assert.equal(loadedSheet?.getCell("A2").value, 123);
  assert.equal(loadedSheet?.getCell("A3").formula?.expression, "A2*2");
  assert.equal(loadedSheet?.getCell("A3").formula?.result, 246);
  assert.equal(loadedSheet?.getCell("B1").style?.bold, true);
});

test("XlsxDocument saves and loads workbook by file path", async () => {
  const tmpDir = await mkdtemp(path.join(os.tmpdir(), "xlsx-js-"));
  const filePath = path.join(tmpDir, "workbook.xlsx");

  try {
    const document = new XlsxDocument();
    const workbook = document.createWorkbook();
    const sheet = workbook.addWorksheet("FilePath");
    sheet.setCellValue(0, 0, "FromPath");
    sheet.addTable({ name: "Table1", range: "A1:B3" });

    await document.writeToPath(filePath, workbook, { includeStyles: true });
    const writtenBytes = await readFile(filePath);
    assert.equal(writtenBytes.length > 0, true);

    const loaded = await document.load(filePath, { preserveStyles: true });
    const loadedSheet = loaded.getWorksheet("FilePath");
    assert.ok(loadedSheet);
    assert.equal(loadedSheet?.getCell("A1").value, "FromPath");
    assert.equal(loadedSheet?.listTables().length, 1);
    assert.equal(loadedSheet?.getTable("Table1")?.range, "A1:B3");
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
});

test("Roundtrip preserves existing drawing/chart references", async () => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
  );
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="ChartSheet" sheetId="1" r:id="rId1"/></sheets>
</workbook>`
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>before</t></is></c></row></sheetData>
  <drawing r:id="rId1"/>
</worksheet>`
  );
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`
  );
  zip.file("xl/drawings/drawing1.xml", "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>");
  zip.file("xl/charts/chart1.xml", "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>");

  const sourceBytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  const document = new XlsxDocument();
  const workbook = await document.load(sourceBytes);
  workbook.getWorksheet("ChartSheet")?.setCellValue(0, 0, "after");
  const outputBytes = await document.serialize(workbook);

  const outZip = await JSZip.loadAsync(outputBytes);
  const sheetXml = await outZip.file("xl/worksheets/sheet1.xml")?.async("string");
  assert.ok(sheetXml?.includes("<drawing r:id=\"rId1\"/>"));
  assert.ok(outZip.file("xl/drawings/drawing1.xml"));
  assert.ok(outZip.file("xl/charts/chart1.xml"));
});

test("Roundtrip preserves drawing when source sheetData is self-closing", async () => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
  );
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="ChartSheet" sheetId="1" r:id="rId1"/></sheets>
</workbook>`
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rId1"/>
</worksheet>`
  );
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`
  );
  zip.file("xl/drawings/drawing1.xml", "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>");
  zip.file("xl/charts/chart1.xml", "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>");

  const sourceBytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  const document = new XlsxDocument();
  const workbook = await document.load(sourceBytes);
  workbook.getWorksheet("ChartSheet")?.setCellValue(0, 0, "after");
  const outputBytes = await document.serialize(workbook);

  const outZip = await JSZip.loadAsync(outputBytes);
  const sheetXml = await outZip.file("xl/worksheets/sheet1.xml")?.async("string");
  assert.ok(sheetXml?.includes("<drawing r:id=\"rId1\"/>"));
  assert.ok(sheetXml?.includes("<sheetData>"));
  assert.ok(outZip.file("xl/charts/chart1.xml"));
});

test("No-op roundtrip preserves chart series range and aggregate flag", async () => {
  const chartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser>
          <c:cat><c:numRef><c:f>Sheet1!A1:B1000</c:f></c:numRef></c:cat>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:extLst><c:ext><c16r3:dataDisplayOptions16 aggregate="1"/></c:ext></c:extLst>
  </c:chart>
</c:chartSpace>`;

  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
  );
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`
  );
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`
  );
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>X</t></is></c></row></sheetData>
  <drawing r:id="rId1"/>
</worksheet>`
  );
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`
  );
  zip.file("xl/drawings/drawing1.xml", "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>");
  zip.file("xl/charts/chart1.xml", chartXml);

  const sourceBytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  const document = new XlsxDocument();
  const workbook = await document.load(sourceBytes);
  const outputBytes = await document.serialize(workbook);
  const outZip = await JSZip.loadAsync(outputBytes);
  const outChartXml = await outZip.file("xl/charts/chart1.xml")?.async("string");

  assert.equal(outChartXml, chartXml);
});

test("Chart authoring API creates line and pie charts", async () => {
  const document = new XlsxDocument();
  const workbook = document.createWorkbook();
  const sheet = workbook.addWorksheet("Sheet1");
  sheet.setCellValue(0, 0, "Category");
  sheet.setCellValue(0, 1, "Value");
  sheet.setCellValue(1, 0, "A");
  sheet.setCellValue(1, 1, 10);
  sheet.setCellValue(2, 0, "B");
  sheet.setCellValue(2, 1, 20);

  sheet.addChart({
    id: "line-1",
    type: "line",
    title: "Line Chart",
    series: [{ categories: "Sheet1!A2:A3", values: "Sheet1!B2:B3", name: "Series 1" }]
  });
  sheet.addChart({
    id: "pie-1",
    type: "pie",
    title: "Pie Chart",
    series: [{ categories: "Sheet1!A2:A3", values: "Sheet1!B2:B3", name: "Series 1" }]
  });

  const bytes = await document.serialize(workbook);
  const zip = await JSZip.loadAsync(bytes);
  const drawingXml = await zip.file("xl/drawings/drawing1.xml")?.async("string");
  const chart1Xml = await zip.file("xl/charts/chart1.xml")?.async("string");
  const chart2Xml = await zip.file("xl/charts/chart2.xml")?.async("string");
  const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");

  assert.ok(sheetXml?.includes("<drawing r:id=\"rIdXlsxJsDrawing1\"/>"));
  assert.ok(drawingXml?.includes("<c:chart r:id=\"rIdChart1\"/>"));
  assert.ok(drawingXml?.includes("<c:chart r:id=\"rIdChart2\"/>"));
  assert.ok(chart1Xml?.includes("<c:lineChart>"));
  assert.ok(chart2Xml?.includes("<c:pieChart>"));

  const loaded = await document.load(bytes);
  const loadedSheet = loaded.getWorksheet("Sheet1");
  assert.ok(loadedSheet);
  assert.equal(loadedSheet?.listCharts().length, 2);
  assert.equal(loadedSheet?.getChart("line-1")?.type, "line");
  assert.equal(loadedSheet?.getChart("pie-1")?.type, "pie");
});
