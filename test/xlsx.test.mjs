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
  assert.equal(worksheet.getCell(0, 0).value, "hello");
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

test("XlsxDocument delegates load/save operations", async () => {
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

  const saved = await document.save(workbook, { includeStyles: true });
  assert.deepEqual([...saved], [1, 2, 3]);
});

test("XlsxDocument saves and loads workbook from buffer", async () => {
  const document = new XlsxDocument();
  const workbook = document.createWorkbook();
  const sheet = workbook.addWorksheet("Roundtrip");
  sheet.setCellValue(0, 0, "Hello");
  sheet.setCellValue(1, 0, 123);
  sheet.getCell(2, 0).setFormula("A2*2", 246);
  sheet.getCell(0, 1).setStyle({ bold: true, fontName: "Calibri" });

  const bytes = await document.save(workbook, { includeStyles: true });
  assert.equal(bytes instanceof Uint8Array, true);
  assert.equal(bytes.length > 0, true);

  const loaded = await document.load(bytes, { preserveStyles: true });
  const loadedSheet = loaded.getWorksheet("Roundtrip");
  assert.ok(loadedSheet);
  assert.equal(loadedSheet?.getCell(0, 0).value, "Hello");
  assert.equal(loadedSheet?.getCell(1, 0).value, 123);
  assert.equal(loadedSheet?.getCell(2, 0).formula?.expression, "A2*2");
  assert.equal(loadedSheet?.getCell(2, 0).formula?.result, 246);
  assert.equal(loadedSheet?.getCell(0, 1).style?.bold, true);
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

    await document.saveToPath(filePath, workbook, { includeStyles: true });
    const writtenBytes = await readFile(filePath);
    assert.equal(writtenBytes.length > 0, true);

    const loaded = await document.load(filePath, { preserveStyles: true });
    const loadedSheet = loaded.getWorksheet("FilePath");
    assert.ok(loadedSheet);
    assert.equal(loadedSheet?.getCell(0, 0).value, "FromPath");
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
  const outputBytes = await document.save(workbook);

  const outZip = await JSZip.loadAsync(outputBytes);
  const sheetXml = await outZip.file("xl/worksheets/sheet1.xml")?.async("string");
  assert.ok(sheetXml?.includes("<drawing r:id=\"rId1\"/>"));
  assert.ok(outZip.file("xl/drawings/drawing1.xml"));
  assert.ok(outZip.file("xl/charts/chart1.xml"));
});
