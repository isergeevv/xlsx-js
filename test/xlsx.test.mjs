import test from "node:test";
import assert from "node:assert/strict";

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
