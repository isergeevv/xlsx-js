import test from "node:test";
import assert from "node:assert/strict";

import {
  Cell,
  CellRange,
  Chart,
  Table,
  Workbook,
  Worksheet
} from "../dist/index.js";

test("Cell: default value is null and formula/style undefined", () => {
  const cell = new Cell();
  assert.equal(cell.value, null);
  assert.equal(cell.formula, undefined);
  assert.equal(cell.style, undefined);
});

test("Cell: constructor initial value", () => {
  assert.equal(new Cell(42).value, 42);
  assert.equal(new Cell("x").value, "x");
  assert.equal(new Cell(false).value, false);
  assert.ok(new Cell(new Date("2020-01-01")).value instanceof Date);
});

test("Cell: setValue replaces value and clears formula", () => {
  const cell = new Cell("a");
  cell.setFormula("SUM(A1:A2)", 99);
  assert.ok(cell.formula);
  cell.setValue("plain");
  assert.equal(cell.value, "plain");
  assert.equal(cell.formula, undefined);
});

test("Cell: setFormula stores expression and optional result", () => {
  const cell = new Cell();
  cell.setFormula("1+1");
  assert.deepEqual(cell.formula, { expression: "1+1", result: undefined });
  cell.setFormula("A1*B1", 3.14);
  assert.deepEqual(cell.formula, { expression: "A1*B1", result: 3.14 });
});

test("Cell: setStyle clones and getters reflect updates", () => {
  const cell = new Cell();
  const style = { bold: true, fontSize: 12, numFmt: "0.00" };
  cell.setStyle(style);
  assert.deepEqual(cell.style, style);
  style.bold = false;
  assert.equal(cell.style?.bold, true);
  cell.setStyle({ italic: true });
  assert.equal(cell.style?.italic, true);
  assert.equal(cell.style?.bold, undefined);
});

test("Cell: method chaining returns same instance", () => {
  const cell = new Cell();
  const chained = cell.setValue(1).setFormula("A1", 1).setStyle({});
  assert.equal(chained, cell);
});

test("Cell: onChange callback fires on mutations", () => {
  let calls = 0;
  const cell = new Cell(null, () => {
    calls += 1;
  });
  cell.setValue(1);
  cell.setFormula("A1");
  cell.setStyle({ bold: true });
  assert.equal(calls, 3);
});

test("Table: default headerRow true and totalsRow false", () => {
  const t = new Table({ name: "T", range: "A1:C3" });
  assert.equal(t.headerRow, true);
  assert.equal(t.totalsRow, false);
});

test("Table: explicit headerRow and totalsRow", () => {
  const t = new Table({ name: "T", range: "A1:A1", headerRow: false, totalsRow: true });
  assert.equal(t.headerRow, false);
  assert.equal(t.totalsRow, true);
});

test("Table: rename and range mutators chain", () => {
  const t = new Table({ name: "Old", range: "A1:B2" });
  t.rename("New").setRange("D10:F20").setHeaderRow(false).setTotalsRow(true);
  assert.equal(t.name, "New");
  assert.equal(t.range, "D10:F20");
  assert.equal(t.headerRow, false);
  assert.equal(t.totalsRow, true);
});

test("Table: rename rejects empty name", () => {
  const t = new Table({ name: "X", range: "A1:A1" });
  assert.throws(() => t.rename("   "), /cannot be empty/);
  assert.throws(() => t.rename(""), /cannot be empty/);
  assert.equal(t.name, "X");
});

test("CellRange: fromA1 and toA1 roundtrip for single row block", () => {
  const r = CellRange.fromA1("A1:C1");
  assert.deepEqual(r.start, { row: 0, col: 0 });
  assert.deepEqual(r.end, { row: 0, col: 2 });
  assert.equal(r.toA1(), "A1:C1");
});

test("CellRange: multi-row multi-column", () => {
  const r = CellRange.fromA1("B2:D5");
  assert.deepEqual(r.start, { row: 1, col: 1 });
  assert.deepEqual(r.end, { row: 4, col: 3 });
  assert.equal(r.toA1(), "B2:D5");
});

test("CellRange: column Z and AA addresses", () => {
  const r = CellRange.fromA1("Z1:AA2");
  assert.equal(r.start.col, 25);
  assert.equal(r.end.col, 26);
  assert.equal(r.toA1(), "Z1:AA2");
});

test("CellRange: getters return defensive copies", () => {
  const r = CellRange.fromA1("A1:B2");
  const s = r.start;
  s.row = 99;
  assert.deepEqual(r.start, { row: 0, col: 0 });
});

test("CellRange: invalid range throws", () => {
  assert.throws(() => CellRange.fromA1("A1"), /Invalid A1 range/);
  assert.throws(() => CellRange.fromA1(""), /Invalid A1 range/);
  assert.throws(() => CellRange.fromA1("A1:"), /Invalid A1 range/);
});

test("CellRange: invalid address throws", () => {
  assert.throws(() => CellRange.fromA1("1A:Z9"), /Invalid A1 address/);
});

test("Workbook: metadata getter returns shallow copy", () => {
  const wb = new Workbook({ createdBy: "u1", modifiedBy: "u2" });
  const m = wb.metadata;
  m.createdBy = "tamper";
  assert.equal(wb.metadata.createdBy, "u1");
});

test("Workbook: addWorksheet duplicate name throws", () => {
  const wb = new Workbook();
  wb.addWorksheet("S1");
  assert.throws(() => wb.addWorksheet("S1"), /already exists/);
});

test("Workbook: getWorksheet returns same instance", () => {
  const wb = new Workbook();
  const s = wb.addWorksheet("A");
  assert.equal(wb.getWorksheet("A"), s);
  assert.equal(wb.getWorksheet("missing"), undefined);
});

test("Workbook: removeWorksheet", () => {
  const wb = new Workbook();
  wb.addWorksheet("X");
  assert.equal(wb.removeWorksheet("X"), true);
  assert.equal(wb.removeWorksheet("X"), false);
});

test("Workbook: renameWorksheet moves key and updates sheet name", () => {
  const wb = new Workbook();
  const s = wb.addWorksheet("Old");
  wb.renameWorksheet("Old", "New");
  assert.equal(wb.getWorksheet("Old"), undefined);
  assert.equal(wb.getWorksheet("New"), s);
  assert.equal(s.name, "New");
});

test("Workbook: renameWorksheet missing source throws", () => {
  const wb = new Workbook();
  assert.throws(() => wb.renameWorksheet("Nope", "X"), /does not exist/);
});

test("Workbook: renameWorksheet duplicate target throws", () => {
  const wb = new Workbook();
  wb.addWorksheet("A");
  wb.addWorksheet("B");
  assert.throws(() => wb.renameWorksheet("A", "B"), /already exists/);
});

test("Workbook: listWorksheets order follows insertion", () => {
  const wb = new Workbook();
  wb.addWorksheet("Third");
  wb.addWorksheet("Fourth");
  const names = wb.listWorksheets().map((w) => w.name);
  assert.deepEqual(names, ["Third", "Fourth"]);
});

test("Worksheet: stable id when provided in options", () => {
  const ws = new Worksheet({ name: "N", id: "fixed-id" });
  assert.equal(ws.id, "fixed-id");
  assert.equal(ws.name, "N");
});

test("Worksheet: getCell returns same reference for same coordinates", () => {
  const ws = new Worksheet({ name: "S" });
  assert.equal(ws.getCell(5, 7), ws.getCell(5, 7));
});

test("Worksheet: setCellValue chains and getCell reflects value", () => {
  const ws = new Worksheet({ name: "S" });
  ws.setCellValue(0, 0, "a").setCellValue(0, 1, 2);
  assert.equal(ws.getCell(0, 0).value, "a");
  assert.equal(ws.getCell(0, 1).value, 2);
});

test("Worksheet: deleteCell removes storage until cell is touched again", () => {
  const ws = new Worksheet({ name: "S" });
  ws.setCellValue(1, 1, 1);
  assert.equal(ws.deleteCell(1, 1), true);
  assert.equal(ws.listCells().some((e) => e.row === 1 && e.col === 1), false);
  const cell = ws.getCell(1, 1);
  assert.equal(cell.value, null);
  assert.equal(ws.deleteCell(1, 1), true);
});

test("Worksheet: listCells includes only populated cells", () => {
  const ws = new Worksheet({ name: "S" });
  ws.setCellValue(2, 0, "x");
  ws.setCellValue(0, 3, "y");
  const entries = ws.listCells();
  assert.equal(entries.length, 2);
  const keys = new Set(entries.map((e) => `${e.row}:${e.col}`));
  assert.ok(keys.has("2:0"));
  assert.ok(keys.has("0:3"));
});

test("Worksheet: addTable duplicate throws", () => {
  const ws = new Worksheet({ name: "S" });
  ws.addTable({ name: "T1", range: "A1:B2" });
  assert.throws(() => ws.addTable({ name: "T1", range: "A1:A1" }), /already exists/);
});

test("Worksheet: table get remove list", () => {
  const ws = new Worksheet({ name: "S" });
  const t = ws.addTable({ name: "T", range: "A1:A1" });
  assert.equal(ws.getTable("T"), t);
  assert.equal(ws.listTables().length, 1);
  assert.equal(ws.removeTable("T"), true);
  assert.equal(ws.removeTable("T"), false);
});

test("Worksheet: addChart duplicate id throws", () => {
  const ws = new Worksheet({ name: "S" });
  ws.addChart({
    id: "c1",
    type: "line",
    series: [{ values: "A1:A2" }]
  });
  assert.throws(
    () =>
      ws.addChart({
        id: "c1",
        type: "pie",
        series: [{ values: "B1:B2" }]
      }),
    /already exists/
  );
});

test("Worksheet: chart get remove list", () => {
  const ws = new Worksheet({ name: "S" });
  const c = ws.addChart({ id: "x", type: "line", series: [{ values: "A1:A2" }] });
  assert.equal(ws.getChart("x"), c);
  assert.equal(ws.listCharts().length, 1);
  assert.equal(ws.removeChart("x"), true);
  assert.equal(ws.removeChart("x"), false);
});

test("Worksheet: markClean clears dirty after load simulation", () => {
  const ws = new Worksheet({ name: "S" });
  ws.setCellValue(0, 0, 1);
  assert.equal(ws.isDirty, true);
  ws.markClean();
  assert.equal(ws.isDirty, false);
});

test("Worksheet: rename marks dirty", () => {
  const ws = new Worksheet({ name: "A" });
  ws.markClean();
  ws.rename("B");
  assert.equal(ws.isDirty, true);
  assert.equal(ws.name, "B");
});

test("Chart: fixed id from options", () => {
  const c = new Chart({ id: "my-chart", type: "pie", series: [{ values: "A1:A2" }] });
  assert.equal(c.id, "my-chart");
});

test("Chart: default position matches documented anchor", () => {
  const c = new Chart({ type: "line", series: [{ values: "A1:A2" }] });
  assert.deepEqual(c.position.from, { row: 0, col: 4 });
  assert.deepEqual(c.position.to, { row: 20, col: 12 });
});

test("Chart: custom position is stored", () => {
  const pos = {
    from: { row: 1, col: 1 },
    to: { row: 10, col: 5 }
  };
  const c = new Chart({ type: "line", series: [{ values: "A1:A2" }], position: pos });
  assert.deepEqual(c.position, pos);
  pos.from.row = 99;
  assert.equal(c.position.from.row, 1);
});

test("Chart: series getter returns copies", () => {
  const c = new Chart({
    type: "line",
    series: [{ values: "A1:A2", categories: "B1:B2", name: "S1" }]
  });
  const s = c.series;
  s[0].values = "tampered";
  assert.equal(c.series[0].values, "A1:A2");
});

test("Chart: setTitle setSeries setPosition chain", () => {
  const c = new Chart({ type: "pie", series: [{ values: "A1:A2" }] });
  const out = c.setTitle("T").setSeries([{ values: "C1:C5" }]).setPosition({
    from: { row: 0, col: 0 },
    to: { row: 5, col: 5 }
  });
  assert.equal(out, c);
  assert.equal(c.title, "T");
  assert.equal(c.series[0].values, "C1:C5");
});

test("Chart: onChange fires for mutators", () => {
  let n = 0;
  const c = new Chart({ type: "line", series: [{ values: "A1:A2" }], title: "a" }, () => {
    n += 1;
  });
  c.setTitle("b");
  c.setSeries([{ values: "B1:B2" }]);
  c.setPosition({ from: { row: 0, col: 0 }, to: { row: 1, col: 1 } });
  assert.equal(n, 3);
});

test("integration: workbook with multiple sheets tables and cells", () => {
  const wb = new Workbook({ createdBy: "integration" });
  const s1 = wb.addWorksheet("Data");
  const s2 = wb.addWorksheet("Summary");
  s1.setCellValue(0, 0, "key");
  s1.addTable({ name: "tbl", range: "A1:D10", headerRow: true, totalsRow: true });
  s2.setCellValue(0, 0, new Date("2024-06-15"));
  assert.equal(wb.listWorksheets().length, 2);
  assert.equal(s1.listTables().length, 1);
  assert.ok(s2.getCell(0, 0).value instanceof Date);
});

test("Worksheet: getCell without mutation does not mark dirty", () => {
  const ws = new Worksheet({ name: "S" });
  ws.markClean();
  ws.getCell(5, 5);
  assert.equal(ws.isDirty, false);
});

test("Worksheet: setCellValue on new cell marks dirty", () => {
  const ws = new Worksheet({ name: "S" });
  ws.markClean();
  ws.setCellValue(0, 0, "x");
  assert.equal(ws.isDirty, true);
});

test("CellRange: constructor preserves addresses", () => {
  const r = new CellRange({ row: 3, col: 2 }, { row: 7, col: 9 });
  assert.deepEqual(r.start, { row: 3, col: 2 });
  assert.deepEqual(r.end, { row: 7, col: 9 });
  assert.equal(r.toA1(), "C4:J8");
});

test("Table: getters reflect independent property reads", () => {
  const t = new Table({ name: "N", range: "X1:Y2" });
  assert.equal(t.name, "N");
  assert.equal(t.range, "X1:Y2");
});
