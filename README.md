# @isergeevv/xlsx-js

**Read, edit, and write Excel `.xlsx` files in Node.js** with a small TypeScript API: workbooks, sheets, values, styles, tables, and charts—then save to a buffer or disk.

---

## Install

The package is published to **GitHub Packages** (see your project’s npm config for `@isergeevv` if needed), then:

```bash
npm install @isergeevv/xlsx-js
```

**Requirements:** Node.js. The library ships ESM and CommonJS entry points plus TypeScript types.

---

## 1) Create a workbook and save it

Use a single `XlsxDocument` for loading and saving. For cells, public APIs on `Worksheet` use **Excel A1** address strings (e.g. `A1`, `B2`, `AA3`).

```ts
import { XlsxDocument } from "@isergeevv/xlsx-js";

const xlsx = new XlsxDocument();
const workbook = xlsx.createWorkbook();
const sheet = workbook.addWorksheet("Sheet1");

sheet.setCellValue("A1", "Hello");
sheet.setCellValue("A2", 42);
// or: sheet.getCell("A1").value

// Write to a file
await xlsx.writeToPath("./out.xlsx", workbook, { includeStyles: true });

// Or get bytes (e.g. for HTTP response)
const bytes = await xlsx.serialize(workbook, { includeStyles: true });
```

---

## 2) Open an existing `.xlsx`

Pass a **file path** or a **Uint8Array** buffer. Use **`{ preserveStyles: true }`** when you care about fonts, bold, and number formats coming back on save.

```ts
import { XlsxDocument } from "@isergeevv/xlsx-js";

const xlsx = new XlsxDocument();
const path = "./report.xlsx";

const workbook = await xlsx.load(path, { preserveStyles: true });
const sheet = workbook.getWorksheet("Sheet1");
if (sheet) {
  const label = sheet.getCell("A1").value;
  const total = sheet.getCell("B10").value;
  console.log(label, total);
}

// Round-trip: edit and write back
sheet?.setCellValue("C5", "Updated");
await xlsx.writeToPath(path, workbook, { includeStyles: true });
```

---

## 3) Cell values, formulas, and styles

| Goal | How |
|------|-----|
| **Set** a value at an address | `setCellValue("B4", value)` |
| **Read** and mutate a cell (same A1) | `getCell("B4")` then `.value`, `.setValue()`, `.setFormula(expr, result?)`, `.setStyle({ ... })` |
| Clear a stored cell | `deleteCell("B4")` |

`CellPrimitive` is `string | number | boolean | Date | null`.

**Example: style and formula**

```ts
sheet.getCell("B1").setStyle({ bold: true, fontName: "Calibri", fontSize: 11 });
sheet.getCell("A3").setFormula("A1+A2", 100);
```

---

## 4) Working with A1 **ranges** (not only single cells)

`CellRange` parses a rectangular range and converts back to a string. Handy for validation or reusing the same area as text (e.g. in chart series).

```ts
import { CellRange } from "@isergeevv/xlsx-js";

const range = CellRange.fromA1("A1:C10");
console.log(range.start, range.end); // 0-based row/col
console.log(range.toA1()); // "A1:C10"
```

Single-cell A1↔`CellAddress` helpers: `CellRange.addressFromA1("B4")`, `CellRange.addressToA1({ row: 3, col: 1 })`.

---

## 5) Excel tables and new rows

Define a **table** with an A1 **range** string, then append or insert **table rows** with values.

```ts
const sheet = workbook.getWorksheet("Data")!;

sheet.addTable({
  name: "Sales",
  range: "A1:C5",
  headerRow: true,
  totalsRow: false
});

// Append a row to the table (and sheet), filling the table’s columns
sheet.addTableRow("Sales", { values: ["Q4", 1200, 0.15] });

// Insert before the row of this A1 cell (column is ignored; same rule as addRow)
sheet.addTableRow("Sales", { at: "A3", values: { 0: "Rush", 2: 0.2 } });
```

Sheets also support **`addRow`**: append an empty row, or insert with **`{ at: "A1" }`**-style (only the **row** of the cell is used) to shift content and update tables/charts/formulas.

---

## 6) Charts: line and pie

You can **create** line and pie charts (series use **A1 range strings** like `Sheet1!A2:A10`) and **re-save** workbooks: existing chart parts are **preserved on roundtrip** when you load and write files.

```ts
sheet.addChart({
  id: "revenue",
  type: "line",
  title: "Revenue",
  series: [
    {
      name: "Series 1",
      categories: "Sheet1!A2:A5",
      values: "Sheet1!B2:B5"
    }
  ]
});

// Optional two-cell A1 anchor (default is roughly E1:M21), otherwise defaults apply
sheet.getChart("revenue")?.setPosition({
  from: "A2",
  to: "E21"
});
```

---

## 7) Workbook metadata (optional)

```ts
import { Workbook } from "@isergeevv/xlsx-js";

const workbook = new Workbook({ createdBy: "My App", createdAt: new Date() });
// … add worksheets, then save with XlsxDocument
```

---

## 8) Load / save options

- **`load(input, { preserveStyles?: boolean })`** — `input` is a path or `Uint8Array`.
- **`serialize(workbook, { includeStyles?: boolean })`** — returns `Uint8Array`.
- **`writeToPath(path, workbook, { includeStyles?: boolean })`**.

Set style-related flags to `true` if you need roundtripping of cell formatting.

---

## Current scope and limits

- Focused on **.xlsx** (Open XML) in **Node**; the API is **TypeScript-friendly** with runtime checks for Excel’s row/column limits.
- **Charts:** creating **line** and **pie** is supported; opening files with other chart types is aimed at **preservation on save**, not full in-memory editing of every series type.
- The project is under active development; for behavior details, the test suite in the repo is the most precise reference.

---

## License

[MIT](LICENSE)
