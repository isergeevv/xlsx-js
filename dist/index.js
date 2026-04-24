import { readFile, writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

class Cell {
    constructor(value = null, onChange) {
        this._value = value;
        this._onChange = onChange;
    }
    get value() {
        return this._value;
    }
    get formula() {
        return this._formula;
    }
    get style() {
        return this._style;
    }
    setValue(value) {
        this._value = value;
        this._formula = undefined;
        this._onChange?.();
        return this;
    }
    setFormula(formulaExpression, result) {
        this._formula = { expression: formulaExpression, result };
        this._onChange?.();
        return this;
    }
    setStyle(style) {
        this._style = { ...style };
        this._onChange?.();
        return this;
    }
}

class CellRange {
    static fromA1(range) {
        const [left, right] = range.split(":");
        if (!left || !right) {
            throw new Error(`Invalid A1 range "${range}"`);
        }
        return new CellRange(CellRange._parseAddress(left), CellRange._parseAddress(right));
    }
    /** One cell, e.g. `B4` (1-based row/column in Excel, stored as 0-based in {@link CellAddress}). */
    static addressFromA1(a1) {
        return CellRange._parseAddress(a1);
    }
    static addressToA1(address) {
        return CellRange._addressToA1({ ...address });
    }
    constructor(start, end) {
        this._start = start;
        this._end = end;
    }
    get start() {
        return { ...this._start };
    }
    get end() {
        return { ...this._end };
    }
    toA1() {
        return `${CellRange._addressToA1(this._start)}:${CellRange._addressToA1(this._end)}`;
    }
    static _parseAddress(address) {
        const match = /^([A-Z]+)(\d+)$/i.exec(address.trim());
        if (!match) {
            throw new Error(`Invalid A1 address "${address}"`);
        }
        const [, colText, rowText] = match;
        return {
            row: Number(rowText) - 1,
            col: CellRange._columnToIndex(colText.toUpperCase())
        };
    }
    static _columnToIndex(column) {
        let value = 0;
        for (let i = 0; i < column.length; i += 1) {
            value = value * 26 + (column.charCodeAt(i) - 64);
        }
        return value - 1;
    }
    static _addressToA1(address) {
        return `${CellRange._indexToColumn(address.col)}${address.row + 1}`;
    }
    static _indexToColumn(colIndex) {
        let n = colIndex + 1;
        let out = "";
        while (n > 0) {
            const rem = (n - 1) % 26;
            out = String.fromCharCode(65 + rem) + out;
            n = Math.floor((n - 1) / 26);
        }
        return out;
    }
}

/** Excel .xlsx grid size (OOXML / Excel 2007+). */
const EXCEL_MAX_ROW_1BASED = 1_048_576;
const EXCEL_MAX_ROW_0BASED = EXCEL_MAX_ROW_1BASED - 1;
const EXCEL_MAX_COL_1BASED = 16_384; // XFD
const EXCEL_MAX_COL_0BASED = EXCEL_MAX_COL_1BASED - 1;

/**
 * Shifts A1-style references when a row is inserted before index `insertBefore0` (0-based).
 * Unqualified refs and refs to `worksheetName` are updated. External workbooks (`[...]Sheet!`) are skipped.
 *
 * Limitations: does not parse string literals, INDIRECT/R1C1, structured table refs, or defined names; may miss edge-case
 * formula tokens. Cross-sheet refs only shift when the sheet name matches `worksheetName`.
 */
/** Unquoted sheet segment must not swallow `[book]sheet` as the sheet name; require a normal identifier start. */
const SHEET_NAME_MATCH = /(\[[^\]]*\])?(?:(?:'((?:[^']|'')*)')|([A-Za-z0-9_][A-Za-z0-9_.]*))!((?:\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)|(?:\d+:\d+))/g;
const UNQUALIFIED_A1 = /(?<![A-Za-z0-9_$.!:])(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)(?!!)(?![A-Za-z0-9_])/g;
const UNQUALIFIED_ROW_RANGE = /(?<![A-Za-z0-9_$.!:])(\d+:\d+)(?![A-Za-z0-9_$])/g;
function sheetNamesMatch(refSheet, worksheetName) {
    return refSheet.localeCompare(worksheetName, undefined, { sensitivity: "accent" }) === 0;
}
function shiftRowOnlyRange(addr, insertBefore0) {
    const parts = addr.split(":");
    if (parts.length !== 2) {
        return addr;
    }
    const r1 = Number(parts[0]);
    const r2 = Number(parts[1]);
    if (!Number.isInteger(r1) || !Number.isInteger(r2)) {
        return addr;
    }
    const bump = (oneBased) => {
        const row0 = oneBased - 1;
        return row0 >= insertBefore0 ? oneBased + 1 : oneBased;
    };
    return `${bump(r1)}:${bump(r2)}`;
}
function shiftSingleA1(ref, insertBefore0) {
    const m = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(ref.trim());
    if (!m) {
        return ref;
    }
    const colAbs = m[1];
    const col = m[2].toUpperCase();
    const rowAbs = m[3];
    const rowStr = m[4];
    const rowNum = Number(rowStr);
    const row0 = rowNum - 1;
    if (row0 >= insertBefore0) {
        return `${colAbs}${col}${rowAbs}${rowNum + 1}`;
    }
    return `${colAbs}${col}${rowAbs}${rowStr}`;
}
function shiftRangeToken(addr, insertBefore0) {
    const trimmed = addr.trim();
    if (/^\d+:\d+$/.test(trimmed)) {
        return shiftRowOnlyRange(trimmed, insertBefore0);
    }
    const parts = trimmed.split(":");
    if (parts.length === 1) {
        return shiftSingleA1(parts[0], insertBefore0);
    }
    if (parts.length === 2) {
        return `${shiftSingleA1(parts[0], insertBefore0)}:${shiftSingleA1(parts[1], insertBefore0)}`;
    }
    return addr;
}
function replaceQualifiedSheets(input, worksheetName, insertBefore0) {
    return input.replace(SHEET_NAME_MATCH, (full, book, quotedSheet, unquotedSheet, addr) => {
        if (book) {
            return full;
        }
        const sheet = quotedSheet !== undefined ? String(quotedSheet).replace(/''/g, "'") : String(unquotedSheet ?? "").trim();
        if (!sheetNamesMatch(sheet, worksheetName)) {
            return full;
        }
        const prefixLen = full.length - addr.length;
        const prefix = full.slice(0, prefixLen);
        return prefix + shiftRangeToken(addr, insertBefore0);
    });
}
function replaceUnqualifiedA1(input, insertBefore0) {
    return input.replace(UNQUALIFIED_A1, (full) => shiftRangeToken(full, insertBefore0));
}
function replaceUnqualifiedRowRanges(input, insertBefore0) {
    return input.replace(UNQUALIFIED_ROW_RANGE, (full) => shiftRowOnlyRange(full, insertBefore0));
}
function shiftRefsInStringForRowInsert(input, worksheetName, insertBefore0) {
    let out = replaceQualifiedSheets(input, worksheetName, insertBefore0);
    out = replaceUnqualifiedA1(out, insertBefore0);
    out = replaceUnqualifiedRowRanges(out, insertBefore0);
    return out;
}

const DEFAULT_POSITION_ANCHOR = (() => {
    const from = CellRange.addressFromA1("E1");
    const to = CellRange.addressFromA1("M21");
    return { from, to };
})();
class Chart {
    constructor(options, onChange) {
        this._id = options.id ?? `chart_${Date.now()}_${Math.floor(Math.random() * 1_000_000)}`;
        this._type = options.type;
        this._title = options.title;
        this._series = options.series.map((series) => ({ ...series }));
        if (options.position) {
            this._from = { ...CellRange.addressFromA1(options.position.from) };
            this._to = { ...CellRange.addressFromA1(options.position.to) };
        }
        else {
            this._from = { ...DEFAULT_POSITION_ANCHOR.from };
            this._to = { ...DEFAULT_POSITION_ANCHOR.to };
        }
        this._onChange = onChange;
    }
    get id() {
        return this._id;
    }
    get type() {
        return this._type;
    }
    get title() {
        return this._title;
    }
    get series() {
        return this._series.map((entry) => ({ ...entry }));
    }
    get position() {
        return {
            from: CellRange.addressToA1(this._from),
            to: CellRange.addressToA1(this._to)
        };
    }
    setTitle(title) {
        this._title = title;
        this._onChange?.();
        return this;
    }
    setSeries(series) {
        this._series = series.map((entry) => ({ ...entry }));
        this._onChange?.();
        return this;
    }
    setPosition(position) {
        this._from = { ...CellRange.addressFromA1(position.from) };
        this._to = { ...CellRange.addressFromA1(position.to) };
        this._onChange?.();
        return this;
    }
    /**
     * Shifts anchor rows and series range strings when a row is inserted on `worksheetName` before
     * the **row** of A1 `beforeA1` (the column in `beforeA1` is ignored).
     */
    applyRowInsertBefore(beforeA1, worksheetName) {
        const insertBefore = CellRange.addressFromA1(beforeA1).row;
        let changed = false;
        const bumpRow = (row) => (row >= insertBefore ? row + 1 : row);
        const nextFromRow = bumpRow(this._from.row);
        const nextToRow = bumpRow(this._to.row);
        if (nextFromRow > EXCEL_MAX_ROW_0BASED || nextToRow > EXCEL_MAX_ROW_0BASED) {
            throw new Error(`Chart anchor row cannot exceed Excel maximum (${EXCEL_MAX_ROW_0BASED} zero-based, ${EXCEL_MAX_ROW_0BASED + 1} rows)`);
        }
        if (nextFromRow !== this._from.row || nextToRow !== this._to.row) {
            this._from = { row: nextFromRow, col: this._from.col };
            this._to = { row: nextToRow, col: this._to.col };
            changed = true;
        }
        const nextSeries = this._series.map((entry) => {
            const values = shiftRefsInStringForRowInsert(entry.values, worksheetName, insertBefore);
            const categories = entry.categories
                ? shiftRefsInStringForRowInsert(entry.categories, worksheetName, insertBefore)
                : undefined;
            const name = entry.name ? shiftRefsInStringForRowInsert(entry.name, worksheetName, insertBefore) : undefined;
            if (values !== entry.values || categories !== entry.categories || name !== entry.name) {
                changed = true;
            }
            return { ...entry, values, categories, name };
        });
        this._series = nextSeries;
        if (changed) {
            this._onChange?.();
        }
    }
}

class Table {
    constructor(options) {
        this._name = options.name;
        this._range = options.range;
        this._headerRow = options.headerRow ?? true;
        this._totalsRow = options.totalsRow ?? false;
    }
    get name() {
        return this._name;
    }
    get range() {
        return this._range;
    }
    get headerRow() {
        return this._headerRow;
    }
    get totalsRow() {
        return this._totalsRow;
    }
    rename(nextName) {
        this._assertName(nextName);
        this._name = nextName;
        return this;
    }
    setRange(nextRange) {
        this._range = nextRange;
        return this;
    }
    setHeaderRow(enabled) {
        this._headerRow = enabled;
        return this;
    }
    setTotalsRow(enabled) {
        this._totalsRow = enabled;
        return this;
    }
    /**
     * Grows the table range by one row at the bottom, or updates the range as if a sheet row were inserted before `at`
     * (aligned with `Worksheet.addRow` table bookkeeping). Does not move cell contents; use `Worksheet.addRow` when
     * inserting rows in populated sheets.
     */
    addRow(options) {
        const parsed = CellRange.fromA1(this._range);
        const sr = parsed.start.row;
        const sc = parsed.start.col;
        const er = parsed.end.row;
        const ec = parsed.end.col;
        if (options?.at === undefined) {
            if (er + 1 > EXCEL_MAX_ROW_0BASED) {
                throw new Error(`Table range cannot extend past row ${EXCEL_MAX_ROW_0BASED + 1}`);
            }
            this._range = new CellRange({ row: sr, col: sc }, { row: er + 1, col: ec }).toA1();
            return this;
        }
        const at = CellRange.addressFromA1(options.at).row;
        if (!Number.isInteger(at) || at < 0) {
            throw new Error("Row index must be a non-negative integer");
        }
        if (at < sr) {
            if (er + 1 > EXCEL_MAX_ROW_0BASED) {
                throw new Error(`Table range cannot extend past row ${EXCEL_MAX_ROW_0BASED + 1}`);
            }
            this._range = new CellRange({ row: sr + 1, col: sc }, { row: er + 1, col: ec }).toA1();
        }
        else if (at <= er + 1) {
            if (er + 1 > EXCEL_MAX_ROW_0BASED) {
                throw new Error(`Table range cannot extend past row ${EXCEL_MAX_ROW_0BASED + 1}`);
            }
            this._range = new CellRange({ row: sr, col: sc }, { row: er + 1, col: ec }).toA1();
        }
        return this;
    }
    _assertName(name) {
        if (!name.trim()) {
            throw new Error("Table name cannot be empty");
        }
    }
}

class Worksheet {
    constructor(options) {
        this._cells = new Map();
        this._tables = new Map();
        this._charts = new Map();
        this._id = options.id ?? `ws_${Date.now()}_${Math.floor(Math.random() * 1_000_000)}`;
        this._name = options.name;
        this._dirty = false;
    }
    get id() {
        return this._id;
    }
    get name() {
        return this._name;
    }
    get isDirty() {
        return this._dirty;
    }
    markClean() {
        this._dirty = false;
        return this;
    }
    rename(nextName) {
        this._name = nextName;
        this._dirty = true;
        return this;
    }
    getCell(a1) {
        const { row, col } = CellRange.addressFromA1(a1);
        return this._getOrCreateAt(row, col);
    }
    setCellValue(a1, value) {
        const { row, col } = CellRange.addressFromA1(a1);
        this._getOrCreateAt(row, col).setValue(value);
        return this;
    }
    deleteCell(a1) {
        const { row, col } = CellRange.addressFromA1(a1);
        Worksheet._assertAddressInGrid(row, col);
        const deleted = this._cells.delete(Worksheet._key({ row, col }));
        if (deleted) {
            this._dirty = true;
        }
        return deleted;
    }
    /**
     * Adds a logical row: either appends after the last used row (no cell moves), or inserts before the row
     * given by `options.at` (A1; column ignored) and shifts existing cells at that row and below down by one. When
     * inserting, table ranges, chart anchors/series strings, and formula text on this sheet are adjusted for the
     * new row (best-effort A1 / row-range rewriting, not a full formula parse).
     * @returns The 0-based row index of the new empty row.
     */
    addRow(options) {
        if (options?.at !== undefined) {
            const at = CellRange.addressFromA1(options.at).row;
            if (!Number.isInteger(at) || at < 0) {
                throw new Error("Row index must be a non-negative integer");
            }
            if (at > EXCEL_MAX_ROW_0BASED) {
                throw new Error(`Row index cannot exceed ${EXCEL_MAX_ROW_0BASED} (Excel max row, 0-based)`);
            }
            for (const { row } of this.listCells()) {
                if (row >= at && row >= EXCEL_MAX_ROW_0BASED) {
                    throw new Error(`Cannot insert row at ${at}: cell at row ${row} would move past Excel limit (${EXCEL_MAX_ROW_0BASED + 1} rows)`);
                }
            }
            const entries = this.listCells().filter((e) => e.row >= at);
            entries.sort((a, b) => (b.row - a.row) || (b.col - a.col));
            for (const { row, col, cell } of entries) {
                this._cells.delete(Worksheet._key({ row, col }));
                this._cells.set(Worksheet._key({ row: row + 1, col }), cell);
            }
            for (const table of this._tables.values()) {
                table.addRow({ at: options.at });
            }
            for (const chart of this._charts.values()) {
                chart.applyRowInsertBefore(options.at, this._name);
            }
            for (const { cell } of this.listCells()) {
                const f = cell.formula;
                if (f?.expression) {
                    const nextExpr = shiftRefsInStringForRowInsert(f.expression, this._name, at);
                    if (nextExpr !== f.expression) {
                        cell.setFormula(nextExpr, f.result);
                    }
                }
            }
            this._dirty = true;
            return at;
        }
        let maxRow = -1;
        for (const key of this._cells.keys()) {
            const row = Number(key.split(":")[0]);
            if (row > maxRow) {
                maxRow = row;
            }
        }
        const next = maxRow + 1;
        if (next > EXCEL_MAX_ROW_0BASED) {
            throw new Error(`Next row index ${next} exceeds Excel maximum (${EXCEL_MAX_ROW_0BASED + 1} rows)`);
        }
        return next;
    }
    /**
     * Appends a row to the table’s range (or inserts before `at` via `addRow`, which shifts the sheet), then optionally
     * writes `values` across the table’s columns on that new row.
     * @returns 0-based sheet row index of the new table row.
     */
    addTableRow(tableName, options) {
        const table = this._tables.get(tableName);
        if (!table) {
            throw new Error(`Table "${tableName}" does not exist in worksheet "${this.name}"`);
        }
        const parsed = CellRange.fromA1(table.range);
        const sr = parsed.start.row;
        const sc = parsed.start.col;
        const er = parsed.end.row;
        const ec = parsed.end.col;
        const colCount = ec - sc + 1;
        if (options?.at !== undefined) {
            const at = CellRange.addressFromA1(options.at).row;
            if (at < sr || at > er + 1) {
                throw new Error(`Table row insert "at" (${at}) must satisfy ${sr} <= at <= ${er + 1} for this table`);
            }
            this.addRow({ at: options.at });
            this._writeTableRowValues(at, sc, colCount, options.values);
            return at;
        }
        const newRow = er + 1;
        if (newRow > EXCEL_MAX_ROW_0BASED) {
            throw new Error(`Cannot append table row past Excel maximum (${EXCEL_MAX_ROW_0BASED + 1} rows)`);
        }
        table.addRow();
        this._writeTableRowValues(newRow, sc, colCount, options?.values);
        return newRow;
    }
    addTable(options) {
        if (this._tables.has(options.name)) {
            throw new Error(`Table "${options.name}" already exists in worksheet "${this.name}"`);
        }
        const t = new Table(options);
        this._tables.set(t.name, t);
        this._dirty = true;
        return t;
    }
    getTable(name) {
        return this._tables.get(name);
    }
    removeTable(name) {
        const removed = this._tables.delete(name);
        if (removed) {
            this._dirty = true;
        }
        return removed;
    }
    listTables() {
        return [...this._tables.values()];
    }
    addChart(options) {
        const chart = new Chart(options, () => {
            this._dirty = true;
        });
        if (this._charts.has(chart.id)) {
            throw new Error(`Chart "${chart.id}" already exists in worksheet "${this.name}"`);
        }
        this._charts.set(chart.id, chart);
        this._dirty = true;
        return chart;
    }
    getChart(id) {
        return this._charts.get(id);
    }
    removeChart(id) {
        const removed = this._charts.delete(id);
        if (removed) {
            this._dirty = true;
        }
        return removed;
    }
    listCharts() {
        return [...this._charts.values()];
    }
    listCells() {
        return [...this._cells.entries()].map(([key, cell]) => {
            const [rowText, colText] = key.split(":");
            return {
                row: Number(rowText),
                col: Number(colText),
                cell
            };
        });
    }
    _getOrCreateAt(row, col) {
        Worksheet._assertAddressInGrid(row, col);
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
    _writeTableRowValues(row, startCol, colCount, values) {
        if (values === undefined) {
            return;
        }
        if (Array.isArray(values)) {
            const n = Math.min(colCount, values.length);
            for (let i = 0; i < n; i += 1) {
                this.setCellValue(CellRange.addressToA1({ row, col: startCol + i }), values[i]);
            }
            return;
        }
        for (const [key, v] of Object.entries(values)) {
            const offset = Number(key);
            if (!Number.isInteger(offset) || offset < 0 || offset >= colCount) {
                throw new Error(`Table row value key "${key}" must be an integer column offset in [0, ${colCount - 1}]`);
            }
            this.setCellValue(CellRange.addressToA1({ row, col: startCol + offset }), v);
        }
    }
    static _key(address) {
        return `${address.row}:${address.col}`;
    }
    static _assertAddressInGrid(row, col) {
        if (!Number.isInteger(row) || row < 0 || row > EXCEL_MAX_ROW_0BASED) {
            throw new Error(`Row index must be an integer in [0, ${EXCEL_MAX_ROW_0BASED}] (${EXCEL_MAX_ROW_0BASED + 1} rows max)`);
        }
        if (!Number.isInteger(col) || col < 0 || col > EXCEL_MAX_COL_0BASED) {
            throw new Error(`Column index must be an integer in [0, ${EXCEL_MAX_COL_0BASED}] (${EXCEL_MAX_COL_0BASED + 1} columns max)`);
        }
    }
}

class Workbook {
    constructor(metadata = {}) {
        this._sheets = new Map();
        this._metadata = { ...metadata };
    }
    get metadata() {
        return { ...this._metadata };
    }
    addWorksheet(name) {
        if (this._sheets.has(name)) {
            throw new Error(`Worksheet "${name}" already exists`);
        }
        const sheet = new Worksheet({ name });
        this._sheets.set(name, sheet);
        return sheet;
    }
    getWorksheet(name) {
        return this._sheets.get(name);
    }
    removeWorksheet(name) {
        return this._sheets.delete(name);
    }
    listWorksheets() {
        return [...this._sheets.values()];
    }
    renameWorksheet(from, to) {
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

const _textEncoder = new TextEncoder();
const _textDecoder = new TextDecoder();
async function writeZip(entries) {
    const zip = new JSZip();
    for (const entry of entries) {
        zip.file(entry.name, entry.data);
    }
    const data = await zip.generateAsync({
        type: "uint8array",
        compression: "DEFLATE",
        compressionOptions: { level: 9 }
    });
    return data;
}
async function readZip(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const out = new Map();
    const paths = Object.keys(zip.files);
    for (const path of paths) {
        const file = zip.files[path];
        if (!file || file.dir) {
            continue;
        }
        out.set(path, await file.async("uint8array"));
    }
    return out;
}
function encodeText(value) {
    return _textEncoder.encode(value);
}
function decodeText(value) {
    return _textDecoder.decode(value);
}

const _snapshotByWorkbook = new WeakMap();
function setWorkbookSnapshot(workbook, snapshot) {
    _snapshotByWorkbook.set(workbook, snapshot);
}
function getWorkbookSnapshot(workbook) {
    return _snapshotByWorkbook.get(workbook);
}

function xmlEscape(value) {
    return value
        .split("&").join("&amp;")
        .split("<").join("&lt;")
        .split(">").join("&gt;")
        .split('"').join("&quot;")
        .split("'").join("&apos;");
}
function xmlUnescape(value) {
    return value
        .split("&lt;").join("<")
        .split("&gt;").join(">")
        .split("&quot;").join('"')
        .split("&apos;").join("'")
        .split("&amp;").join("&");
}
function readTagText(xml, tagName) {
    const regex = new RegExp(`<${tagName}[^>]*>([\\s\\S]*?)<\\/${tagName}>`, "i");
    const match = regex.exec(xml);
    return match ? xmlUnescape(match[1]) : undefined;
}
function getAttribute(tag, attributeName) {
    const regex = new RegExp(`${attributeName}="([^"]*)"`, "i");
    const match = regex.exec(tag);
    return match ? xmlUnescape(match[1]) : undefined;
}

class XlsxParser {
    async parse(input, options = {}) {
        const bytes = typeof input === "string" ? new Uint8Array(await readFile(input)) : input;
        const zipEntries = await readZip(bytes);
        const workbookXml = this._getTextEntry(zipEntries, "xl/workbook.xml");
        const workbookMeta = this._parseMetadata(zipEntries.get("xl/xlsxjs.json"));
        const workbook = new Workbook(workbookMeta.workbook);
        const workbookRelsXml = this._getTextEntry(zipEntries, "xl/_rels/workbook.xml.rels");
        const sheets = this._parseSheetEntries(workbookXml, workbookRelsXml);
        const sheetXmlByName = new Map();
        const sheetPathByName = new Map();
        for (const sheet of sheets) {
            const worksheet = workbook.addWorksheet(sheet.name);
            const xml = this._getTextEntry(zipEntries, sheet.path);
            sheetXmlByName.set(sheet.name, xml);
            sheetPathByName.set(sheet.name, sheet.path);
            this._parseWorksheetXml(xml, worksheet);
            const sheetMeta = workbookMeta.sheets.get(sheet.name);
            for (const table of sheetMeta?.tables ?? []) {
                worksheet.addTable(table);
            }
            for (const chart of sheetMeta?.charts ?? []) {
                worksheet.addChart(this._chartOptionsFromSnapshot(chart));
            }
            if (options.preserveStyles) {
                for (const styleEntry of sheetMeta?.styles ?? []) {
                    worksheet
                        .getCell(CellRange.addressToA1({ row: styleEntry.row, col: styleEntry.col }))
                        .setStyle(styleEntry.style);
                }
            }
            worksheet.markClean();
        }
        setWorkbookSnapshot(workbook, {
            entries: zipEntries,
            sheetPathByName,
            sheetXmlByName
        });
        return workbook;
    }
    _chartOptionsFromSnapshot(raw) {
        return {
            id: raw.id,
            type: raw.type,
            title: raw.title,
            series: raw.series,
            position: this._chartPositionFromSnapshot(raw.position)
        };
    }
    _chartPositionFromSnapshot(p) {
        if (p == null) {
            return undefined;
        }
        const o = p;
        if (typeof o.from === "string" && typeof o.to === "string") {
            return { from: o.from, to: o.to };
        }
        const from = o.from;
        const to = o.to;
        if (from != null &&
            to != null &&
            typeof from.row === "number" &&
            typeof from.col === "number" &&
            typeof to.row === "number" &&
            typeof to.col === "number") {
            return { from: CellRange.addressToA1(from), to: CellRange.addressToA1(to) };
        }
        return undefined;
    }
    _getTextEntry(entries, name) {
        const raw = entries.get(name);
        if (!raw) {
            throw new Error(`Missing XLSX entry: ${name}`);
        }
        return decodeText(raw);
    }
    _parseSheetEntries(workbookXml, workbookRelsXml) {
        const relPathById = this._parseWorkbookRelationshipTargets(workbookRelsXml);
        const sheetList = [];
        const sheetRegex = /<sheet\b[^>]*>/g;
        let match = sheetRegex.exec(workbookXml);
        while (match) {
            const tag = match[0];
            const name = getAttribute(tag, "name");
            const relId = getAttribute(tag, "r:id");
            const target = relId ? relPathById.get(relId) : undefined;
            if (name && target) {
                sheetList.push({ name, path: `xl/${target}` });
            }
            match = sheetRegex.exec(workbookXml);
        }
        return sheetList;
    }
    _parseWorkbookRelationshipTargets(workbookRelsXml) {
        const mapping = new Map();
        const relRegex = /<Relationship\b[^>]*>/g;
        let match = relRegex.exec(workbookRelsXml);
        while (match) {
            const tag = match[0];
            const id = getAttribute(tag, "Id");
            const target = getAttribute(tag, "Target");
            if (id && target) {
                mapping.set(id, target);
            }
            match = relRegex.exec(workbookRelsXml);
        }
        return mapping;
    }
    _parseWorksheetXml(xml, worksheet) {
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
            const cell = worksheet.getCell(CellRange.addressToA1({ row, col }));
            if (formula) {
                cell.setFormula(formula, primitive);
            }
            else {
                cell.setValue(primitive);
            }
            match = cellRegex.exec(xml);
        }
    }
    _parsePrimitive(raw, type) {
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
    _fromA1(reference) {
        const m = /^([A-Z]+)(\d+)$/i.exec(reference);
        if (!m) {
            throw new Error(`Invalid cell reference ${reference}`);
        }
        const row = Number(m[2]) - 1;
        let col = 0;
        const colText = m[1].toUpperCase();
        for (let i = 0; i < colText.length; i += 1) {
            col = col * 26 + (colText.charCodeAt(i) - 64);
        }
        return { row, col: col - 1 };
    }
    _parseMetadata(raw) {
        if (!raw) {
            return { workbook: {}, sheets: new Map() };
        }
        const parsed = JSON.parse(decodeText(raw));
        const sheets = new Map();
        for (const sheet of parsed.sheets ?? []) {
            sheets.set(sheet.name, {
                styles: sheet.styles ?? [],
                tables: sheet.tables ?? [],
                charts: sheet.charts ?? []
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

class XlsxWriter {
    async write(workbook, options = {}) {
        const worksheets = workbook.listWorksheets();
        const snapshot = getWorkbookSnapshot(workbook);
        const sheetPathByName = snapshot?.sheetPathByName ?? new Map();
        const sheetXmlByName = snapshot?.sheetXmlByName ?? new Map();
        const baseEntries = new Map(snapshot?.entries ?? []);
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
            baseEntries.set("[Content_Types].xml", encodeText(this._contentTypesXml(worksheets.length, resolvedSheetPaths, chartAssets.contentTypeOverrides)));
            baseEntries.set("_rels/.rels", encodeText(this._rootRelsXml()));
            baseEntries.set("xl/workbook.xml", encodeText(this._workbookXml(worksheets)));
            baseEntries.set("xl/_rels/workbook.xml.rels", encodeText(this._workbookRelsXml(worksheets, resolvedSheetPaths)));
        }
        baseEntries.set("xl/xlsxjs.json", encodeText(this._metadataJson(workbook, options)));
        return writeZip([...baseEntries.entries()].map(([name, data]) => ({ name, data })));
    }
    async writeToPath(path, workbook, options = {}) {
        const buffer = await this.write(workbook, options);
        await writeFile(path, buffer);
    }
    _worksheetXml(sheet, baseXml, drawingRelId) {
        const rows = new Map();
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
            }
            else if (/<sheetData\s*\/>/.test(baseXml)) {
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
    _cellXml(row, col, value, formula) {
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
    _serializePrimitive(value) {
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
    _metadataJson(workbook, options) {
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
    _sanitizeStyle(style) {
        if (!style) {
            return null;
        }
        return { ...style };
    }
    _contentTypesXml(sheetCount, sheetPathByName, extraOverrides) {
        let overrides = `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`;
        const worksheetPaths = new Set([...sheetPathByName.values()]);
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
    _rootRelsXml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
    }
    _workbookXml(worksheets) {
        const sheetsXml = worksheets
            .map((sheet, index) => `<sheet name="${xmlEscape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
            .join("");
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheetsXml}</sheets>
</workbook>`;
    }
    _workbookRelsXml(worksheets, sheetPathByName) {
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
    _columnName(col) {
        let n = col + 1;
        let out = "";
        while (n > 0) {
            const rem = (n - 1) % 26;
            out = String.fromCharCode(65 + rem) + out;
            n = Math.floor((n - 1) / 26);
        }
        return out;
    }
    _resolveSheetPaths(worksheets, existing) {
        const used = new Set(existing.values());
        const resolved = new Map();
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
    _upsertDrawingTag(xml, drawingRelId) {
        const drawingTag = `<drawing r:id="${drawingRelId}"/>`;
        if (/<drawing\b[^>]*\/>/.test(xml)) {
            return xml.replace(/<drawing\b[^>]*\/>/, drawingTag);
        }
        return xml.replace(/<\/worksheet>/, `  ${drawingTag}\n</worksheet>`);
    }
    _buildChartAssets(worksheets, sheetPathByName) {
        const bySheetPath = new Map();
        const files = [];
        const overrides = [];
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
            const chartEntries = [];
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
            const drawingRelsXml = this._drawingRelsXml(chartEntries.map((entry) => ({ id: entry.chartRelId, targetPath: entry.chartPath })));
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
    _sheetRelsPath(sheetPath) {
        const segments = sheetPath.split("/");
        const fileName = segments.length > 0 ? segments[segments.length - 1] : "sheet1.xml";
        return `xl/worksheets/_rels/${fileName}.rels`;
    }
    _sheetRelsXml(drawingRelId, drawingPath) {
        const target = drawingPath.replace(/^xl\/drawings\//, "../drawings/");
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${drawingRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="${target}"/>
</Relationships>`;
    }
    _drawingXml(anchors) {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${anchors.join("")}
</xdr:wsDr>`;
    }
    _drawingRelsXml(relations) {
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
    _chartAnchorXml(position, chartRelId, chartIndex) {
        const from = CellRange.addressFromA1(position.from);
        const to = CellRange.addressFromA1(position.to);
        return `<xdr:twoCellAnchor>
    <xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
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
    _chartXml(chart, chartIndex) {
        const seriesXml = chart.series
            .map((series, index) => {
            const nameXml = series.name ? `<c:tx><c:v>${xmlEscape(series.name)}</c:v></c:tx>` : "";
            const catXml = series.categories
                ? `<c:cat><c:strRef><c:f>${xmlEscape(series.categories)}</c:f></c:strRef></c:cat>`
                : "";
            return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${nameXml}${catXml}<c:val><c:numRef><c:f>${xmlEscape(series.values)}</c:f></c:numRef></c:val></c:ser>`;
        })
            .join("");
        const titleXml = chart.title
            ? `<c:title><c:tx><c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>${xmlEscape(chart.title)}</a:t></a:r></a:p></c:rich></c:tx></c:title>`
            : "";
        const chartTypeXml = chart.type === "pie"
            ? `<c:pieChart><c:varyColors val="1"/>${seriesXml}</c:pieChart>`
            : `<c:lineChart><c:grouping val="standard"/>${seriesXml}<c:axId val="${500000 + chartIndex}"/><c:axId val="${600000 + chartIndex}"/></c:lineChart>
<c:catAx><c:axId val="${500000 + chartIndex}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="${600000 + chartIndex}"/><c:tickLblPos val="nextTo"/></c:catAx>
<c:valAx><c:axId val="${600000 + chartIndex}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="${500000 + chartIndex}"/><c:tickLblPos val="nextTo"/></c:valAx>`;
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

class XlsxDocument {
    constructor(parser = new XlsxParser(), writer = new XlsxWriter()) {
        this._parser = parser;
        this._writer = writer;
    }
    createWorkbook() {
        return new Workbook();
    }
    async load(input, options = {}) {
        return this._parser.parse(input, options);
    }
    async serialize(workbook, options = {}) {
        return this._writer.write(workbook, options);
    }
    async writeToPath(path, workbook, options = {}) {
        return this._writer.writeToPath(path, workbook, options);
    }
}

export { Cell, CellRange, Chart, EXCEL_MAX_COL_0BASED, EXCEL_MAX_COL_1BASED, EXCEL_MAX_ROW_0BASED, EXCEL_MAX_ROW_1BASED, Table, Workbook, Worksheet, XlsxDocument, XlsxParser, XlsxWriter };
//# sourceMappingURL=index.js.map
