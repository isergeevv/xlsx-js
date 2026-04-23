/**
 * Shifts A1-style references when a row is inserted before index `insertBefore0` (0-based).
 * Unqualified refs and refs to `worksheetName` are updated. External workbooks (`[...]Sheet!`) are skipped.
 *
 * Limitations: does not parse string literals, INDIRECT/R1C1, structured table refs, or defined names; may miss edge-case
 * formula tokens. Cross-sheet refs only shift when the sheet name matches `worksheetName`.
 */

/** Unquoted sheet segment must not swallow `[book]sheet` as the sheet name; require a normal identifier start. */
const SHEET_NAME_MATCH =
  /(\[[^\]]*\])?(?:(?:'((?:[^']|'')*)')|([A-Za-z0-9_][A-Za-z0-9_.]*))!((?:\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)|(?:\d+:\d+))/g;

const UNQUALIFIED_A1 =
  /(?<![A-Za-z0-9_$.!:])(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)(?!!)(?![A-Za-z0-9_])/g;

const UNQUALIFIED_ROW_RANGE = /(?<![A-Za-z0-9_$.!:])(\d+:\d+)(?![A-Za-z0-9_$])/g;

function sheetNamesMatch(refSheet: string, worksheetName: string): boolean {
  return refSheet.localeCompare(worksheetName, undefined, { sensitivity: "accent" }) === 0;
}

function shiftRowOnlyRange(addr: string, insertBefore0: number): string {
  const parts = addr.split(":");
  if (parts.length !== 2) {
    return addr;
  }
  const r1 = Number(parts[0]);
  const r2 = Number(parts[1]);
  if (!Number.isInteger(r1) || !Number.isInteger(r2)) {
    return addr;
  }
  const bump = (oneBased: number): number => {
    const row0 = oneBased - 1;
    return row0 >= insertBefore0 ? oneBased + 1 : oneBased;
  };
  return `${bump(r1)}:${bump(r2)}`;
}

function shiftSingleA1(ref: string, insertBefore0: number): string {
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

function shiftRangeToken(addr: string, insertBefore0: number): string {
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

function replaceQualifiedSheets(input: string, worksheetName: string, insertBefore0: number): string {
  return input.replace(SHEET_NAME_MATCH, (full, book, quotedSheet, unquotedSheet, addr) => {
    if (book) {
      return full;
    }
    const sheet =
      quotedSheet !== undefined ? String(quotedSheet).replace(/''/g, "'") : String(unquotedSheet ?? "").trim();
    if (!sheetNamesMatch(sheet, worksheetName)) {
      return full;
    }
    const prefixLen = full.length - addr.length;
    const prefix = full.slice(0, prefixLen);
    return prefix + shiftRangeToken(addr, insertBefore0);
  });
}

function replaceUnqualifiedA1(input: string, insertBefore0: number): string {
  return input.replace(UNQUALIFIED_A1, (full) => shiftRangeToken(full, insertBefore0));
}

function replaceUnqualifiedRowRanges(input: string, insertBefore0: number): string {
  return input.replace(UNQUALIFIED_ROW_RANGE, (full) => shiftRowOnlyRange(full, insertBefore0));
}

export function shiftRefsInStringForRowInsert(
  input: string,
  worksheetName: string,
  insertBefore0: number
): string {
  let out = replaceQualifiedSheets(input, worksheetName, insertBefore0);
  out = replaceUnqualifiedA1(out, insertBefore0);
  out = replaceUnqualifiedRowRanges(out, insertBefore0);
  return out;
}
