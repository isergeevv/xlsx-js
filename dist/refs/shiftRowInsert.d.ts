/**
 * Shifts A1-style references when a row is inserted before index `insertBefore0` (0-based).
 * Unqualified refs and refs to `worksheetName` are updated. External workbooks (`[...]Sheet!`) are skipped.
 *
 * Limitations: does not parse string literals, INDIRECT/R1C1, structured table refs, or defined names; may miss edge-case
 * formula tokens. Cross-sheet refs only shift when the sheet name matches `worksheetName`.
 */
export declare function shiftRefsInStringForRowInsert(input: string, worksheetName: string, insertBefore0: number): string;
//# sourceMappingURL=shiftRowInsert.d.ts.map