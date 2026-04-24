import type { CellAddress } from "../types";
export declare class CellRange {
    private readonly _start;
    private readonly _end;
    static fromA1(range: string): CellRange;
    /** One cell, e.g. `B4` (1-based row/column in Excel, stored as 0-based in {@link CellAddress}). */
    static addressFromA1(a1: string): CellAddress;
    static addressToA1(address: CellAddress): string;
    constructor(start: CellAddress, end: CellAddress);
    get start(): CellAddress;
    get end(): CellAddress;
    toA1(): string;
    private static _parseAddress;
    private static _columnToIndex;
    private static _addressToA1;
    private static _indexToColumn;
}
//# sourceMappingURL=CellRange.d.ts.map