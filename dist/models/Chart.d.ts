import type { ChartOptions, ChartPosition, ChartSeriesOptions, ChartType } from "../types";
export declare class Chart {
    private readonly _id;
    private _type;
    private _title?;
    private _series;
    private _from;
    private _to;
    private readonly _onChange?;
    constructor(options: ChartOptions, onChange?: () => void);
    get id(): string;
    get type(): ChartType;
    get title(): string | undefined;
    get series(): ChartSeriesOptions[];
    get position(): ChartPosition;
    setTitle(title: string | undefined): this;
    setSeries(series: ChartSeriesOptions[]): this;
    setPosition(position: ChartPosition): this;
    /**
     * Shifts anchor rows and series range strings when a row is inserted on `worksheetName` before
     * the **row** of A1 `beforeA1` (the column in `beforeA1` is ignored).
     */
    applyRowInsertBefore(beforeA1: string, worksheetName: string): void;
}
//# sourceMappingURL=Chart.d.ts.map