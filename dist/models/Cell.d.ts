import type { CellFormula, CellStyle, CellPrimitive } from "../types";
export declare class Cell {
    private _value;
    private _formula?;
    private _style?;
    private readonly _onChange?;
    constructor(value?: CellPrimitive, onChange?: () => void);
    get value(): CellPrimitive;
    get formula(): CellFormula | undefined;
    get style(): CellStyle | undefined;
    setValue(value: CellPrimitive): this;
    setFormula(formulaExpression: string, result?: CellPrimitive): this;
    setStyle(style: CellStyle): this;
}
//# sourceMappingURL=Cell.d.ts.map