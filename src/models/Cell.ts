import type { CellFormula, CellStyle, CellPrimitive } from "../types";

export class Cell {
  private _value: CellPrimitive;
  private _formula?: CellFormula;
  private _style?: CellStyle;
  private readonly _onChange?: () => void;

  constructor(value: CellPrimitive = null, onChange?: () => void) {
    this._value = value;
    this._onChange = onChange;
  }

  public get value(): CellPrimitive {
    return this._value;
  }

  public get formula(): CellFormula | undefined {
    return this._formula;
  }

  public get style(): CellStyle | undefined {
    return this._style;
  }

  public setValue(value: CellPrimitive): this {
    this._value = value;
    this._formula = undefined;
    this._onChange?.();
    return this;
  }

  public setFormula(formulaExpression: string, result?: CellPrimitive): this {
    this._formula = { expression: formulaExpression, result };
    this._onChange?.();
    return this;
  }

  public setStyle(style: CellStyle): this {
    this._style = { ...style };
    this._onChange?.();
    return this;
  }
}
