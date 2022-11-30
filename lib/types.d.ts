// Type definitions for XLSX-Populate
// Project: https://github.com/dtjohnson/xlsx-populate
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped
// TypeScript Version: 3.4
export = XlsxPopulate;
declare class XlsxPopulate {
  static MIME_TYPE: string;
  static dateToNumber(date: Date): number;
  static fromBlankAsync(): Promise<XlsxPopulate.Workbook>;
  static fromDataAsync(
    data:
      | string
      | number[]
      | ArrayBuffer
      | Uint8Array
      | Buffer
      | Blob
      | Promise<any>,
    opts?: object
  ): Promise<XlsxPopulate.Workbook>;
  static fromFileAsync(
    path: string,
    opts?: any
  ): Promise<XlsxPopulate.Workbook>;
  static numberToDate(number: number): Date;
}

declare namespace XlsxPopulate {
  class Workbook {
    activeSheet(): Sheet;
    activeSheet(sheet: Sheet | string | number): Workbook;
    addSheet(name: string, indexOrBeforeSheet?: number | string | Sheet): Sheet;
    definedName(): string[];
    definedName(name: string): undefined | string | Cell | Range | Row | Column;
    definedName(
      name: string,
      refersTo: string | Cell | Range | Row | Column
    ): Workbook;
    deleteSheet(sheet: Sheet | string | number): Workbook;
    find(pattern: string | RegExp, replacement?: string | Function): boolean;
    moveSheet(
      sheet: Sheet | string | number,
      indexOrBeforeSheet?: number | string | Sheet
    ): Workbook;
    outputAsync(opts?: { password?: string }): Promise<Buffer>;
    outputAsync(opts: { type: "base64"; password?: string }): Promise<string>;
    outputAsync(opts: {
      type: "binarystring";
      password?: string;
    }): Promise<string>;
    outputAsync(opts: {
      type: "uint8array";
      password?: string;
    }): Promise<Uint8Array>;
    outputAsync(opts: {
      type: "arraybuffer";
      password?: string;
    }): Promise<ArrayBuffer>;
    outputAsync(opts: { type: "blob"; password?: string }): Promise<Blob>;
    outputAsync(opts: {
      type: "nodebuffer";
      password?: string;
    }): Promise<Buffer>;
    sheet(sheetNameOrIndex: number | string): Sheet;
    sheets(): Sheet[];
    property(name: string): any;
    property(names: string[]): { [key: string]: any };
    property(properties: { [key: string]: any }): Workbook;
    property(name: string, value: any): Workbook;
    properties(): CoreProperties;
    toFileAsync(path: string, opts?: object): Promise<void>;
    cloneSheet(
      from: Sheet,
      name: string,
      indexOrBeforeSheet?: number | string | Sheet
    ): Sheet;
  }

  class Sheet {
    _rows: Row[]; // private field

    active(): boolean;
    active(active: boolean): Sheet;
    activeCell(): Cell;
    activeCell(cell: string | Cell): Sheet;
    activeCell(rowNumber: number, columnNameOrNumber: string | number): Sheet;
    cell(address: string): Cell;
    cell(rowNumber: number, columnNameOrNumber: string | number): Cell;
    column(columnNameOrNumber: string | number): Column;
    definedName(name: string): undefined | string | Cell | Range | Row | Column;
    definedName(
      name: string,
      refersTo: string | Cell | Range | Row | Column
    ): Workbook;
    delete(): Workbook;
    find(
      pattern: string | RegExp,
      replacement?: string | Function
    ): Array<Cell>;
    gridLinesVisible(): boolean;
    gridLinesVisible(selected: boolean): Sheet;
    hidden(): boolean | string;
    hidden(hidden: boolean): Sheet;
    move(indexOrBeforeSheet?: number | string | Sheet): Sheet;
    merged(): Range[];
    merged(address: string): boolean;
    merged(address: string, merged: boolean): Sheet;
    name(): string;
    name(name: string): Sheet;
    range(address: string): Range;
    range(startCell: string | Cell, endCell: string | Cell): Range;
    range(
      startRowNumber: number,
      startColumnNameOrNumber: string | number,
      endRowNumber: number,
      endColumnNameOrNumber: string | number
    ): Range;
    autoFilter(): Sheet;
    autoFilter(range: Range): Sheet;
    rows(): Row[];
    row(rowNumber: number): Row;
    tabColor(): undefined | Color;
    tabColor(): Color | string | number;
    tabSelected(): boolean;
    tabSelected(selected: boolean): Sheet;
    usedRange(): Range | undefined;
    workbook(): Workbook;
    pageBreaks(): Object;
    verticalPageBreaks(): PageBreaks;
    horizontalPageBreaks(): PageBreaks;
    comment(address: string, comment: Comment | undefined): Sheet;
    conditionalFormatting(address: string, conditionalFormatting: ConditionalFormatting | undefined): Sheet;
    hyperlink(address: string): string | undefined;
    hyperlink(address: string, hyperlink: string, internal?: boolean): Sheet;
    hyperlink(address: string, opts: object | Cell): Sheet;
    printOptions(attributeName: string): boolean;
    printOptions(
      attributeName: string,
      attributeEnabled: undefined | boolean
    ): Sheet;
    printGridLines(): boolean;
    printGridLines(enabled: undefined | boolean): Sheet;
    panes(opts: PanesOptions): Sheet;
    freezePanes(xSplit: number, ySplit: number): Sheet;
    freezePanes(topLeftCell: string): Sheet;
    splitPanes(xSplit: number, ySplit: number): Sheet;
    resetPanes(): Sheet;
    pageMargins(attributeName: string): number;
    pageMargins(
      attributeName: string,
      attributeStringValue: undefined | number | string
    ): Sheet;
    pageMarginsPreset(): string;
    pageMarginsPreset(presetName: undefined | string): Sheet;
    pageMarginsPreset(presetName: string, presetAttributes: object): Sheet;
    sharedFormulas(): Record<string, { ref: string; formula: string }>;
    sharedFormulas(id: string): { ref: string; formula: string };
    sharedFormulas(
      id: string,
      sharedFormula: { ref: string; formula: string }
    ): Sheet;
    protected(): boolean;
    protected(
      password: string,
      options?: Partial<{
        objects: boolean;
        scenarios: boolean;
        selectLockedCells: boolean;
        selectUnlockedCells: boolean;
        formatCells: boolean;
        formatColumns: boolean;
        formatRows: boolean;
        insertColumns: boolean;
        insertRows: boolean;
        insertHyperlinks: boolean;
        deleteColumns: boolean;
        deleteRows: boolean;
        sort: boolean;
        autoFilter: boolean;
        pivotTables: boolean;
      }>
    ): Sheet;
  }

  class Row {
    address(opts?: object): string;
    cells(): Cell[];
    cell(columnNameOrNumber: string | number): Cell;
    height(): undefined | number;
    height(height: number): Row;
    hidden(): boolean;
    hidden(hidden: boolean): Row;
    rowNumber(): number;
    sheet(): Sheet;
    style(name: string): any;
    style(names: string[]): { [key: string]: any };
    style(name: string, value: any): Cell;
    style(styles: { [key: string]: any }): Cell;
    style(style: Style): Cell;
    workbook(): Workbook;
    addPageBreak(): Row;
  }

  class Cell {
    active(): boolean;
    active(active: boolean): Cell;
    address(opts?: object): string;
    column(): Column;
    clear(): Cell;
    columnName(): number;
    columnNumber(): number;
    find(pattern: string | RegExp, replacement?: string | Function): boolean;
    formula(): string;
    formula(formula: string): Cell;
    comment(comment: Comment | undefined): Cell;
    hyperlink(): string | undefined;
    hyperlink(hyperlink: string | Cell | undefined): Cell;
    hyperlink(opts: Object | Cell): Cell;
    dataValidation(): object | undefined;
    dataValidation(dataValidation: string | object | undefined): Cell;
    tap(callback: Function): Cell;
    thru(callback: Function): any;
    rangeTo(cell: Cell | string): Range;
    relativeCell(rowOffset: number, columnOffset: number): Cell;
    row(): Row;
    rowNumber(): number;
    sheet(): Sheet;
    style(name: string): any;
    style(names: string[]): { [key: string]: any };
    style(name: string, value: any): Cell;
    style(name: any[][]): Range;
    style(styles: { [key: string]: any }): Cell;
    style(style: Style): Cell;
    value(): string | boolean | number | Date | undefined;
    value(value: string | boolean | number | null | undefined): Cell;
    value(): Range;
    workbook(): Workbook;
    addHorizontalPageBreak(): Cell;
  }

  class Column {
    address(opts?: object): string;
    cell(rowNumber: number): Cell;
    columnName(): string;
    columnNumber(): number;
    hidden(): boolean;
    hidden(hidden: boolean): Column;
    sheet(): Sheet;
    style(name: string): any;
    style(names: string[]): { [key: string]: any };
    style(name: string, value: any): Cell;
    style(styles: { [key: string]: any }): Cell;
    style(style: Style): Cell;
    width(): undefined | number;
    width(width: number): Column;
    workbook(): Workbook;
    addPageBreak(): Column;
  }

  class PanesOptions {
    activePane: string;
    state: string;
    topLeftCell: string;
    xSplit: number;
    ySplit: number;
  }

  class CoreProperties {
    [key: string]: any;
  }

  class Range {
    address(opts?: object): string;
    cell(ri: number, ci: number): Cell;
    autoFilter(): Range;
    cells(): [Cell][];
    clear(): Range;
    endCell(): Cell;
    forEach(callback: Function): Range;
    formula(): string | undefined;
    formula(formula: string): Range;
    map(callback: Function): any[][];
    merged(): boolean;
    merged(merged: boolean): Range;
    dataValidation(): object | undefined;
    dataValidation(dataValidation: object | undefined): Range;
    reduce(callback: Function, initialValue?: any): any;
    sheet(): Sheet;
    startCell(): Cell;
    style(name: string): any[][];
    style(names: string[]): { [key: string]: any[][] };
    style(name: string): Range;
    style(name: string, value: any): Range;
    style(styles: { [key: string]: Function | any[][] | any }): Range;
    style(style: Style): Range;
    tap(callback: Function): Range;
    thru(callback: Function): any;
    value(): any[][];
    value(callback: Function): Range;
    value(values: any[][]): Range;
    value(value: any): Range;
    workbook(): Workbook;
  }

  class PageBreaks {
    count: number;
    list: any[];
    add(id: number): PageBreaks;
    remove(index: number): PageBreaks;
  }

  class Color {
    rgb?: string;
    theme?: number;
    tint?: number;
  }

  class Style {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean | string;
    strikethrough?: boolean;
    subscript?: boolean;
    superscript?: boolean;
    fontSize?: number;
    fontFamily?: string;
    fontColor?: Color | string;
    horizontalAlignment?: string;
    justifyLastLine?: boolean;
    indent?: number;
    verticalAlignment?: string;
    wrapText?: boolean;
    shrinkToFit?: boolean;
    textDirection?: string;
    textRotation?: number;
    angleTextCounterclockwise?: boolean;
    angleTextClockwise?: boolean;
    rotateTextUp?: boolean;
    rotateTextDown?: boolean;
    verticalText?: boolean;
    fill?: SolidFill | PatternFill | GradientFill;
    border?: Borders | Border;
    borderColor?: Color | string;
    borderStyle?: string;
    leftBorderColor?: Color | string;
    rightBorderColor?: Color | string;
    topBorderColor?: Color | string;
    bottomBorderColor?: Color | string;
    diagonalBorderColor?: Color | string;
    leftBorderStyle?: string;
    rightBorderStyle?: string;
    topBorderStyle?: string;
    bottomBorderStyle?: string;
    diagonalBorderStyle?: string;
    diagonalBorderDirection?: string;
  }

  class SolidFill {
    type: string;
    color: Color | string;
  }

  class PatternFill {
    type: string;
    pattern: string;
    foreground: Color | string;
    background: Color | string;
  }

  class Border {
    style: string;
    color: Color | string;
    direction?: string;
  }

  class Borders {
    left?: Border | string;
    right?: Border | string;
    top?: Border | string;
    bottom?: Border | string;
    diagonal?: Border | string;
  }

  class GradientFill {
    type: string;
    gradientType?: string;
    stops: {
      position: number;
      color: Color | string;
    }[];
    angle?: number;
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
  }

  class Comment {
    text: string;
    width: string;
    height: string;
  }

  class ConditionalFormatting {
    type: string;
    formula: string;
    priority: number;
    style: Style;
  }

  class FormulaError {
    error(): string;
  }
  namespace FormulaError {
    const DIV0: FormulaError;
    const NA: FormulaError;
    const NAME: FormulaError;
    const NULL: FormulaError;
    const NUM: FormulaError;
    const REF: FormulaError;
    const VALUE: FormulaError;
  }
}
