## Classes

<dl>
<dt><a href="#Cell">Cell</a></dt>
<dd></dd>
<dt><a href="#Row">Row</a></dt>
<dd></dd>
<dt><a href="#Sheet">Sheet</a></dt>
<dd><p>A sheet in a workbook.</p>
</dd>
<dt><a href="#Workbook">Workbook</a></dt>
<dd></dd>
</dl>

<a name="Cell"></a>

## Cell
**Kind**: global class  

* [Cell](#Cell)
    * [new Cell(row, cellNode)](#new_Cell_new)
    * [.toString()](#Cell+toString) ⇒ <code>string</code>
    * [.getRow()](#Cell+getRow) ⇒ <code>[Row](#Row)</code>
    * [.getSheet()](#Cell+getSheet) ⇒ <code>[Sheet](#Sheet)</code>
    * [.getAddress()](#Cell+getAddress) ⇒ <code>string</code>
    * [.getRowNumber()](#Cell+getRowNumber) ⇒ <code>number</code>
    * [.getColumnNumber()](#Cell+getColumnNumber) ⇒ <code>number</code>
    * [.getColumnName()](#Cell+getColumnName) ⇒ <code>number</code>
    * [.getFullAddress()](#Cell+getFullAddress) ⇒ <code>string</code>
    * [.setValue(value)](#Cell+setValue) ⇒ <code>[Cell](#Cell)</code>
    * [.getRelativeCell(rowOffset, columnOffset)](#Cell+getRelativeCell) ⇒ <code>[Cell](#Cell)</code>
    * [.setFormula(formula, [calculatedValue], [sharedIndex], [sharedRef])](#Cell+setFormula) ⇒ <code>[Cell](#Cell)</code>
    * [.shareFormulaUntil(lastSharedCell)](#Cell+shareFormulaUntil) ⇒ <code>[Cell](#Cell)</code>

<a name="new_Cell_new"></a>

### new Cell(row, cellNode)
Initializes a new Cell.


| Param | Type | Description |
| --- | --- | --- |
| row | <code>[Row](#Row)</code> | The parent row. |
| cellNode | <code>Element</code> | The cell node. |

<a name="Cell+toString"></a>

### cell.toString() ⇒ <code>string</code>
Get node information.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> - The cell information.  
<a name="Cell+getRow"></a>

### cell.getRow() ⇒ <code>[Row](#Row)</code>
Gets the parent row of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Row](#Row)</code> - The parent row.  
<a name="Cell+getSheet"></a>

### cell.getSheet() ⇒ <code>[Sheet](#Sheet)</code>
Gets the parent sheet.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Sheet](#Sheet)</code> - The parent sheet.  
<a name="Cell+getAddress"></a>

### cell.getAddress() ⇒ <code>string</code>
Gets the address of the cell (e.g. "A5").

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> - The cell address.  
<a name="Cell+getRowNumber"></a>

### cell.getRowNumber() ⇒ <code>number</code>
Gets the row number of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>number</code> - The row number.  
<a name="Cell+getColumnNumber"></a>

### cell.getColumnNumber() ⇒ <code>number</code>
Gets the column number of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>number</code> - The column number.  
<a name="Cell+getColumnName"></a>

### cell.getColumnName() ⇒ <code>number</code>
Gets the column name of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>number</code> - The column name.  
<a name="Cell+getFullAddress"></a>

### cell.getFullAddress() ⇒ <code>string</code>
Gets the full address of the cell including sheet (e.g. "Sheet1!A5").

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> - The full address.  
<a name="Cell+setValue"></a>

### cell.setValue(value) ⇒ <code>[Cell](#Cell)</code>
Sets the value of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| value | <code>\*</code> | The value to set. |

<a name="Cell+getRelativeCell"></a>

### cell.getRelativeCell(rowOffset, columnOffset) ⇒ <code>[Cell](#Cell)</code>
Returns a cell with a relative position to the offsets provided.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The relative cell.  

| Param | Type | Description |
| --- | --- | --- |
| rowOffset | <code>number</code> | Offset from this.getRowNumber(). |
| columnOffset | <code>number</code> | Offset from this.getColumnNumber(). |

<a name="Cell+setFormula"></a>

### cell.setFormula(formula, [calculatedValue], [sharedIndex], [sharedRef]) ⇒ <code>[Cell](#Cell)</code>
Sets the formula for a cell (with optional pre-calculated value).

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| formula | <code>string</code> | The formula to set. |
| [calculatedValue] | <code>\*</code> | The pre-calculated value. |
| [sharedIndex] | <code>number</code> | Unique non-negative integer value to represent the formula. |
| [sharedRef] | <code>string</code> | Range of cells referencing this formala, for example: "A1:A3". |

<a name="Cell+shareFormulaUntil"></a>

### cell.shareFormulaUntil(lastSharedCell) ⇒ <code>[Cell](#Cell)</code>
If this cell is the source of a shared formula,then assign all cells from this cell to lastSharedCell its shared index.Note that lastSharedCell must share the same row or column, such that  this.getRowNumber() <= lastSharedCell.getRowNumber()      AND,  this.getColumnNumber() <= lastSharedCell.getColumnNumber()

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The shared formula source cell.  

| Param | Type | Description |
| --- | --- | --- |
| lastSharedCell | <code>\*</code> | String address or cell to share formula up until. |

<a name="Row"></a>

## Row
**Kind**: global class  

* [Row](#Row)
    * [new Row(sheet, rowNode)](#new_Row_new)
    * [.getSheet()](#Row+getSheet) ⇒ <code>[Sheet](#Sheet)</code>
    * [.getRowNumber()](#Row+getRowNumber) ⇒ <code>number</code>
    * [.getCell(columnNumber)](#Row+getCell) ⇒ <code>[Cell](#Cell)</code>

<a name="new_Row_new"></a>

### new Row(sheet, rowNode)
Initializes a new Row.


| Param | Type | Description |
| --- | --- | --- |
| sheet | <code>[Sheet](#Sheet)</code> | The parent sheet. |
| rowNode | <code>Element</code> | The row's node. |

<a name="Row+getSheet"></a>

### row.getSheet() ⇒ <code>[Sheet](#Sheet)</code>
Gets the parent sheet.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>[Sheet](#Sheet)</code> - The parent sheet.  
<a name="Row+getRowNumber"></a>

### row.getRowNumber() ⇒ <code>number</code>
Gets the row number of the row.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>number</code> - The row number.  
<a name="Row+getCell"></a>

### row.getCell(columnNumber) ⇒ <code>[Cell](#Cell)</code>
Gets the cell in the row with the provided column number.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell with the provided column number.  

| Param | Type | Description |
| --- | --- | --- |
| columnNumber | <code>number</code> | The column number. |

<a name="Sheet"></a>

## Sheet
A sheet in a workbook.

**Kind**: global class  

* [Sheet](#Sheet)
    * [new Sheet(workbook, sheetNode, sheetXML)](#new_Sheet_new)
    * [.getWorkbook()](#Sheet+getWorkbook) ⇒ <code>[Workbook](#Workbook)</code>
    * [.getName()](#Sheet+getName) ⇒ <code>string</code>
    * [.setName(name)](#Sheet+setName) ⇒ <code>undefined</code>
    * [.getRow(rowNumber)](#Sheet+getRow) ⇒ <code>[Row](#Row)</code>
    * [.getCell(address)](#Sheet+getCell) ⇒ <code>[Cell](#Cell)</code>
    * [.getCell(rowNumber, columnNumber)](#Sheet+getCell) ⇒ <code>[Cell](#Cell)</code>

<a name="new_Sheet_new"></a>

### new Sheet(workbook, sheetNode, sheetXML)
Initializes a new Sheet.


| Param | Type | Description |
| --- | --- | --- |
| workbook | <code>[Workbook](#Workbook)</code> | The parent workbook. |
| sheetNode | <code>Element</code> | The node defining the sheet metadat in the workbook.xml. |
| sheetXML | <code>Document</code> | The XML defining the sheet data in worksheets/sheetN.xml. |

<a name="Sheet+getWorkbook"></a>

### sheet.getWorkbook() ⇒ <code>[Workbook](#Workbook)</code>
Gets the parent workbook.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Workbook](#Workbook)</code> - The parent workbook.  
<a name="Sheet+getName"></a>

### sheet.getName() ⇒ <code>string</code>
Gets the name of the sheet.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>string</code> - The name of the sheet.  
<a name="Sheet+setName"></a>

### sheet.setName(name) ⇒ <code>undefined</code>
Set the name of the sheet.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  

| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | The new name of the sheet. |

<a name="Sheet+getRow"></a>

### sheet.getRow(rowNumber) ⇒ <code>[Row](#Row)</code>
Gets the row with the given number.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Row](#Row)</code> - The row with the given number.  

| Param | Type | Description |
| --- | --- | --- |
| rowNumber | <code>number</code> | The row number. |

<a name="Sheet+getCell"></a>

### sheet.getCell(address) ⇒ <code>[Cell](#Cell)</code>
Gets the cell with the given address.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| address | <code>string</code> | The address of the cell. |

<a name="Sheet+getCell"></a>

### sheet.getCell(rowNumber, columnNumber) ⇒ <code>[Cell](#Cell)</code>
Gets the cell with the given row and column numbers.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| rowNumber | <code>number</code> | The row number of the cell. |
| columnNumber | <code>number</code> | The column number of the cell. |

<a name="Workbook"></a>

## Workbook
**Kind**: global class  

* [Workbook](#Workbook)
    * [new Workbook(data)](#new_Workbook_new)
    * [.createSheet(sheetName, [index])](#Workbook+createSheet) ⇒ <code>[Sheet](#Sheet)</code>
    * [.getSheet(sheetNameOrIndex)](#Workbook+getSheet) ⇒ <code>[Sheet](#Sheet)</code>
    * [.getNamedCell(cellName)](#Workbook+getNamedCell) ⇒ <code>[Cell](#Cell)</code>
    * [.outputAsync()](#Workbook+outputAsync) ⇒ <code>Buffer</code>

<a name="new_Workbook_new"></a>

### new Workbook(data)
Initializes a new Workbook.


| Param | Type | Description |
| --- | --- | --- |
| data | <code>Buffer</code> | File buffer of the Excel workbook. |

<a name="Workbook+createSheet"></a>

### workbook.createSheet(sheetName, [index]) ⇒ <code>[Sheet](#Sheet)</code>
Create a new sheet.

**Kind**: instance method of <code>[Workbook](#Workbook)</code>  
**Returns**: <code>[Sheet](#Sheet)</code> - The new sheet.  

| Param | Type | Description |
| --- | --- | --- |
| sheetName | <code>string</code> | The name of the sheet. Must be unique. |
| [index] | <code>number</code> | The position of the sheet (0-based). Omit to insert at the end. |

<a name="Workbook+getSheet"></a>

### workbook.getSheet(sheetNameOrIndex) ⇒ <code>[Sheet](#Sheet)</code>
Gets the sheet with the provided name or index (0-based).

**Kind**: instance method of <code>[Workbook](#Workbook)</code>  
**Returns**: <code>[Sheet](#Sheet)</code> - The sheet, if found.  

| Param | Type | Description |
| --- | --- | --- |
| sheetNameOrIndex | <code>string</code> &#124; <code>number</code> | The sheet name or index. |

<a name="Workbook+getNamedCell"></a>

### workbook.getNamedCell(cellName) ⇒ <code>[Cell](#Cell)</code>
Get a named cell. (Assumes names with workbook scope pointing to single cells.)

**Kind**: instance method of <code>[Workbook](#Workbook)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell, if found.  

| Param | Type | Description |
| --- | --- | --- |
| cellName | <code>string</code> | The name of the cell. |

<a name="Workbook+outputAsync"></a>

### workbook.outputAsync() ⇒ <code>Buffer</code>
Gets the output.

**Kind**: instance method of <code>[Workbook](#Workbook)</code>  
**Returns**: <code>Buffer</code> - A node buffer for the generated Excel workbook.  
