## Classes

<dl>
<dt><a href="#Cell">Cell</a></dt>
<dd><p>A workbook cell.</p>
</dd>
<dt><a href="#Row">Row</a></dt>
<dd></dd>
<dt><a href="#Sheet">Sheet</a></dt>
<dd><p>A sheet in a workbook.</p>
</dd>
</dl>

<a name="Cell"></a>

## Cell
A workbook cell.

**Kind**: global class  

* [Cell](#Cell)
    * [new Cell(row, cellNode)](#new_Cell_new)
    * [.address()](#Cell+address) ⇒ <code>string</code>
    * [.clear()](#Cell+clear) ⇒ <code>[Cell](#Cell)</code>
    * [.columnName()](#Cell+columnName) ⇒ <code>number</code>
    * [.columnNumber()](#Cell+columnNumber) ⇒ <code>number</code>
    * [.fullAddress()](#Cell+fullAddress) ⇒ <code>string</code>
    * [.relativeCell(rowOffset, columnOffset)](#Cell+relativeCell) ⇒ <code>[Cell](#Cell)</code>
    * [.row()](#Cell+row) ⇒ <code>[Row](#Row)</code>
    * [.rowNumber()](#Cell+rowNumber) ⇒ <code>number</code>
    * [.sheet()](#Cell+sheet) ⇒ <code>[Sheet](#Sheet)</code>
    * [.value([value])](#Cell+value) ⇒ <code>\*</code> &#124; <code>[Cell](#Cell)</code>
    * [.workbook()](#Cell+workbook) ⇒ <code>Workbook</code>
    * [.formula(formula, [calculatedValue], [sharedIndex], [sharedRef])](#Cell+formula) ⇒ <code>[Cell](#Cell)</code>
    * [.shareFormulaUntil(lastSharedCell)](#Cell+shareFormulaUntil) ⇒ <code>[Cell](#Cell)</code>
    * [.toString()](#Cell+toString) ⇒ <code>string</code>

<a name="new_Cell_new"></a>

### new Cell(row, cellNode)
Initializes a new Cell.


| Param | Type | Description |
| --- | --- | --- |
| row | <code>[Row](#Row)</code> | The parent row. |
| cellNode | <code>Element</code> | The cell node. |

<a name="Cell+address"></a>

### cell.address() ⇒ <code>string</code>
Gets the address of the cell (e.g. "A5").

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> - The cell address.  
<a name="Cell+clear"></a>

### cell.clear() ⇒ <code>[Cell](#Cell)</code>
Clears the contents from the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  
<a name="Cell+columnName"></a>

### cell.columnName() ⇒ <code>number</code>
Gets the column name of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>number</code> - The column name.  
<a name="Cell+columnNumber"></a>

### cell.columnNumber() ⇒ <code>number</code>
Gets the column number of the cell (1-based).

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>number</code> - The column number.  
<a name="Cell+fullAddress"></a>

### cell.fullAddress() ⇒ <code>string</code>
Gets the full address of the cell including sheet (e.g. "Sheet1!A5").

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> - The full address.  
<a name="Cell+relativeCell"></a>

### cell.relativeCell(rowOffset, columnOffset) ⇒ <code>[Cell](#Cell)</code>
Returns a cell with a relative position given the offsets provided.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The relative cell.  

| Param | Type | Description |
| --- | --- | --- |
| rowOffset | <code>number</code> | The row offset (0 for the current row). |
| columnOffset | <code>number</code> | The column offset (0 for the current column). |

<a name="Cell+row"></a>

### cell.row() ⇒ <code>[Row](#Row)</code>
Gets the parent row of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Row](#Row)</code> - The parent row.  
<a name="Cell+rowNumber"></a>

### cell.rowNumber() ⇒ <code>number</code>
Gets the row number of the cell (1-based).

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>number</code> - The row number.  
<a name="Cell+sheet"></a>

### cell.sheet() ⇒ <code>[Sheet](#Sheet)</code>
Gets the parent sheet.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Sheet](#Sheet)</code> - The parent sheet.  
<a name="Cell+value"></a>

### cell.value([value]) ⇒ <code>\*</code> &#124; <code>[Cell](#Cell)</code>
Gets or sets the value of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>\*</code> &#124; <code>[Cell](#Cell)</code> - The value of the cell or the cell if setting.  

| Param | Type | Description |
| --- | --- | --- |
| [value] | <code>\*</code> | The value to set. |

<a name="Cell+workbook"></a>

### cell.workbook() ⇒ <code>Workbook</code>
Gets the parent workbook.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>Workbook</code> - The parent workbook.  
<a name="Cell+formula"></a>

### cell.formula(formula, [calculatedValue], [sharedIndex], [sharedRef]) ⇒ <code>[Cell](#Cell)</code>
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
If this cell is the source of a shared formula,then assign all cells from this cell to lastSharedCell its shared index.Note that lastSharedCell must share the same row or column, such that  this.rowNumber() <= lastSharedCell.rowNumber()      AND,  this.columnNumber() <= lastSharedCell.columnNumber()

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The shared formula source cell.  

| Param | Type | Description |
| --- | --- | --- |
| lastSharedCell | <code>\*</code> | String address or cell to share formula up until. |

<a name="Cell+toString"></a>

### cell.toString() ⇒ <code>string</code>
Get node information.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> - The cell information.  
<a name="Row"></a>

## Row
**Kind**: global class  

* [Row](#Row)
    * [new Row(sheet, rowNode)](#new_Row_new)
    * [.sheet()](#Row+sheet) ⇒ <code>[Sheet](#Sheet)</code>
    * [.rowNumber()](#Row+rowNumber) ⇒ <code>number</code>
    * [.cell(columnNumber)](#Row+cell) ⇒ <code>[Cell](#Cell)</code>

<a name="new_Row_new"></a>

### new Row(sheet, rowNode)
Initializes a new Row.


| Param | Type | Description |
| --- | --- | --- |
| sheet | <code>[Sheet](#Sheet)</code> | The parent sheet. |
| rowNode | <code>Element</code> | The row's node. |

<a name="Row+sheet"></a>

### row.sheet() ⇒ <code>[Sheet](#Sheet)</code>
Gets the parent sheet.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>[Sheet](#Sheet)</code> - The parent sheet.  
<a name="Row+rowNumber"></a>

### row.rowNumber() ⇒ <code>number</code>
Gets the row number of the row.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>number</code> - The row number.  
<a name="Row+cell"></a>

### row.cell(columnNumber) ⇒ <code>[Cell](#Cell)</code>
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
    * [.workbook()](#Sheet+workbook) ⇒ <code>Workbook</code>
    * [.name(name)](#Sheet+name) ⇒ <code>undefined</code>
    * [.row(rowNumber)](#Sheet+row) ⇒ <code>[Row](#Row)</code>
    * [.cell(address)](#Sheet+cell) ⇒ <code>[Cell](#Cell)</code>
    * [.cell(rowNumber, columnNumber)](#Sheet+cell) ⇒ <code>[Cell](#Cell)</code>

<a name="new_Sheet_new"></a>

### new Sheet(workbook, sheetNode, sheetXML)
Initializes a new Sheet.


| Param | Type | Description |
| --- | --- | --- |
| workbook | <code>Workbook</code> | The parent workbook. |
| sheetNode | <code>Element</code> | The node defining the sheet metadat in the workbook.xml. |
| sheetXML | <code>Document</code> | The XML defining the sheet data in worksheets/sheetN.xml. |

<a name="Sheet+workbook"></a>

### sheet.workbook() ⇒ <code>Workbook</code>
Gets the parent workbook.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>Workbook</code> - The parent workbook.  
<a name="Sheet+name"></a>

### sheet.name(name) ⇒ <code>undefined</code>
Set the name of the sheet.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  

| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | The new name of the sheet. |

<a name="Sheet+row"></a>

### sheet.row(rowNumber) ⇒ <code>[Row](#Row)</code>
Gets the row with the given number.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Row](#Row)</code> - The row with the given number.  

| Param | Type | Description |
| --- | --- | --- |
| rowNumber | <code>number</code> | The row number. |

<a name="Sheet+cell"></a>

### sheet.cell(address) ⇒ <code>[Cell](#Cell)</code>
Gets the cell with the given address.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| address | <code>string</code> | The address of the cell. |

<a name="Sheet+cell"></a>

### sheet.cell(rowNumber, columnNumber) ⇒ <code>[Cell](#Cell)</code>
Gets the cell with the given row and column numbers.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| rowNumber | <code>number</code> | The row number of the cell. |
| columnNumber | <code>number</code> | The column number of the cell. |

