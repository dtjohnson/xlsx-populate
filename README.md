[![view on npm](http://img.shields.io/npm/v/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![npm module downloads per month](http://img.shields.io/npm/dm/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![Build Status](https://travis-ci.org/dtjohnson/xlsx-populate.svg?branch=master)](https://travis-ci.org/dtjohnson/xlsx-populate)
[![Dependency Status](https://david-dm.org/dtjohnson/xlsx-populate.svg)](https://david-dm.org/dtjohnson/xlsx-populate)

# xlsx-populate
TODO

## Table of Contents
- [Styles](#styles)
- [API Reference](#api-reference)

## Styles

borderStyles:
hair
dotted
dashDotDot
dashed
mediumDashDotDot
 thin
 slantDashDot
 mediumDashDot
 mediumDashed
 medium
 thick
 double
 
```js
cell.style("border", true);
cell.style("border", "thin");
cell.style("borderStyle", "thin");
cell.style("borderColor", "ff0000");
cell.style("borderTop", true);
cell.style("borderLeft", "dotted");
cell.style("borderBottom", { style: "dashed", color: 5 });
cell.style("border", {
    top: true,
    left: "medium",
    diagonal: {
        style: "hair",
        direction: "both"
    }
});
```

## API Reference
### Classes

<dl>
<dt><a href="#Cell">Cell</a></dt>
<dd><p>A cell</p>
</dd>
<dt><a href="#Row">Row</a></dt>
<dd><p>A row.</p>
</dd>
</dl>

<a name="Cell"></a>

### Cell
A cell

**Kind**: global class  

* [Cell](#Cell)
    * [.address()](#Cell+address) ⇒ <code>string</code>
    * [.clear()](#Cell+clear) ⇒ <code>[Cell](#Cell)</code>
    * [.columnName()](#Cell+columnName) ⇒ <code>number</code>
    * [.columnNumber()](#Cell+columnNumber) ⇒ <code>number</code>
    * [.fullAddress()](#Cell+fullAddress) ⇒ <code>string</code>
    * [.relativeCell(rowOffset, columnOffset)](#Cell+relativeCell) ⇒ <code>[Cell](#Cell)</code>
    * [.row()](#Cell+row) ⇒ <code>[Row](#Row)</code>
    * [.rowNumber()](#Cell+rowNumber) ⇒ <code>number</code>
    * [.sheet()](#Cell+sheet) ⇒ <code>Sheet</code>
    * [.value([value])](#Cell+value) ⇒ <code>string</code> &#124; <code>boolean</code> &#124; <code>number</code> &#124; <code>Date</code> &#124; <code>null</code> &#124; <code>[Cell](#Cell)</code>
    * [.workbook()](#Cell+workbook) ⇒ <code>Workbook</code>

<a name="Cell+address"></a>

#### cell.address() ⇒ <code>string</code>
Gets the address of the cell (e.g. "A5").

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> - The cell address.  
<a name="Cell+clear"></a>

#### cell.clear() ⇒ <code>[Cell](#Cell)</code>
Clears the contents from the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  
<a name="Cell+columnName"></a>

#### cell.columnName() ⇒ <code>number</code>
Gets the column name of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>number</code> - The column name.  
<a name="Cell+columnNumber"></a>

#### cell.columnNumber() ⇒ <code>number</code>
Gets the column number of the cell (1-based).

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>number</code> - The column number.  
<a name="Cell+fullAddress"></a>

#### cell.fullAddress() ⇒ <code>string</code>
Gets the full address of the cell including sheet (e.g. "Sheet1!A5").

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> - The full address.  
<a name="Cell+relativeCell"></a>

#### cell.relativeCell(rowOffset, columnOffset) ⇒ <code>[Cell](#Cell)</code>
Returns a cell with a relative position given the offsets provided.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The relative cell.  

| Param | Type | Description |
| --- | --- | --- |
| rowOffset | <code>number</code> | The row offset (0 for the current row). |
| columnOffset | <code>number</code> | The column offset (0 for the current column). |

<a name="Cell+row"></a>

#### cell.row() ⇒ <code>[Row](#Row)</code>
Gets the parent row of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Row](#Row)</code> - The parent row.  
<a name="Cell+rowNumber"></a>

#### cell.rowNumber() ⇒ <code>number</code>
Gets the row number of the cell (1-based).

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>number</code> - The row number.  
<a name="Cell+sheet"></a>

#### cell.sheet() ⇒ <code>Sheet</code>
Gets the parent sheet.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>Sheet</code> - The parent sheet.  
<a name="Cell+value"></a>

#### cell.value([value]) ⇒ <code>string</code> &#124; <code>boolean</code> &#124; <code>number</code> &#124; <code>Date</code> &#124; <code>null</code> &#124; <code>[Cell](#Cell)</code>
Gets or sets the value of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> &#124; <code>boolean</code> &#124; <code>number</code> &#124; <code>Date</code> &#124; <code>null</code> &#124; <code>[Cell](#Cell)</code> - The value of the cell or the cell if setting.  

| Param | Type | Description |
| --- | --- | --- |
| [value] | <code>string</code> &#124; <code>boolean</code> &#124; <code>number</code> &#124; <code>Date</code> &#124; <code>null</code> &#124; <code>undefined</code> | The value to set. |

<a name="Cell+workbook"></a>

#### cell.workbook() ⇒ <code>Workbook</code>
Gets the parent workbook.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>Workbook</code> - The parent workbook.  
<a name="Row"></a>

### Row
A row.

**Kind**: global class  
