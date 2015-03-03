[![view on npm](http://img.shields.io/npm/v/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![npm module downloads per month](http://img.shields.io/npm/dm/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![Build Status](https://travis-ci.org/dtjohnson/xlsx-populate.svg?branch=master)](https://travis-ci.org/dtjohnson/xlsx-populate)
[![Dependency Status](https://david-dm.org/dtjohnson/xlsx-populate.svg)](https://david-dm.org/dtjohnson/xlsx-populate)

# xlsx-populate
Node.js module to populate Excel XLSX templates. This module does not parse Excel workbooks. There are [good modules](https://github.com/SheetJS/js-xlsx) for this already. The purpose of this module is to open existing Excel XLSX workbook templates that have styling in place and populate with data.

## Installation

    $ npm install xlsx-populate

## Usage
Here is a basic example:
```js
var Workbook = require('xlsx-populate');

// Load the input workbook from file.
var workbook = Workbook.fromFileSync("./Book1.xlsx");

// Modify the workbook.
workbook.getSheet("Sheet1").getCell("A1").setValue("This is neat!");

// Write to file.
workbook.toFileSync("./out.xlsx");
```

### Getting Sheets
You can get sheets from a Workbook object by either name or index (0-based):
```js
// Get sheet with name "Sheet1".
var sheet = workbook.getSheet("Sheet1");

// Get the first sheet.
var sheet = workbook.getSheet(0);
```

### Getting Cells
You can get a cell from a sheet by either address or row and column:
```js
// Get cell "A5" by address.
var cell = sheet.getCell("A5");

// Get cell "A5" by row and column.
var cell = sheet.getCell(5, 1);
```

You can also get named cells directly from the Workbook:
```js
// Get cell named "Foo".
var cell = sheet.getNamedCell("Foo");
```

### Setting Cell Contents
You can set the cell value or formula:
```js
cell.setValue("foo");
cell.setValue(5.6);
cell.setFormula("SUM(A1:A5)");
```

### Serving from Express
You can serve the workbook with [express](http://expressjs.com/) with a route like this:
```js
router.get("/download", function (req, res) {
    // Open the workbook.
    var workbook = Workbook.fromFile("input.xlsx", function (err, workbook) {
        if (err) return res.status(500).send(err);

        // Make edits.
        workbook.getSheet(0).getCell("A1").setValue("foo");

        // Set the output file name.
        res.attachment("output.xlsx");

        // Send the workbook.
        res.send(workbook.output());
    });
});
```

## Classes
<dl>
<dt><a href="#Workbook">Workbook</a></dt>
<dd></dd>
<dt><a href="#Sheet">Sheet</a></dt>
<dd></dd>
<dt><a href="#Cell">Cell</a></dt>
<dd></dd>
</dl>
<a name="Workbook"></a>

## Workbook
**Kind**: global class

* [Workbook](#Workbook)
  * [new Workbook(data)](#new_Workbook_new)
  * _instance_
    * [.getSheet(sheetNameOrIndex)](#Workbook#getSheet) ⇒ <code>[Sheet](#Sheet)</code>
    * [.getNamedCell(cellName)](#Workbook#getNamedCell) ⇒ <code>[Cell](#Cell)</code>
    * [.output()](#Workbook#output) ⇒ <code>Buffer</code>
    * [.toFile(path, cb)](#Workbook#toFile)
    * [.toFileSync(path)](#Workbook#toFileSync)
  * _static_
    * [.fromFile(path, cb)](#Workbook.fromFile)
    * [.fromFileSync(path)](#Workbook.fromFileSync) ⇒ <code>[Workbook](#Workbook)</code>

<a name="new_Workbook_new"></a>
### new Workbook(data)
Initializes a new Workbook.


| Param | Type |
| --- | --- |
| data | <code>Buffer</code> |

<a name="Workbook#getSheet"></a>
### workbook.getSheet(sheetNameOrIndex) ⇒ <code>[Sheet](#Sheet)</code>
Gets the sheet with the provided name or index (0-based).

**Kind**: instance method of <code>[Workbook](#Workbook)</code>

| Param | Type |
| --- | --- |
| sheetNameOrIndex | <code>string</code> \| <code>number</code> |

<a name="Workbook#getNamedCell"></a>
### workbook.getNamedCell(cellName) ⇒ <code>[Cell](#Cell)</code>
Get a named cell. (Assumes names with workbook scope pointing to single cells.)

**Kind**: instance method of <code>[Workbook](#Workbook)</code>

| Param | Type |
| --- | --- |
| cellName | <code>string</code> |

<a name="Workbook#output"></a>
### workbook.output() ⇒ <code>Buffer</code>
Gets the output.

**Kind**: instance method of <code>[Workbook](#Workbook)</code>
<a name="Workbook#toFile"></a>
### workbook.toFile(path, cb)
Writes to file with the given path.

**Kind**: instance method of <code>[Workbook](#Workbook)</code>

| Param | Type |
| --- | --- |
| path | <code>string</code> |
| cb | <code>function</code> |

<a name="Workbook#toFileSync"></a>
### workbook.toFileSync(path)
Wirtes to file with the given path synchronously.

**Kind**: instance method of <code>[Workbook](#Workbook)</code>

| Param | Type |
| --- | --- |
| path | <code>string</code> |

<a name="Workbook.fromFile"></a>
### Workbook.fromFile(path, cb)
Creates a Workbook from the file with the given path.

**Kind**: static method of <code>[Workbook](#Workbook)</code>

| Param | Type |
| --- | --- |
| path | <code>string</code> |
| cb | <code>function</code> |

<a name="Workbook.fromFileSync"></a>
### Workbook.fromFileSync(path) ⇒ <code>[Workbook](#Workbook)</code>
Creates a Workbook from the file with the given path synchronously.

**Kind**: static method of <code>[Workbook](#Workbook)</code>

| Param |
| --- |
| path |

<a name="Sheet"></a>
## Sheet
**Kind**: global class

* [Sheet](#Sheet)
  * [new Sheet(workbook, name, sheetNode, sheetXML)](#new_Sheet_new)
  * [.getWorkbook()](#Sheet#getWorkbook) ⇒ <code>[Workbook](#Workbook)</code>
  * [.getName()](#Sheet#getName) ⇒ <code>string</code>
  * [.getCell()](#Sheet#getCell) ⇒ <code>[Cell](#Cell)</code>

<a name="new_Sheet_new"></a>
### new Sheet(workbook, name, sheetNode, sheetXML)
Initializes a new Sheet.


| Param | Type | Description |
| --- | --- | --- |
| workbook | <code>[Workbook](#Workbook)</code> |  |
| name | <code>string</code> |  |
| sheetNode | <code>etree.Element</code> | The node defining the sheet in the workbook.xml. |
| sheetXML | <code>etree.Element</code> |  |

<a name="Sheet#getWorkbook"></a>
### sheet.getWorkbook() ⇒ <code>[Workbook](#Workbook)</code>
Gets the parent workbook.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>
<a name="Sheet#getName"></a>
### sheet.getName() ⇒ <code>string</code>
Gets the name of the sheet.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>
<a name="Sheet#getCell"></a>
### sheet.getCell() ⇒ <code>[Cell](#Cell)</code>
Gets the cell with either the provided row and column or address.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>
<a name="Cell"></a>
## Cell
**Kind**: global class

* [Cell](#Cell)
  * [new Cell(sheet, row, column, cellNode)](#new_Cell_new)
  * [.getSheet()](#Cell#getSheet) ⇒ <code>[Sheet](#Sheet)</code>
  * [.getRow()](#Cell#getRow) ⇒ <code>number</code>
  * [.getColumn()](#Cell#getColumn) ⇒ <code>number</code>
  * [.getAddress()](#Cell#getAddress) ⇒ <code>string</code>
  * [.getFullAddress()](#Cell#getFullAddress) ⇒ <code>string</code>
  * [.setValue(value)](#Cell#setValue) ⇒ <code>[Cell](#Cell)</code>
  * [.setFormula(formula, [calculatedValue])](#Cell#setFormula) ⇒ <code>[Cell](#Cell)</code>

<a name="new_Cell_new"></a>
### new Cell(sheet, row, column, cellNode)
Initializes a new Cell.


| Param | Type |
| --- | --- |
| sheet | <code>[Sheet](#Sheet)</code> |
| row | <code>number</code> |
| column | <code>number</code> |
| cellNode | <code>etree.SubElement</code> |

<a name="Cell#getSheet"></a>
### cell.getSheet() ⇒ <code>[Sheet](#Sheet)</code>
Gets the parent sheet.

**Kind**: instance method of <code>[Cell](#Cell)</code>
<a name="Cell#getRow"></a>
### cell.getRow() ⇒ <code>number</code>
Gets the row of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>
<a name="Cell#getColumn"></a>
### cell.getColumn() ⇒ <code>number</code>
Gets the column of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>
<a name="Cell#getAddress"></a>
### cell.getAddress() ⇒ <code>string</code>
Gets the address of the cell (e.g. "A5").

**Kind**: instance method of <code>[Cell](#Cell)</code>
<a name="Cell#getFullAddress"></a>
### cell.getFullAddress() ⇒ <code>string</code>
Gets the full address of the cell including sheet (e.g. "Sheet1!A5").

**Kind**: instance method of <code>[Cell](#Cell)</code>
<a name="Cell#setValue"></a>
### cell.setValue(value) ⇒ <code>[Cell](#Cell)</code>
Sets the value of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>

| Param | Type |
| --- | --- |
| value | <code>\*</code> |

<a name="Cell#setFormula"></a>
### cell.setFormula(formula, [calculatedValue]) ⇒ <code>[Cell](#Cell)</code>
Sets the formula for a cell (with optional precalculated value).

**Kind**: instance method of <code>[Cell](#Cell)</code>

| Param | Type |
| --- | --- |
| formula | <code>string</code> |
| [calculatedValue] | <code>\*</code> |
