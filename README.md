[![view on npm](http://img.shields.io/npm/v/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![npm module downloads per month](http://img.shields.io/npm/dm/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![Build Status](https://travis-ci.org/dtjohnson/xlsx-populate.svg?branch=master)](https://travis-ci.org/dtjohnson/xlsx-populate)
[![Dependency Status](https://david-dm.org/dtjohnson/xlsx-populate.svg)](https://david-dm.org/dtjohnson/xlsx-populate)

# xlsx-populate
TODO

## Table of Contents
- [Setup Development Environment](#setup-development-environment)
  * [Install node and gulp globally](#install-node-and-gulp-globally)
  * [Git clone the project](#git-clone-the-project)
  * [Install xlsx-populate libraries](#install-xlsx-populate-libraries)
  * [Gulp tasks](#gulp-tasks)
- [Styles](#styles)
- [API Reference](#api-reference)

## Setup Development Environment

To contribute, ensure that npm (node package manager) and git are installed. Then continue with the following instructions.

### Install node and gulp globally
```bash
npm install --global node gulp
```

### Git clone the project
```bash
git clone git@github.com:dtjohnson/xlsx-populate.git
cd xlsx-populate
```

### Install xlsx-populate libraries
```bash
npm install
npm install --only=dev # Install dev tools
alias node="node --harmony"  # Run node in ES6 mode
```

### Gulp tasks

* __browserify__ - build client-side javascript xlsx-populate bundle
* __lint__ - check project source code style
* __unit__ - unit test project
* __blank__ - build blank xlsx files for default load
* __docs__ - build docs: generate README.md from docs/template.md and source code
* __test__ - run lint and unit test project
* __watch__ - listen for new project changes and then run associated gulp task
* __default__ - run all gulp tasks

Please review [gulp documentation](https://github.com/gulpjs/gulp) to learn more. Here are a few examples:

```
gulp lint  # checks code style
gulp browserify  # outputs bundle.js for web development
```

## Styles

* bold: Boolean
* italic: Boolean
* underline: Boolean or 'double'
* strikethough: Boolean
* subscript: Boolean
* superscript: Boolean
* fontSize: Number > 0
* fontFamily: String
* fontColor: hex String or theme Number
* fontTint: Number [-1, 1] The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means 100% darken and 1.0 means 100% lighten. Also, 0.0 means no change.
* horizontalAlignment: left, center, right, fill, justify, centerContinuous, distributed
* justifyLastLine: Boolean (akak 'Justified Distributed'. Only applies when horizontalAlignment === 'distributed') A boolean value indicating if the cells justified or distributed alignment should be used on the last line of text. (This is typical for East Asian alignments but not typical in other contexts.)
* indent: Number > 0
* verticalAlignment: top, center, bottom, justify, distributed
* wrapText: Boolean
* shrinkToFit: Boolean
* textDirection: 'left-to-right', 'right-to-left'
* textRotation: Number [-90, 90] counter clockwise rotation (negatives are clockwise)
* angleTextCounterclockwise: Boolean. textRotation = 45
* angleTextClockwise: Boolean. textRotation = -45
* rotateTextUp: Boolean. textRotation = 90
* rotateTextDown: Boolean. textRotation = -90
* verticalText: Boolean. Special rotation that shows text vertical but individual letters are oriented normally 
* fill pattern: gray125, darkGray, mediumGray, lightGray, gray0625, darkHorizontal, darkVertical, darkDown, darkUp, darkGrid, darkTrellis, lightHorizontal, lightVertical, lightDown, lightUp, lightGrid, lightTrellis
* path gradient: A box is drawn between top, left, right, and bottom. That is used to draw gradient
* borderStyle: hair, dotted, dashDotDot, dashed, mediumDashDotDot, thin, slantDashDot, mediumDashDot, mediumDashed, medium, thick, double

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

