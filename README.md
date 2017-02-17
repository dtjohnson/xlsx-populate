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

* __browserify__ - build client-side javascript project bundle
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
gulp browserify  # outputs browser/xlsx-populate.js for web applications
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
<dt><a href="#Column">Column</a></dt>
<dd><p>A column.</p>
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
    * _instance_
        * [.address([opts])](#Cell+address) ⇒ <code>string</code>
        * [.column()](#Cell+column) ⇒ <code>[Column](#Column)</code>
        * [.clear()](#Cell+clear) ⇒ <code>[Cell](#Cell)</code>
        * [.columnName()](#Cell+columnName) ⇒ <code>number</code>
        * [.columnNumber()](#Cell+columnNumber) ⇒ <code>number</code>
        * [.formula()](#Cell+formula) ⇒ <code>string</code>
        * [.formula(formula)](#Cell+formula) ⇒ <code>[Cell](#Cell)</code>
        * [.find(pattern, [replacement])](#Cell+find) ⇒ <code>boolean</code>
        * [.groupWith(selections)](#Cell+groupWith) ⇒ <code>Group</code>
        * [.tap(callback)](#Cell+tap) ⇒ <code>[Cell](#Cell)</code>
        * [.thru(callback)](#Cell+thru) ⇒ <code>\*</code>
        * [.rangeTo(cell)](#Cell+rangeTo) ⇒ <code>Range</code>
        * [.relativeCell(rowOffset, columnOffset)](#Cell+relativeCell) ⇒ <code>[Cell](#Cell)</code>
        * [.row()](#Cell+row) ⇒ <code>[Row](#Row)</code>
        * [.rowNumber()](#Cell+rowNumber) ⇒ <code>number</code>
        * [.sheet()](#Cell+sheet) ⇒ <code>Sheet</code>
        * [.style(name)](#Cell+style) ⇒ <code>\*</code>
        * [.style(names)](#Cell+style) ⇒ <code>object.&lt;string, \*&gt;</code>
        * [.style(name, value)](#Cell+style) ⇒ <code>[Cell](#Cell)</code>
        * [.style(styles)](#Cell+style) ⇒ <code>[Cell](#Cell)</code>
        * [.value()](#Cell+value) ⇒ <code>string</code> &#124; <code>boolean</code> &#124; <code>number</code> &#124; <code>Date</code> &#124; <code>undefined</code>
        * [.value(value)](#Cell+value) ⇒ <code>[Cell](#Cell)</code>
        * [.workbook()](#Cell+workbook) ⇒ <code>Workbook</code>
    * _inner_
        * [~tapCallback](#Cell..tapCallback) ⇒ <code>undefined</code>
        * [~thruCallback](#Cell..thruCallback) ⇒ <code>\*</code>

<a name="Cell+address"></a>

#### cell.address([opts]) ⇒ <code>string</code>
Get the address of the column.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> - The address  

| Param | Type | Description |
| --- | --- | --- |
| [opts] | <code>Object</code> | Options |
| [opts.includeSheetName] | <code>boolean</code> | Include the sheet name in the address. |
| [opts.rowAnchored] | <code>boolean</code> | Anchor the row. |
| [opts.columnAnchored] | <code>boolean</code> | Anchor the column. |

<a name="Cell+column"></a>

#### cell.column() ⇒ <code>[Column](#Column)</code>
Gets the parent column of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Column](#Column)</code> - The parent column.  
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
<a name="Cell+formula"></a>

#### cell.formula() ⇒ <code>string</code>
Gets the formula in the cell. Note that if a formula was set as part of a range, the getter will return 'SHARED'. This is a limitation that may be addressed in a future release.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> - - The formula in the cell.  
<a name="Cell+formula"></a>

#### cell.formula(formula) ⇒ <code>[Cell](#Cell)</code>
Sets the formula in the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| formula | <code>string</code> | The formula to set. |

<a name="Cell+find"></a>

#### cell.find(pattern, [replacement]) ⇒ <code>boolean</code>
Find the given pattern in the cell and optionally replace it.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>boolean</code> - A flag indicating if the pattern was found.  

| Param | Type | Description |
| --- | --- | --- |
| pattern | <code>string</code> &#124; <code>RegExp</code> | The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches. |
| [replacement] | <code>string</code> &#124; <code>function</code> | The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in the cell will be replaced. |

<a name="Cell+groupWith"></a>

#### cell.groupWith(selections) ⇒ <code>Group</code>
Create a Group with this cell and other selections.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>Group</code> - The group.  

| Param | Type | Description |
| --- | --- | --- |
| selections | <code>[Cell](#Cell)</code> &#124; <code>Range</code> &#124; <code>Group</code> | The selections. |

<a name="Cell+tap"></a>

#### cell.tap(callback) ⇒ <code>[Cell](#Cell)</code>
Invoke a callback on the cell and return the cell. Useful for method chaining.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| callback | <code>[tapCallback](#Cell..tapCallback)</code> | The callback function. |

<a name="Cell+thru"></a>

#### cell.thru(callback) ⇒ <code>\*</code>
Invoke a callback on the cell and return the value provided by the callback. Useful for method chaining.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>\*</code> - The return value of the callaback.  

| Param | Type | Description |
| --- | --- | --- |
| callback | <code>[thruCallback](#Cell..thruCallback)</code> | The callback function. |

<a name="Cell+rangeTo"></a>

#### cell.rangeTo(cell) ⇒ <code>Range</code>
Create a range from this cell and another.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>Range</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| cell | <code>[Cell](#Cell)</code> | The other cell to range to. |

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
<a name="Cell+style"></a>

#### cell.style(name) ⇒ <code>\*</code>
Gets an individual style.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>\*</code> - The style.  

| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | The name of the style. |

<a name="Cell+style"></a>

#### cell.style(names) ⇒ <code>object.&lt;string, \*&gt;</code>
Gets multiple styles.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>object.&lt;string, \*&gt;</code> - Object whose keys are the style names and values are the styles.  

| Param | Type | Description |
| --- | --- | --- |
| names | <code>Array.&lt;string&gt;</code> | The names of the style. |

<a name="Cell+style"></a>

#### cell.style(name, value) ⇒ <code>[Cell](#Cell)</code>
Sets an individual style.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | The name of the style. |
| value | <code>\*</code> | The value to set. |

<a name="Cell+style"></a>

#### cell.style(styles) ⇒ <code>[Cell](#Cell)</code>
Sets multiple styles.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| styles | <code>object.&lt;string, \*&gt;</code> | Object whose keys are the style names and values are the styles to set. |

<a name="Cell+value"></a>

#### cell.value() ⇒ <code>string</code> &#124; <code>boolean</code> &#124; <code>number</code> &#124; <code>Date</code> &#124; <code>undefined</code>
Gets the value of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>string</code> &#124; <code>boolean</code> &#124; <code>number</code> &#124; <code>Date</code> &#124; <code>undefined</code> - The value of the cell.  
<a name="Cell+value"></a>

#### cell.value(value) ⇒ <code>[Cell](#Cell)</code>
Sets the value of the cell.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| value | <code>string</code> &#124; <code>boolean</code> &#124; <code>number</code> &#124; <code>Date</code> &#124; <code>null</code> &#124; <code>undefined</code> | The value to set. |

<a name="Cell+workbook"></a>

#### cell.workbook() ⇒ <code>Workbook</code>
Gets the parent workbook.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>Workbook</code> - The parent workbook.  
<a name="Cell..tapCallback"></a>

#### Cell~tapCallback ⇒ <code>undefined</code>
Callback used by tap.

**Kind**: inner typedef of <code>[Cell](#Cell)</code>  

| Param | Type | Description |
| --- | --- | --- |
| cell | <code>[Cell](#Cell)</code> | The cell |

<a name="Cell..thruCallback"></a>

#### Cell~thruCallback ⇒ <code>\*</code>
Callback used by thru.

**Kind**: inner typedef of <code>[Cell](#Cell)</code>  
**Returns**: <code>\*</code> - The value to return from thru.  

| Param | Type | Description |
| --- | --- | --- |
| cell | <code>[Cell](#Cell)</code> | The cell |

<a name="Column"></a>

### Column
A column.

**Kind**: global class  

* [Column](#Column)
    * [.address([opts])](#Column+address) ⇒ <code>string</code>
    * [.cell(rowNumber)](#Column+cell) ⇒ <code>[Cell](#Cell)</code>
    * [.columnName()](#Column+columnName) ⇒ <code>string</code>
    * [.columnNumber()](#Column+columnNumber) ⇒ <code>number</code>
    * [.hidden([hidden])](#Column+hidden) ⇒ <code>boolean</code> &#124; <code>[Column](#Column)</code>
    * [.sheet()](#Column+sheet) ⇒ <code>Sheet</code>
    * [.width([width])](#Column+width) ⇒ <code>undefined</code> &#124; <code>number</code> &#124; <code>[Column](#Column)</code>
    * [.workbook()](#Column+workbook) ⇒ <code>Workbook</code>

<a name="Column+address"></a>

#### column.address([opts]) ⇒ <code>string</code>
Get the address of the column.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>string</code> - The address  

| Param | Type | Description |
| --- | --- | --- |
| [opts] | <code>Object</code> | Options |
| [opts.includeSheetName] | <code>boolean</code> | Include the sheet name in the address. |
| [opts.anchored] | <code>boolean</code> | Anchor the address. |

<a name="Column+cell"></a>

#### column.cell(rowNumber) ⇒ <code>[Cell](#Cell)</code>
Get a cell within the column.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell in the column with the given row number.  

| Param | Type | Description |
| --- | --- | --- |
| rowNumber | <code>number</code> | The row number. |

<a name="Column+columnName"></a>

#### column.columnName() ⇒ <code>string</code>
Get the name of the column.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>string</code> - The column name.  
<a name="Column+columnNumber"></a>

#### column.columnNumber() ⇒ <code>number</code>
Get the number of the column.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>number</code> - The column number.  
<a name="Column+hidden"></a>

#### column.hidden([hidden]) ⇒ <code>boolean</code> &#124; <code>[Column](#Column)</code>
Gets or sets whether the column is hidden.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>boolean</code> &#124; <code>[Column](#Column)</code> - A flag indicating whether the column is hidden if getting, the column if setting.  

| Param | Type | Description |
| --- | --- | --- |
| [hidden] | <code>boolean</code> | A flag indicating whether to hide the column. |

<a name="Column+sheet"></a>

#### column.sheet() ⇒ <code>Sheet</code>
Get the parent sheet.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>Sheet</code> - The parent sheet.  
<a name="Column+width"></a>

#### column.width([width]) ⇒ <code>undefined</code> &#124; <code>number</code> &#124; <code>[Column](#Column)</code>
Gets or sets the width.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>undefined</code> &#124; <code>number</code> &#124; <code>[Column](#Column)</code> - The width (or undefined) if getting, the column if setting.  

| Param | Type | Description |
| --- | --- | --- |
| [width] | <code>number</code> | The width of the column. |

<a name="Column+workbook"></a>

#### column.workbook() ⇒ <code>Workbook</code>
Get the parent workbook.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>Workbook</code> - The parent workbook.  
<a name="Row"></a>

### Row
A row.

**Kind**: global class  

* [Row](#Row)
    * [.address([opts])](#Row+address) ⇒ <code>string</code>
    * [.cell(columnNameOrNumber)](#Row+cell) ⇒ <code>[Cell](#Cell)</code>
    * [.height([height])](#Row+height) ⇒ <code>undefined</code> &#124; <code>number</code> &#124; <code>[Row](#Row)</code>
    * [.hidden([hidden])](#Row+hidden) ⇒ <code>boolean</code> &#124; <code>[Row](#Row)</code>
    * [.rowNumber()](#Row+rowNumber) ⇒ <code>number</code>
    * [.sheet()](#Row+sheet) ⇒ <code>Sheet</code>
    * [.workbook()](#Row+workbook) ⇒ <code>Workbook</code>

<a name="Row+address"></a>

#### row.address([opts]) ⇒ <code>string</code>
Get the address of the row.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>string</code> - The address  

| Param | Type | Description |
| --- | --- | --- |
| [opts] | <code>Object</code> | Options |
| [opts.includeSheetName] | <code>boolean</code> | Include the sheet name in the address. |
| [opts.anchored] | <code>boolean</code> | Anchor the address. |

<a name="Row+cell"></a>

#### row.cell(columnNameOrNumber) ⇒ <code>[Cell](#Cell)</code>
Get a cell in the row.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| columnNameOrNumber | <code>string</code> &#124; <code>number</code> | The name or number of the column. |

<a name="Row+height"></a>

#### row.height([height]) ⇒ <code>undefined</code> &#124; <code>number</code> &#124; <code>[Row](#Row)</code>
Gets or sets the row height.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>undefined</code> &#124; <code>number</code> &#124; <code>[Row](#Row)</code> - The height (or undefined) if getting, the row if setting.  

| Param | Type | Description |
| --- | --- | --- |
| [height] | <code>number</code> | The height of the row. |

<a name="Row+hidden"></a>

#### row.hidden([hidden]) ⇒ <code>boolean</code> &#124; <code>[Row](#Row)</code>
Gets or sets whether the row is hidden.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>boolean</code> &#124; <code>[Row](#Row)</code> - A flag indicating whether the row is hidden if getting, the row if setting.  

| Param | Type | Description |
| --- | --- | --- |
| [hidden] | <code>boolean</code> | A flag indicating whether to hide the row. |

<a name="Row+rowNumber"></a>

#### row.rowNumber() ⇒ <code>number</code>
Gets the row number.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>number</code> - The row number.  
<a name="Row+sheet"></a>

#### row.sheet() ⇒ <code>Sheet</code>
Gets the parent sheet of the row.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>Sheet</code> - The parent sheet.  
<a name="Row+workbook"></a>

#### row.workbook() ⇒ <code>Workbook</code>
Get the parent workbook.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>Workbook</code> - The parent workbook.  

