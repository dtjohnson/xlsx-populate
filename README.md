[![view on npm](http://img.shields.io/npm/v/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![npm module downloads per month](http://img.shields.io/npm/dm/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![Build Status](https://travis-ci.org/dtjohnson/xlsx-populate.svg?branch=master)](https://travis-ci.org/dtjohnson/xlsx-populate)
[![Dependency Status](https://david-dm.org/dtjohnson/xlsx-populate.svg)](https://david-dm.org/dtjohnson/xlsx-populate)

# xlsx-populate
Excel XLSX parser/generator written in JavaScript with Node.js and browser support, jQuery/d3-style method chaining, and a focus on keeping existing workbook features and styles in tact.

## Table of Contents
- [Installation](#installation)
  * [Node.js](#nodejs)
  * [Browser](#browser)
- [Usage](#usage)
  * [Populating Data](#populating-data)
  * [Parsing Data](#parsing-data)
  * [Method Chaining](#method-chaining)
  * [Ranges](#ranges)
  * [Rows and Columns](#rows-and-columns)
  * [Styles](#styles)
  * [Dates](#dates)
  * [Serving from Express](#serving-from-express)
  * [Browser Usage](#browser-usage)
- [Setup Development Environment](#setup-development-environment)
  * [Install node and gulp globally](#install-node-and-gulp-globally)
  * [Git clone the project](#git-clone-the-project)
  * [Install xlsx-populate libraries](#install-xlsx-populate-libraries)
  * [Gulp tasks](#gulp-tasks)
- [Style Reference](#style-reference)
- [API Reference](#api-reference)

## Installation

### Node.js
```
npm install xlsx-populate
```
Note that xlsx-populate uses ES6 features so only Node.js v4+ is supported.

### Browser

xlsx-populate is written first for Node.js. We use [browserify](http://browserify.org/) and [babelify](https://github.com/babel/babelify) to transpile and pack up the module for use in the browser.

You have a number of options to include the code in the browser. You can download the combined, minified code from the browser directory in this repository or you can install with bower:
```
bower install xlsx-populate
```
After including the module in the browser, it is available globally as `Workbook`.

Alternatively, you can require this module using [browserify](http://browserify.org/). Since xlsx-populate uses ES6 features, you will also need to use [babelify](https://github.com/babel/babelify) with [babel-preset-es2015](https://www.npmjs.com/package/babel-preset-es2015).

## Usage

### Populating Data

Here is a basic example:
```js
const Workbook = require('xlsx-populate');

// Load a new blank workbook
Workbook.fromBlankAsync()
    .then(workbook => {
        // Modify the workbook.
        workbook.sheet("Sheet1").cell("A1").value("This is neat!");
        
        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    });
```

### Parsing Data

You can pull data out of existing workbooks using `value` as a getter without any arguments:
```js
const Workbook = require('xlsx-populate');

// Load an existing workbook
Workbook.fromFileAsync("./Book1.xlsx")
    .then(workbook => {
        // Modify the workbook.
        const value = workbook.sheet("Sheet1").cell("A1").value();
        
        // Log the value.
        console.log(value);
    });
```

### Method Chaining

TODO

### Ranges
xlsx-populate also supports ranges of cells to allow parsing/manipulate of multiple cells at once.
```js
const r = workbook.sheet(0).range("A1:C3");

// Set all cell values to the same value:
r.values(5);

// Set the values using a 2D array:
r.values([
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
]);

// Set the values using a callback function:
r.values((cell, ri, ci, range) => Math.random());
```

A common use case is to simply pull all of the values out all at once. You can easily do that with the [Sheet.usedRange](#Sheet+usedRange) method.
```js
// Get 2D array of all values in the worksheet.
const values = workbook.sheet("Sheet1").usedRange().values();
```

### Rows and Columns

TODO

### Styles
xlsx-populate supports a wide range of cell formatting. See the [Style Reference](#style-reference) for the various options.

To set/set a cell style:
```js
// Set a single style
cell.style("bold", true);

// Set multiple styles
cell.style({ bold: true, italic: true });

// Get a single style
const bold = cell.style("bold"); // true
 
// Get multiple styles
const styles = cell.style(["bold", "italic"]); // { bold: true, italic: true } 
```

Similarly for ranges:
```js
// Set all cells in range with a single style
range.style("bold", true);

// Set with a 2D array
range.style("bold", [[true, false], [false, true]]);

// Set with a callback function
range.style("bold", (cell, ri, ci, range) => Math.random() > 0.5);

// Set multiple styles using any combination
range.style({
    bold: true,
    italic: [[true, false], [false, true]],
    underline: (cell, ri, ci, range) => Math.random() > 0.5
});
```

Some styles take values that are more complex objects:
```js
cell.style("fill", {
    type: "pattern",
    pattern: "darkDown",
    foreground: {
        rgb: "ff0000"
    },
    background: {
        theme: 3,
        tint: 0.4
    }
});
```

There are often shortcuts for the setters, but the getters will always return the full objects:
```js
cell.style("fill", "0000ff");

const fill = cell.style("fill");
/*
fill is now set to:
{
    type: "solid",
    color: {
        rgb: "0000ff"
    }
}
*/
```

### Dates

Excel stores date/times as the number of days since 1/1/1900 ([sort of](https://en.wikipedia.org/wiki/Leap_year_bug)). It just applies a number formatting to make the number appear as a date. So to set a date value, you will need to also set a number format for a date if one doesn't already exist in the cell:
```js
cell.value(new Date()).style("numberFormat", "dddd, mmmm dd, yyyy");
```
When fetching the value of the cell, it will be returned as a number. To convert it to a date use [Workbook.numberToDate](#Workbook.numberToDate):
```js
const num = cell.value(); // 42788
const date = Workbook.numberToDate(num); // Wed Feb 22 2017 00:00:00 GMT-0500 (Eastern Standard Time)
```

### Serving from Express
You can serve the workbook from [express](http://expressjs.com/) or other web servers with something like this:
```js
router.get("/download", function (req, res, next) {
    // Open the workbook.
    Workbook.fromFileAsync("input.xlsx")
        .then(workbook => {
            // Make edits.
            workbook.sheet(0).cell("A1").value("foo");
            
            // Get the output
            return workbook.outputAsync();
        })
        .then(data => {
            // Set the output file name.
            res.attachment("output.xlsx");
            
            // Send the workbook.
            res.send(data);
        })
        .catch(next);
});
```

### Browser Usage
TODO


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

## Style Reference

### Styles
|Style Name|Type|Description|
| ------------- | ------------- | ----- |
|bold|`boolean`|`true` for bold, `false` for not bold|
|italic|`boolean`|`true` for italic, `false` for not italic|
|underline|`boolean|string`|`true` for single underline, `false` for no underline, `'double'` for double-underline|
|strikethrough|`boolean`|`true` for strikethrough `false` for not strikethrough|
|subscript|`boolean`|`true` for subscript, `false` for not subscript (cannot be combined with superscript)|
|superscript|`boolean`|`true` for superscript, `false` for not superscript (cannot be combined with subscript)|
|fontSize|`number`|Font size in points. Must be greater than 0.|
|fontFamily|`string`|Name of font family.|
|fontColor|`Color|string|number`|Color of the font. If string, will set an RGB color. If number, will set a theme color.|
|horizontalAlignment|`string`|Horizontal alignment. Allowed values: `'left'`, `'center'`, `'right'`, `'fill'`, `'justify'`, `'centerContinuous'`, `'distributed'`|
|justifyLastLine|`boolean`|a.k.a Justified Distributed. Only applies when horizontalAlignment === `'distributed'`) A boolean value indicating if the cells justified or distributed alignment should be used on the last line of text. (This is typical for East Asian alignments but not typical in other contexts.)|
|indent|`number`|Number of indents. Must be greater than or equal to 0.|
|verticalAlignment|`string`|Vertical alignment. Allowed values: `'top'`, `'center'`, `'bottom'`, `'justify'`, `'distributed'`|
|wrapText|`boolean`|`true` to wrap the text in the cell, `false` to not wrap.|
|shrinkToFit|`boolean`|`true` to shrink the text in the cell to fit, `false` to not shrink.|
|textDirection|`string`|Direction of the text. Allowed values: `'left-to-right'`, `'right-to-left'`|
|textRotation|`number`|Counter-clockwise angle of rotation in degrees. Must be [-90, 90] where negative numbers indicated clockwise rotation.|
|angleTextCounterclockwise|`boolean`|Shortcut for textRotation of 45 degrees.|
|angleTextClockwise|`boolean`|Shortcut for textRotation of -45 degrees.|
|rotateTextUp|`boolean`|Shortcut for textRotation of 90 degrees.|
|rotateTextDown|`boolean`|Shortcut for textRotation of -90 degrees.|
|verticalText|`boolean`|Special rotation that shows text vertical but individual letters are oriented normally. `true` to rotate, `false` to not rotate.|
|fill|`SolidFill|PatternFill|GradientFill|string|number`|The cell fill. If string, will set a solid RGB fill. If number, will set a solid theme color fill.|
|border|`Borders|Border|string|boolean}`|The border settings. If string, with set outside borders with to given border style. If true, will set outside border style to `'thin'`.|
|borderColor|`Color|string|number`|Color of the borders. If string, will set an RGB color. If number, will set a theme color.|
|borderStyle|`string`|Style of the outside borders. Allowed values: `'hair'`, `'dotted'`, `'dashDotDot'`, `'dashed'`, `'mediumDashDotDot'`, `'thin'`, `'slantDashDot'`, `'mediumDashDot'`, `'mediumDashed'`, `'medium'`, `'thick'`, `'double'`|
|leftBorder, rightBorder, topBorder, bottomBorder, diagonalBorder|`Border|string|boolean`|The border settings for the given side. If string, with set border to the given border style. If true, will set border style to `'thin'`.|
|leftBorderColor, rightBorderColor, topBorderColor, bottomBorderColor, diagonalBorderColor|`Color|string|number`|Color of the given border. If string, will set an RGB color. If number, will set a theme color.|
|leftBorderStyle, rightBorderStyle, topBorderStyle, bottomBorderStyle, diagonalBorderStyle|`string`|Style of the given side.|
|diagonalBorderDirection|`string`|Direction of the diagonal border(s) from left to right. Allowed values: `'up'`, `'down'`, `'both'`|
|numberFormat|`string`|Number format code. See TODO.|

### Color
|Property|Type|Description|
| ------------- | ------------- | ----- |
|[rgb]|`string`|RGB color code (e.g. `'ff0000'`). Either rgb or theme is required.|
|[theme]|`number`|Index of a theme color. Either rgb or theme is required.|
|[tint]|`number`|Optional tint value of the color from -1 to 1. Particularly useful for theme colors. 0.0 means no tint, -1.0 means 100% darken, and 1.0 means 100% lighten.|

### Borders
|Property|Type|Description|
| ------------- | ------------- | ----- |
|[left]|`Border|string|boolean`|The border settings for the left side. If string, with set border to the given border style. If true, will set border style to `'thin'`.|
|[right]|`Border|string|boolean`|The border settings for the right side. If string, with set border to the given border style. If true, will set border style to `'thin'`.|
|[top]|`Border|string|boolean`|The border settings for the top side. If string, with set border to the given border style. If true, will set border style to `'thin'`.|
|[bottom]|`Border|string|boolean`|The border settings for the bottom side. If string, with set border to the given border style. If true, will set border style to `'thin'`.|
|[diagonal]|`Border|string|boolean`|The border settings for the diagonal side. If string, with set border to the given border style. If true, will set border style to `'thin'`.|

### Border
|Property|Type|Description|
| ------------- | ------------- | ----- |
|style|`string`|Style of the given border.|
|color|`Color|string|number`|Color of the given border. If string, will set an RGB color. If number, will set a theme color.|
|[direction]|`string`|For diagonal border, the direction of the border(s) from left to right. Allowed values: `'up'`, `'down'`, `'both'`|

### SolidFill
|Property|Type|Description|
| ------------- | ------------- | ----- |
|type|`'solid'`||
|color|`Color|string|number`|Color of the fill. If string, will set an RGB color. If number, will set a theme color.|

### PatternFill
|Property|Type|Description|
| ------------- | ------------- | ----- |
|type|`'pattern'`||
|pattern|`string`|Name of the pattern. Allowed values: `'gray125'`, `'darkGray'`, `'mediumGray'`, `'lightGray'`, `'gray0625'`, `'darkHorizontal'`, `'darkVertical'`, `'darkDown'`, `'darkUp'`, `'darkGrid'`, `'darkTrellis'`, `'lightHorizontal'`, `'lightVertical'`, `'lightDown'`, `'lightUp'`, `'lightGrid'`, `'lightTrellis'`.|
|foreground|`Color|string|number`|Color of the foreground. If string, will set an RGB color. If number, will set a theme color.|
|background|`Color|string|number`|Color of the background. If string, will set an RGB color. If number, will set a theme color.|

### GradientFill
|Property|Type|Description|
| ------------- | ------------- | ----- |
|type|`'gradient'`||
|[gradientType]|`string`|Type of gradient. Allowed values: `'linear'` (default), `'path'`. With a path gradient, a path is drawn between the top, left, right, and bottom values and a graident is draw from that path to the outside of the cell.|
|stops|`Array.<{}>`||
|stops[].position|`number`|The position of the stop from 0 to 1.|
|stops[].color|`Color|string|number`|Color of the stop. If string, will set an RGB color. If number, will set a theme color.|
|[angle]|`number`|If linear gradient, the angle of clockwise rotation of the gradient.|
|[left]|`number`|If path gradient, the left position of the path as a percentage from 0 to 1.|
|[right]|`number`|If path gradient, the right position of the path as a percentage from 0 to 1.|
|[top]|`number`|If path gradient, the top position of the path as a percentage from 0 to 1.|
|[bottom]|`number`|If path gradient, the bottom position of the path as a percentage from 0 to 1.|

## API Reference
### Classes

<dl>
<dt><a href="#Cell">Cell</a></dt>
<dd><p>A cell</p>
</dd>
<dt><a href="#Column">Column</a></dt>
<dd><p>A column.</p>
</dd>
<dt><a href="#Range">Range</a></dt>
<dd><p>A range of cells.</p>
</dd>
<dt><a href="#Row">Row</a></dt>
<dd><p>A row.</p>
</dd>
<dt><a href="#Sheet">Sheet</a></dt>
<dd><p>A worksheet.</p>
</dd>
<dt><a href="#Workbook">Workbook</a></dt>
<dd><p>A workbook.</p>
</dd>
</dl>

### Objects

<dl>
<dt><a href="#XlsxPopulate">XlsxPopulate</a> : <code>object</code></dt>
<dd></dd>
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
        * [.find(pattern, [replacement])](#Cell+find) ⇒ <code>boolean</code>
        * [.formula()](#Cell+formula) ⇒ <code>string</code>
        * [.formula(formula)](#Cell+formula) ⇒ <code>[Cell](#Cell)</code>
        * [.tap(callback)](#Cell+tap) ⇒ <code>[Cell](#Cell)</code>
        * [.thru(callback)](#Cell+thru) ⇒ <code>\*</code>
        * [.rangeTo(cell)](#Cell+rangeTo) ⇒ <code>[Range](#Range)</code>
        * [.relativeCell(rowOffset, columnOffset)](#Cell+relativeCell) ⇒ <code>[Cell](#Cell)</code>
        * [.row()](#Cell+row) ⇒ <code>[Row](#Row)</code>
        * [.rowNumber()](#Cell+rowNumber) ⇒ <code>number</code>
        * [.sheet()](#Cell+sheet) ⇒ <code>[Sheet](#Sheet)</code>
        * [.style(name)](#Cell+style) ⇒ <code>\*</code>
        * [.style(names)](#Cell+style) ⇒ <code>object.&lt;string, \*&gt;</code>
        * [.style(name, value)](#Cell+style) ⇒ <code>[Cell](#Cell)</code>
        * [.style(styles)](#Cell+style) ⇒ <code>[Cell](#Cell)</code>
        * [.value()](#Cell+value) ⇒ <code>string</code> &#124; <code>boolean</code> &#124; <code>number</code> &#124; <code>Date</code> &#124; <code>undefined</code>
        * [.value(value)](#Cell+value) ⇒ <code>[Cell](#Cell)</code>
        * [.workbook()](#Cell+workbook) ⇒ <code>[XlsxPopulate](#XlsxPopulate)</code>
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
<a name="Cell+find"></a>

#### cell.find(pattern, [replacement]) ⇒ <code>boolean</code>
Find the given pattern in the cell and optionally replace it.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>boolean</code> - A flag indicating if the pattern was found.  

| Param | Type | Description |
| --- | --- | --- |
| pattern | <code>string</code> &#124; <code>RegExp</code> | The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches. |
| [replacement] | <code>string</code> &#124; <code>function</code> | The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in the cell will be replaced. |

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
**Returns**: <code>\*</code> - The return value of the callback.  

| Param | Type | Description |
| --- | --- | --- |
| callback | <code>[thruCallback](#Cell..thruCallback)</code> | The callback function. |

<a name="Cell+rangeTo"></a>

#### cell.rangeTo(cell) ⇒ <code>[Range](#Range)</code>
Create a range from this cell and another.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| cell | <code>[Cell](#Cell)</code> &#124; <code>string</code> | The other cell or cell address to range to. |

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

#### cell.sheet() ⇒ <code>[Sheet](#Sheet)</code>
Gets the parent sheet.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[Sheet](#Sheet)</code> - The parent sheet.  
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

#### cell.workbook() ⇒ <code>[XlsxPopulate](#XlsxPopulate)</code>
Gets the parent workbook.

**Kind**: instance method of <code>[Cell](#Cell)</code>  
**Returns**: <code>[XlsxPopulate](#XlsxPopulate)</code> - The parent workbook.  
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
    * [.hidden()](#Column+hidden) ⇒ <code>boolean</code>
    * [.hidden(hidden)](#Column+hidden) ⇒ <code>[Column](#Column)</code>
    * [.sheet()](#Column+sheet) ⇒ <code>[Sheet](#Sheet)</code>
    * [.width()](#Column+width) ⇒ <code>undefined</code> &#124; <code>number</code>
    * [.width(width)](#Column+width) ⇒ <code>[Column](#Column)</code>
    * [.workbook()](#Column+workbook) ⇒ <code>[XlsxPopulate](#XlsxPopulate)</code>

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

#### column.hidden() ⇒ <code>boolean</code>
Gets a value indicating whether the column is hidden.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>boolean</code> - A flag indicating whether the column is hidden.  
<a name="Column+hidden"></a>

#### column.hidden(hidden) ⇒ <code>[Column](#Column)</code>
Sets whether the column is hidden.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>[Column](#Column)</code> - The column.  

| Param | Type | Description |
| --- | --- | --- |
| hidden | <code>boolean</code> | A flag indicating whether to hide the column. |

<a name="Column+sheet"></a>

#### column.sheet() ⇒ <code>[Sheet](#Sheet)</code>
Get the parent sheet.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>[Sheet](#Sheet)</code> - The parent sheet.  
<a name="Column+width"></a>

#### column.width() ⇒ <code>undefined</code> &#124; <code>number</code>
Gets the width.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>undefined</code> &#124; <code>number</code> - The width (or undefined).  
<a name="Column+width"></a>

#### column.width(width) ⇒ <code>[Column](#Column)</code>
Sets the width.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>[Column](#Column)</code> - The column.  

| Param | Type | Description |
| --- | --- | --- |
| width | <code>number</code> | The width of the column. |

<a name="Column+workbook"></a>

#### column.workbook() ⇒ <code>[XlsxPopulate](#XlsxPopulate)</code>
Get the parent workbook.

**Kind**: instance method of <code>[Column](#Column)</code>  
**Returns**: <code>[XlsxPopulate](#XlsxPopulate)</code> - The parent workbook.  
<a name="Range"></a>

### Range
A range of cells.

**Kind**: global class  

* [Range](#Range)
    * _instance_
        * [.address([opts])](#Range+address) ⇒ <code>string</code>
        * [.cell(ri, ci)](#Range+cell) ⇒ <code>[Cell](#Cell)</code>
        * [.clear()](#Range+clear) ⇒ <code>[Range](#Range)</code>
        * [.endCell()](#Range+endCell) ⇒ <code>[Cell](#Cell)</code>
        * [.forEach(callback)](#Range+forEach) ⇒ <code>[Range](#Range)</code>
        * [.formula()](#Range+formula) ⇒ <code>string</code> &#124; <code>undefined</code>
        * [.formula(formula)](#Range+formula) ⇒ <code>[Range](#Range)</code>
        * [.map(callback)](#Range+map) ⇒ <code>Array.&lt;Array.&lt;\*&gt;&gt;</code>
        * [.merged()](#Range+merged) ⇒ <code>boolean</code>
        * [.merged(merged)](#Range+merged) ⇒ <code>[Range](#Range)</code>
        * [.reduce(callback, [initialValue])](#Range+reduce) ⇒ <code>\*</code>
        * [.sheet()](#Range+sheet) ⇒ <code>[Sheet](#Sheet)</code>
        * [.startCell()](#Range+startCell) ⇒ <code>[Cell](#Cell)</code>
        * [.style(name)](#Range+style) ⇒ <code>Array.&lt;Array.&lt;\*&gt;&gt;</code>
        * [.style(names)](#Range+style) ⇒ <code>Object.&lt;string, Array.&lt;Array.&lt;\*&gt;&gt;&gt;</code>
        * [.style(name)](#Range+style) ⇒ <code>[Range](#Range)</code>
        * [.style(name)](#Range+style) ⇒ <code>[Range](#Range)</code>
        * [.style(name, value)](#Range+style) ⇒ <code>[Range](#Range)</code>
        * [.style(styles)](#Range+style) ⇒ <code>[Range](#Range)</code>
        * [.tap(callback)](#Range+tap) ⇒ <code>[Range](#Range)</code>
        * [.thru(callback)](#Range+thru) ⇒ <code>\*</code>
        * [.value()](#Range+value) ⇒ <code>Array.&lt;Array.&lt;\*&gt;&gt;</code>
        * [.value()](#Range+value) ⇒ <code>[Range](#Range)</code>
        * [.value()](#Range+value) ⇒ <code>[Range](#Range)</code>
        * [.value(value)](#Range+value) ⇒ <code>[Range](#Range)</code>
        * [.workbook()](#Range+workbook) ⇒ <code>[XlsxPopulate](#XlsxPopulate)</code>
    * _inner_
        * [~forEachCallback](#Range..forEachCallback) ⇒ <code>undefined</code>
        * [~mapCallback](#Range..mapCallback) ⇒ <code>\*</code>
        * [~reduceCallback](#Range..reduceCallback) ⇒ <code>\*</code>
        * [~tapCallback](#Range..tapCallback) ⇒ <code>undefined</code>
        * [~thruCallback](#Range..thruCallback) ⇒ <code>\*</code>

<a name="Range+address"></a>

#### range.address([opts]) ⇒ <code>string</code>
Get the address of the range.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>string</code> - The address.  

| Param | Type | Description |
| --- | --- | --- |
| [opts] | <code>Object</code> | Options |
| [opts.includeSheetName] | <code>boolean</code> | Include the sheet name in the address. |
| [opts.startRowAnchored] | <code>boolean</code> | Anchor the start row. |
| [opts.startColumnAnchored] | <code>boolean</code> | Anchor the start column. |
| [opts.endRowAnchored] | <code>boolean</code> | Anchor the end row. |
| [opts.endColumnAnchored] | <code>boolean</code> | Anchor the end column. |

<a name="Range+cell"></a>

#### range.cell(ri, ci) ⇒ <code>[Cell](#Cell)</code>
Gets a cell within the range.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| ri | <code>number</code> | Row index relative to the top-left corner of the range (0-based). |
| ci | <code>number</code> | Column index relative to the top-left corner of the range (0-based). |

<a name="Range+clear"></a>

#### range.clear() ⇒ <code>[Range](#Range)</code>
Clear the contents of all the cells in the range.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  
<a name="Range+endCell"></a>

#### range.endCell() ⇒ <code>[Cell](#Cell)</code>
Get the end cell of the range.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The end cell.  
<a name="Range+forEach"></a>

#### range.forEach(callback) ⇒ <code>[Range](#Range)</code>
Call a function for each cell in the range. Goes by row then column.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| callback | <code>[forEachCallback](#Range..forEachCallback)</code> | Function called for each cell in the range. |

<a name="Range+formula"></a>

#### range.formula() ⇒ <code>string</code> &#124; <code>undefined</code>
Gets the formula shared across the range.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>string</code> &#124; <code>undefined</code> - The shared formula.  
<a name="Range+formula"></a>

#### range.formula(formula) ⇒ <code>[Range](#Range)</code>
Sets the shared formula in the range. The formula will be translated for each cell.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| formula | <code>string</code> | The formula to set. |

<a name="Range+map"></a>

#### range.map(callback) ⇒ <code>Array.&lt;Array.&lt;\*&gt;&gt;</code>
Creates a 2D array of values by running each cell through a callback.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>Array.&lt;Array.&lt;\*&gt;&gt;</code> - The 2D array of return values.  

| Param | Type | Description |
| --- | --- | --- |
| callback | <code>[mapCallback](#Range..mapCallback)</code> | Function called for each cell in the range. |

<a name="Range+merged"></a>

#### range.merged() ⇒ <code>boolean</code>
Gets a value indicating whether the cells in the range are merged.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>boolean</code> - The value.  
<a name="Range+merged"></a>

#### range.merged(merged) ⇒ <code>[Range](#Range)</code>
Sets a value indicating whether the cells in the range should be merged.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| merged | <code>boolean</code> | True to merge, false to unmerge. |

<a name="Range+reduce"></a>

#### range.reduce(callback, [initialValue]) ⇒ <code>\*</code>
Reduces the range to a single value accumulated from the result of a function called for each cell.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>\*</code> - The accumulated value.  

| Param | Type | Description |
| --- | --- | --- |
| callback | <code>[reduceCallback](#Range..reduceCallback)</code> | Function called for each cell in the range. |
| [initialValue] | <code>\*</code> | The initial value. |

<a name="Range+sheet"></a>

#### range.sheet() ⇒ <code>[Sheet](#Sheet)</code>
Gets the parent sheet of the range.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Sheet](#Sheet)</code> - The parent sheet.  
<a name="Range+startCell"></a>

#### range.startCell() ⇒ <code>[Cell](#Cell)</code>
Gets the start cell of the range.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The start cell.  
<a name="Range+style"></a>

#### range.style(name) ⇒ <code>Array.&lt;Array.&lt;\*&gt;&gt;</code>
Gets a single style for each cell.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>Array.&lt;Array.&lt;\*&gt;&gt;</code> - 2D array of style values.  

| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | The name of the style. |

<a name="Range+style"></a>

#### range.style(names) ⇒ <code>Object.&lt;string, Array.&lt;Array.&lt;\*&gt;&gt;&gt;</code>
Gets multiple styles for each cell.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>Object.&lt;string, Array.&lt;Array.&lt;\*&gt;&gt;&gt;</code> - Object whose keys are style names and values are 2D arrays of style values.  

| Param | Type | Description |
| --- | --- | --- |
| names | <code>Array.&lt;string&gt;</code> | The names of the styles. |

<a name="Range+style"></a>

#### range.style(name) ⇒ <code>[Range](#Range)</code>
Set the style in each cell to the result of a function called for each.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | The name of the style. |
|  | <code>[mapCallback](#Range..mapCallback)</code> | The callback to provide value for the cell. |

<a name="Range+style"></a>

#### range.style(name) ⇒ <code>[Range](#Range)</code>
Sets the style in each cell to the corresponding value in the given 2D array of values.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | The name of the style. |
|  | <code>Array.&lt;Array.&lt;\*&gt;&gt;</code> | The style values to set. |

<a name="Range+style"></a>

#### range.style(name, value) ⇒ <code>[Range](#Range)</code>
Set the style of all cells in the range to a single style value.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | The name of the style. |
| value | <code>\*</code> | The value to set. |

<a name="Range+style"></a>

#### range.style(styles) ⇒ <code>[Range](#Range)</code>
Set multiple styles for the cells in the range.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| styles | <code>object.&lt;string, (Range~mapCallback\|Array.&lt;Array.&lt;\*&gt;&gt;\|\*)&gt;</code> | Object whose keys are style names and values are either function callbacks, 2D arrays of style values, or a single value for all the cells. |

<a name="Range+tap"></a>

#### range.tap(callback) ⇒ <code>[Range](#Range)</code>
Invoke a callback on the range and return the range. Useful for method chaining.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| callback | <code>[tapCallback](#Range..tapCallback)</code> | The callback function. |

<a name="Range+thru"></a>

#### range.thru(callback) ⇒ <code>\*</code>
Invoke a callback on the range and return the value provided by the callback. Useful for method chaining.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>\*</code> - The return value of the callback.  

| Param | Type | Description |
| --- | --- | --- |
| callback | <code>[thruCallback](#Range..thruCallback)</code> | The callback function. |

<a name="Range+value"></a>

#### range.value() ⇒ <code>Array.&lt;Array.&lt;\*&gt;&gt;</code>
Get the values of each cell in the range as a 2D array.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>Array.&lt;Array.&lt;\*&gt;&gt;</code> - The values.  
<a name="Range+value"></a>

#### range.value() ⇒ <code>[Range](#Range)</code>
Set the values in each cell to the result of a function called for each.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
|  | <code>[mapCallback](#Range..mapCallback)</code> | The callback to provide value for the cell. |

<a name="Range+value"></a>

#### range.value() ⇒ <code>[Range](#Range)</code>
Sets the value in each cell to the corresponding value in the given 2D array of values.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
|  | <code>Array.&lt;Array.&lt;\*&gt;&gt;</code> | The values to set. |

<a name="Range+value"></a>

#### range.value(value) ⇒ <code>[Range](#Range)</code>
Set the value of all cells in the range to a single value.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| value | <code>\*</code> | The value to set. |

<a name="Range+workbook"></a>

#### range.workbook() ⇒ <code>[XlsxPopulate](#XlsxPopulate)</code>
Gets the parent workbook.

**Kind**: instance method of <code>[Range](#Range)</code>  
**Returns**: <code>[XlsxPopulate](#XlsxPopulate)</code> - The parent workbook.  
<a name="Range..forEachCallback"></a>

#### Range~forEachCallback ⇒ <code>undefined</code>
Callback used by forEach.

**Kind**: inner typedef of <code>[Range](#Range)</code>  

| Param | Type | Description |
| --- | --- | --- |
| cell | <code>[Cell](#Cell)</code> | The cell. |
| ri | <code>number</code> | The relative row index. |
| ci | <code>number</code> | The relative column index. |
| range | <code>[Range](#Range)</code> | The range. |

<a name="Range..mapCallback"></a>

#### Range~mapCallback ⇒ <code>\*</code>
Callback used by map.

**Kind**: inner typedef of <code>[Range](#Range)</code>  
**Returns**: <code>\*</code> - The value to map to.  

| Param | Type | Description |
| --- | --- | --- |
| cell | <code>[Cell](#Cell)</code> | The cell. |
| ri | <code>number</code> | The relative row index. |
| ci | <code>number</code> | The relative column index. |
| range | <code>[Range](#Range)</code> | The range. |

<a name="Range..reduceCallback"></a>

#### Range~reduceCallback ⇒ <code>\*</code>
Callback used by reduce.

**Kind**: inner typedef of <code>[Range](#Range)</code>  
**Returns**: <code>\*</code> - The value to map to.  

| Param | Type | Description |
| --- | --- | --- |
| accumulator | <code>\*</code> | The accumulated value. |
| cell | <code>[Cell](#Cell)</code> | The cell. |
| ri | <code>number</code> | The relative row index. |
| ci | <code>number</code> | The relative column index. |
| range | <code>[Range](#Range)</code> | The range. |

<a name="Range..tapCallback"></a>

#### Range~tapCallback ⇒ <code>undefined</code>
Callback used by tap.

**Kind**: inner typedef of <code>[Range](#Range)</code>  

| Param | Type | Description |
| --- | --- | --- |
| range | <code>[Range](#Range)</code> | The range. |

<a name="Range..thruCallback"></a>

#### Range~thruCallback ⇒ <code>\*</code>
Callback used by thru.

**Kind**: inner typedef of <code>[Range](#Range)</code>  
**Returns**: <code>\*</code> - The value to return from thru.  

| Param | Type | Description |
| --- | --- | --- |
| range | <code>[Range](#Range)</code> | The range. |

<a name="Row"></a>

### Row
A row.

**Kind**: global class  

* [Row](#Row)
    * [.address([opts])](#Row+address) ⇒ <code>string</code>
    * [.cell(columnNameOrNumber)](#Row+cell) ⇒ <code>[Cell](#Cell)</code>
    * [.height()](#Row+height) ⇒ <code>undefined</code> &#124; <code>number</code>
    * [.height(height)](#Row+height) ⇒ <code>[Row](#Row)</code>
    * [.hidden()](#Row+hidden) ⇒ <code>boolean</code>
    * [.hidden(hidden)](#Row+hidden) ⇒ <code>[Row](#Row)</code>
    * [.rowNumber()](#Row+rowNumber) ⇒ <code>number</code>
    * [.sheet()](#Row+sheet) ⇒ <code>[Sheet](#Sheet)</code>
    * [.workbook()](#Row+workbook) ⇒ <code>[XlsxPopulate](#XlsxPopulate)</code>

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

#### row.height() ⇒ <code>undefined</code> &#124; <code>number</code>
Gets the row height.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>undefined</code> &#124; <code>number</code> - The height (or undefined).  
<a name="Row+height"></a>

#### row.height(height) ⇒ <code>[Row](#Row)</code>
Sets the row height.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>[Row](#Row)</code> - The row.  

| Param | Type | Description |
| --- | --- | --- |
| height | <code>number</code> | The height of the row. |

<a name="Row+hidden"></a>

#### row.hidden() ⇒ <code>boolean</code>
Gets a value indicating whether the row is hidden.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>boolean</code> - A flag indicating whether the row is hidden.  
<a name="Row+hidden"></a>

#### row.hidden(hidden) ⇒ <code>[Row](#Row)</code>
Sets whether the row is hidden.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>[Row](#Row)</code> - The row.  

| Param | Type | Description |
| --- | --- | --- |
| hidden | <code>boolean</code> | A flag indicating whether to hide the row. |

<a name="Row+rowNumber"></a>

#### row.rowNumber() ⇒ <code>number</code>
Gets the row number.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>number</code> - The row number.  
<a name="Row+sheet"></a>

#### row.sheet() ⇒ <code>[Sheet](#Sheet)</code>
Gets the parent sheet of the row.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>[Sheet](#Sheet)</code> - The parent sheet.  
<a name="Row+workbook"></a>

#### row.workbook() ⇒ <code>[XlsxPopulate](#XlsxPopulate)</code>
Get the parent workbook.

**Kind**: instance method of <code>[Row](#Row)</code>  
**Returns**: <code>[XlsxPopulate](#XlsxPopulate)</code> - The parent workbook.  
<a name="Sheet"></a>

### Sheet
A worksheet.

**Kind**: global class  

* [Sheet](#Sheet)
    * [.cell(address)](#Sheet+cell) ⇒ <code>[Cell](#Cell)</code>
    * [.cell(rowNumber, columnNameOrNumber)](#Sheet+cell) ⇒ <code>[Cell](#Cell)</code>
    * [.column(columnNameOrNumber)](#Sheet+column) ⇒ <code>[Column](#Column)</code>
    * [.definedName(name)](#Sheet+definedName) ⇒ <code>undefined</code> &#124; <code>[Cell](#Cell)</code> &#124; <code>[Range](#Range)</code> &#124; <code>[Row](#Row)</code> &#124; <code>[Column](#Column)</code>
    * [.find(pattern, [replacement])](#Sheet+find) ⇒ <code>[Array.&lt;Cell&gt;](#Cell)</code>
    * [.name()](#Sheet+name) ⇒ <code>string</code>
    * [.range(address)](#Sheet+range) ⇒ <code>[Range](#Range)</code>
    * [.range(startCell, endCell)](#Sheet+range) ⇒ <code>[Range](#Range)</code>
    * [.range(startRowNumber, startColumnNameOrNumber, endRowNumber, endColumnNameOrNumber)](#Sheet+range) ⇒ <code>[Range](#Range)</code>
    * [.row(rowNumber)](#Sheet+row) ⇒ <code>[Row](#Row)</code>
    * [.usedRange()](#Sheet+usedRange) ⇒ <code>[Range](#Range)</code> &#124; <code>undefined</code>
    * [.workbook()](#Sheet+workbook) ⇒ <code>[XlsxPopulate](#XlsxPopulate)</code>

<a name="Sheet+cell"></a>

#### sheet.cell(address) ⇒ <code>[Cell](#Cell)</code>
Gets the cell with the given address.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| address | <code>string</code> | The address of the cell. |

<a name="Sheet+cell"></a>

#### sheet.cell(rowNumber, columnNameOrNumber) ⇒ <code>[Cell](#Cell)</code>
Gets the cell with the given row and column numbers.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Cell](#Cell)</code> - The cell.  

| Param | Type | Description |
| --- | --- | --- |
| rowNumber | <code>number</code> | The row number of the cell. |
| columnNameOrNumber | <code>string</code> &#124; <code>number</code> | The column name or number of the cell. |

<a name="Sheet+column"></a>

#### sheet.column(columnNameOrNumber) ⇒ <code>[Column](#Column)</code>
Gets a column in the sheet.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Column](#Column)</code> - The column.  

| Param | Type | Description |
| --- | --- | --- |
| columnNameOrNumber | <code>string</code> &#124; <code>number</code> | The name or number of the column. |

<a name="Sheet+definedName"></a>

#### sheet.definedName(name) ⇒ <code>undefined</code> &#124; <code>[Cell](#Cell)</code> &#124; <code>[Range](#Range)</code> &#124; <code>[Row](#Row)</code> &#124; <code>[Column](#Column)</code>
Gets a defined name scoped to the sheet.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>undefined</code> &#124; <code>[Cell](#Cell)</code> &#124; <code>[Range](#Range)</code> &#124; <code>[Row](#Row)</code> &#124; <code>[Column](#Column)</code> - The named selection or undefined if name not found.  
**Throws**:

- <code>Error</code> Will throw if address in defined name is not supported.


| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | The defined name. |

<a name="Sheet+find"></a>

#### sheet.find(pattern, [replacement]) ⇒ <code>[Array.&lt;Cell&gt;](#Cell)</code>
Find the given pattern in the sheet and optionally replace it.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Array.&lt;Cell&gt;](#Cell)</code> - The matching cells.  

| Param | Type | Description |
| --- | --- | --- |
| pattern | <code>string</code> &#124; <code>RegExp</code> | The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches. |
| [replacement] | <code>string</code> &#124; <code>function</code> | The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in each cell will be replaced. |

<a name="Sheet+name"></a>

#### sheet.name() ⇒ <code>string</code>
Get the name of the sheet.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>string</code> - The sheet name.  
<a name="Sheet+range"></a>

#### sheet.range(address) ⇒ <code>[Range](#Range)</code>
Gets a range from the given range address.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| address | <code>string</code> | The range address (e.g. 'A1:B3'). |

<a name="Sheet+range"></a>

#### sheet.range(startCell, endCell) ⇒ <code>[Range](#Range)</code>
Gets a range from the given cells or cell addresses.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| startCell | <code>string</code> &#124; <code>[Cell](#Cell)</code> | The starting cell or cell address (e.g. 'A1'). |
| endCell | <code>string</code> &#124; <code>[Cell](#Cell)</code> | The ending cell or cell address (e.g. 'B3'). |

<a name="Sheet+range"></a>

#### sheet.range(startRowNumber, startColumnNameOrNumber, endRowNumber, endColumnNameOrNumber) ⇒ <code>[Range](#Range)</code>
Gets a range from the given row numbers and column names or numbers.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Range](#Range)</code> - The range.  

| Param | Type | Description |
| --- | --- | --- |
| startRowNumber | <code>number</code> | The starting cell row number. |
| startColumnNameOrNumber | <code>string</code> &#124; <code>number</code> | The starting cell column name or number. |
| endRowNumber | <code>number</code> | The ending cell row number. |
| endColumnNameOrNumber | <code>string</code> &#124; <code>number</code> | The ending cell column name or number. |

<a name="Sheet+row"></a>

#### sheet.row(rowNumber) ⇒ <code>[Row](#Row)</code>
Gets the row with the given number.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Row](#Row)</code> - The row with the given number.  

| Param | Type | Description |
| --- | --- | --- |
| rowNumber | <code>number</code> | The row number. |

<a name="Sheet+usedRange"></a>

#### sheet.usedRange() ⇒ <code>[Range](#Range)</code> &#124; <code>undefined</code>
Get the range of cells in the sheet that have contained a value or style at any point. Useful for extracting the entire sheet contents.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[Range](#Range)</code> &#124; <code>undefined</code> - The used range or undefined if no cells in the sheet are used.  
<a name="Sheet+workbook"></a>

#### sheet.workbook() ⇒ <code>[XlsxPopulate](#XlsxPopulate)</code>
Gets the parent workbook.

**Kind**: instance method of <code>[Sheet](#Sheet)</code>  
**Returns**: <code>[XlsxPopulate](#XlsxPopulate)</code> - The parent workbook.  
<a name="Workbook"></a>

### Workbook
A workbook.

**Kind**: global class  

* [Workbook](#Workbook)
    * [.definedName(name)](#Workbook+definedName) ⇒ <code>undefined</code> &#124; <code>[Cell](#Cell)</code> &#124; <code>[Range](#Range)</code> &#124; <code>[Row](#Row)</code> &#124; <code>[Column](#Column)</code>
    * [.find(pattern, [replacement])](#Workbook+find) ⇒ <code>boolean</code>
    * [.outputAsync([type])](#Workbook+outputAsync) ⇒ <code>string</code> &#124; <code>Uint8Array</code> &#124; <code>ArrayBuffer</code> &#124; <code>Blob</code> &#124; <code>Buffer</code>
    * [.sheet(sheetNameOrIndex)](#Workbook+sheet) ⇒ <code>[Sheet](#Sheet)</code> &#124; <code>undefined</code>
    * [.toFileAsync(path)](#Workbook+toFileAsync) ⇒ <code>Promise.&lt;undefined&gt;</code>

<a name="Workbook+definedName"></a>

#### workbook.definedName(name) ⇒ <code>undefined</code> &#124; <code>[Cell](#Cell)</code> &#124; <code>[Range](#Range)</code> &#124; <code>[Row](#Row)</code> &#124; <code>[Column](#Column)</code>
Gets a defined name scoped to the workbook.

**Kind**: instance method of <code>[Workbook](#Workbook)</code>  
**Returns**: <code>undefined</code> &#124; <code>[Cell](#Cell)</code> &#124; <code>[Range](#Range)</code> &#124; <code>[Row](#Row)</code> &#124; <code>[Column](#Column)</code> - The named selection or undefined if name not found.  
**Throws**:

- <code>Error</code> Will throw if address in defined name is not supported.


| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | The defined name. |

<a name="Workbook+find"></a>

#### workbook.find(pattern, [replacement]) ⇒ <code>boolean</code>
Find the given pattern in the workbook and optionally replace it.

**Kind**: instance method of <code>[Workbook](#Workbook)</code>  
**Returns**: <code>boolean</code> - A flag indicating if the pattern was found.  

| Param | Type | Description |
| --- | --- | --- |
| pattern | <code>string</code> &#124; <code>RegExp</code> | The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches. |
| [replacement] | <code>string</code> &#124; <code>function</code> | The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in each cell will be replaced. |

<a name="Workbook+outputAsync"></a>

#### workbook.outputAsync([type]) ⇒ <code>string</code> &#124; <code>Uint8Array</code> &#124; <code>ArrayBuffer</code> &#124; <code>Blob</code> &#124; <code>Buffer</code>
Generates the workbook output.

**Kind**: instance method of <code>[Workbook](#Workbook)</code>  
**Returns**: <code>string</code> &#124; <code>Uint8Array</code> &#124; <code>ArrayBuffer</code> &#124; <code>Blob</code> &#124; <code>Buffer</code> - The data.  

| Param | Type | Description |
| --- | --- | --- |
| [type] | <code>string</code> | The type of the data to return. (Supports any supported [JSZip data types](https://stuk.github.io/jszip/documentation/api_jszip/generate_async.html): base64, binarystring, uint8array, arraybuffer, blob, nodebuffer) Defaults to 'nodebuffer' in Node.js and 'blob' in browsers. |

<a name="Workbook+sheet"></a>

#### workbook.sheet(sheetNameOrIndex) ⇒ <code>[Sheet](#Sheet)</code> &#124; <code>undefined</code>
Gets the sheet with the provided name or index (0-based).

**Kind**: instance method of <code>[Workbook](#Workbook)</code>  
**Returns**: <code>[Sheet](#Sheet)</code> &#124; <code>undefined</code> - The sheet or undefined if not found.  

| Param | Type | Description |
| --- | --- | --- |
| sheetNameOrIndex | <code>string</code> &#124; <code>number</code> | The sheet name or index. |

<a name="Workbook+toFileAsync"></a>

#### workbook.toFileAsync(path) ⇒ <code>Promise.&lt;undefined&gt;</code>
Write the workbook to file. (Not supported in browsers.)

**Kind**: instance method of <code>[Workbook](#Workbook)</code>  
**Returns**: <code>Promise.&lt;undefined&gt;</code> - A promise.  

| Param | Type | Description |
| --- | --- | --- |
| path | <code>string</code> | The path of the file to write. |

<a name="XlsxPopulate"></a>

### XlsxPopulate : <code>object</code>
**Kind**: global namespace  

* [XlsxPopulate](#XlsxPopulate) : <code>object</code>
    * [.MIME_TYPE](#XlsxPopulate.MIME_TYPE) : <code>string</code>
    * [.dateToNumber(date)](#XlsxPopulate.dateToNumber) ⇒ <code>number</code>
    * [.fromBlankAsync()](#XlsxPopulate.fromBlankAsync) ⇒ <code>[Promise.&lt;Workbook&gt;](#Workbook)</code>
    * [.fromDataAsync(data)](#XlsxPopulate.fromDataAsync) ⇒ <code>[Promise.&lt;Workbook&gt;](#Workbook)</code>
    * [.fromFileAsync(path)](#XlsxPopulate.fromFileAsync) ⇒ <code>[Promise.&lt;Workbook&gt;](#Workbook)</code>
    * [.numberToDate(number)](#XlsxPopulate.numberToDate) ⇒ <code>Date</code>

<a name="XlsxPopulate.MIME_TYPE"></a>

#### XlsxPopulate.MIME_TYPE : <code>string</code>
The XLSX mime type.

**Kind**: static property of <code>[XlsxPopulate](#XlsxPopulate)</code>  
<a name="XlsxPopulate.dateToNumber"></a>

#### XlsxPopulate.dateToNumber(date) ⇒ <code>number</code>
Convert a date to a number for Excel.

**Kind**: static method of <code>[XlsxPopulate](#XlsxPopulate)</code>  
**Returns**: <code>number</code> - The number.  

| Param | Type | Description |
| --- | --- | --- |
| date | <code>Date</code> | The date. |

<a name="XlsxPopulate.fromBlankAsync"></a>

#### XlsxPopulate.fromBlankAsync() ⇒ <code>[Promise.&lt;Workbook&gt;](#Workbook)</code>
Create a new blank workbook.

**Kind**: static method of <code>[XlsxPopulate](#XlsxPopulate)</code>  
**Returns**: <code>[Promise.&lt;Workbook&gt;](#Workbook)</code> - The workbook.  
<a name="XlsxPopulate.fromDataAsync"></a>

#### XlsxPopulate.fromDataAsync(data) ⇒ <code>[Promise.&lt;Workbook&gt;](#Workbook)</code>
Loads a workbook from a data object. (Supports any supported [JSZip data types](https://stuk.github.io/jszip/documentation/api_jszip/load_async.html).)

**Kind**: static method of <code>[XlsxPopulate](#XlsxPopulate)</code>  
**Returns**: <code>[Promise.&lt;Workbook&gt;](#Workbook)</code> - The workbook.  

| Param | Type | Description |
| --- | --- | --- |
| data | <code>string</code> &#124; <code>Array.&lt;number&gt;</code> &#124; <code>ArrayBuffer</code> &#124; <code>Uint8Array</code> &#124; <code>Buffer</code> &#124; <code>Blob</code> &#124; <code>Promise.&lt;\*&gt;</code> | The data to load. |

<a name="XlsxPopulate.fromFileAsync"></a>

#### XlsxPopulate.fromFileAsync(path) ⇒ <code>[Promise.&lt;Workbook&gt;](#Workbook)</code>
Loads a workbook from file.

**Kind**: static method of <code>[XlsxPopulate](#XlsxPopulate)</code>  
**Returns**: <code>[Promise.&lt;Workbook&gt;](#Workbook)</code> - The workbook.  

| Param | Type | Description |
| --- | --- | --- |
| path | <code>string</code> | The path to the workbook. |

<a name="XlsxPopulate.numberToDate"></a>

#### XlsxPopulate.numberToDate(number) ⇒ <code>Date</code>
Convert an Excel number to a date.

**Kind**: static method of <code>[XlsxPopulate](#XlsxPopulate)</code>  
**Returns**: <code>Date</code> - The date.  

| Param | Type | Description |
| --- | --- | --- |
| number | <code>number</code> | The number. |


