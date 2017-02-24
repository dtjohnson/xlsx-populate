[![view on npm](http://img.shields.io/npm/v/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![npm module downloads per month](http://img.shields.io/npm/dm/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![Build Status](https://travis-ci.org/dtjohnson/xlsx-populate.svg?branch=master)](https://travis-ci.org/dtjohnson/xlsx-populate)
[![Dependency Status](https://david-dm.org/dtjohnson/xlsx-populate.svg)](https://david-dm.org/dtjohnson/xlsx-populate)

# xlsx-populate
Excel XLSX parser/generator written in JavaScript with Node.js and browser support, jQuery/d3-style method chaining, and a focus on keeping existing workbook features and styles in tact.

## Table of Contents
<!-- toc -->

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

### NOTOC-Styles
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

### NOTOC-Color
|Property|Type|Description|
| ------------- | ------------- | ----- |
|[rgb]|`string`|RGB color code (e.g. `'ff0000'`). Either rgb or theme is required.|
|[theme]|`number`|Index of a theme color. Either rgb or theme is required.|
|[tint]|`number`|Optional tint value of the color from -1 to 1. Particularly useful for theme colors. 0.0 means no tint, -1.0 means 100% darken, and 1.0 means 100% lighten.|

### NOTOC-Borders
|Property|Type|Description|
| ------------- | ------------- | ----- |
|[left]|`Border|string|boolean`|The border settings for the left side. If string, with set border to the given border style. If true, will set border style to `'thin'`.|
|[right]|`Border|string|boolean`|The border settings for the right side. If string, with set border to the given border style. If true, will set border style to `'thin'`.|
|[top]|`Border|string|boolean`|The border settings for the top side. If string, with set border to the given border style. If true, will set border style to `'thin'`.|
|[bottom]|`Border|string|boolean`|The border settings for the bottom side. If string, with set border to the given border style. If true, will set border style to `'thin'`.|
|[diagonal]|`Border|string|boolean`|The border settings for the diagonal side. If string, with set border to the given border style. If true, will set border style to `'thin'`.|

### NOTOC-Border
|Property|Type|Description|
| ------------- | ------------- | ----- |
|style|`string`|Style of the given border.|
|color|`Color|string|number`|Color of the given border. If string, will set an RGB color. If number, will set a theme color.|
|[direction]|`string`|For diagonal border, the direction of the border(s) from left to right. Allowed values: `'up'`, `'down'`, `'both'`|

### NOTOC-SolidFill
|Property|Type|Description|
| ------------- | ------------- | ----- |
|type|`'solid'`||
|color|`Color|string|number`|Color of the fill. If string, will set an RGB color. If number, will set a theme color.|

### NOTOC-PatternFill
|Property|Type|Description|
| ------------- | ------------- | ----- |
|type|`'pattern'`||
|pattern|`string`|Name of the pattern. Allowed values: `'gray125'`, `'darkGray'`, `'mediumGray'`, `'lightGray'`, `'gray0625'`, `'darkHorizontal'`, `'darkVertical'`, `'darkDown'`, `'darkUp'`, `'darkGrid'`, `'darkTrellis'`, `'lightHorizontal'`, `'lightVertical'`, `'lightDown'`, `'lightUp'`, `'lightGrid'`, `'lightTrellis'`.|
|foreground|`Color|string|number`|Color of the foreground. If string, will set an RGB color. If number, will set a theme color.|
|background|`Color|string|number`|Color of the background. If string, will set an RGB color. If number, will set a theme color.|

### NOTOC-GradientFill
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
<!-- api -->
