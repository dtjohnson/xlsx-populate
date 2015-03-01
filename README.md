[![view on npm](http://img.shields.io/npm/v/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![npm module downloads per month](http://img.shields.io/npm/dm/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![Build Status](https://travis-ci.org/dtjohnson/xlsx-populate.svg?branch=master)](https://travis-ci.org/dtjohnson/xlsx-populate)
[![Dependency Status](https://david-dm.org/dtjohnson/xlsx-populate.svg)](https://david-dm.org/dtjohnson/xlsx-populate)

# xlsx-populate
Node.js module to populate Excel XLSX templates. This module does no parse Excel workbooks. There are [good modules](https://github.com/SheetJS/js-xlsx) for this already. The purpose of this module is to open existing Excel XLSX workbook templates that have styling in place and populate with data.

## Installation

    $ npm install xlsx-populate

## Usage
```js
var Workbook = require('xlsx-populate');

// Load the input workbook from file.
var workbook = Workbook.fromFileSync("./Book1.xlsx");

// Modify the workbook.
workbook.getSheet("Sheet1").getCell("A1").setValue("This is neat!");

// Write to file.
workbook.toFileSync("./out.xlsx");
```
