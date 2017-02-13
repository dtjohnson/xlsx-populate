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
<dt><a href="#ContentTypes">ContentTypes</a></dt>
<dd><p>A content type collection.</p>
</dd>
<dt><a href="#Relationships">Relationships</a></dt>
<dd><p>A relationship collection.</p>
</dd>
<dt><a href="#Row">Row</a></dt>
<dd><p>A row.</p>
</dd>
<dt><a href="#SharedStrings">SharedStrings</a></dt>
<dd><p>The shared strings table.</p>
</dd>
<dt><a href="#Style">Style</a></dt>
<dd><p>A style.</p>
</dd>
<dt><a href="#_StyleSheet">_StyleSheet</a></dt>
<dd><p>A style sheet.</p>
</dd>
</dl>

### Constants

<dl>
<dt><a href="#STANDARD_CODES">STANDARD_CODES</a></dt>
<dd><p>Standard number format codes
Taken from <a href="http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/">http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/</a></p>
</dd>
<dt><a href="#STARTING_CUSTOM_NUMBER_FORMAT_ID">STARTING_CUSTOM_NUMBER_FORMAT_ID</a></dt>
<dd><p>The starting ID for custom number formats. The first 163 indexes are reserved.</p>
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
<a name="ContentTypes"></a>

### ContentTypes
A content type collection.

**Kind**: global class  

* [ContentTypes](#ContentTypes)
    * [new ContentTypes(node)](#new_ContentTypes_new)
    * [.add(partName, contentType)](#ContentTypes+add) ⇒ <code>Object</code>
    * [.findByPartName(partName)](#ContentTypes+findByPartName) ⇒ <code>Object</code> &#124; <code>undefined</code>
    * [.toObject()](#ContentTypes+toObject) ⇒ <code>Object</code>

<a name="new_ContentTypes_new"></a>

#### new ContentTypes(node)
Creates a new instance of _ContentTypes


| Param | Type | Description |
| --- | --- | --- |
| node | <code>Object</code> | The node. |

<a name="ContentTypes+add"></a>

#### contentTypes.add(partName, contentType) ⇒ <code>Object</code>
Add a new content type.

**Kind**: instance method of <code>[ContentTypes](#ContentTypes)</code>  
**Returns**: <code>Object</code> - The new content type.  

| Param | Type | Description |
| --- | --- | --- |
| partName | <code>string</code> | The part name. |
| contentType | <code>string</code> | The content type. |

<a name="ContentTypes+findByPartName"></a>

#### contentTypes.findByPartName(partName) ⇒ <code>Object</code> &#124; <code>undefined</code>
Find a content type by part name.

**Kind**: instance method of <code>[ContentTypes](#ContentTypes)</code>  
**Returns**: <code>Object</code> &#124; <code>undefined</code> - The matching content type or undefined if not found.  

| Param | Type | Description |
| --- | --- | --- |
| partName | <code>string</code> | The part name. |

<a name="ContentTypes+toObject"></a>

#### contentTypes.toObject() ⇒ <code>Object</code>
Convert the collection to an object.

**Kind**: instance method of <code>[ContentTypes](#ContentTypes)</code>  
**Returns**: <code>Object</code> - The object.  
<a name="Relationships"></a>

### Relationships
A relationship collection.

**Kind**: global class  

* [Relationships](#Relationships)
    * [new Relationships(node)](#new_Relationships_new)
    * [.add(type, target)](#Relationships+add) ⇒ <code>Object</code>
    * [.findByType(type)](#Relationships+findByType) ⇒ <code>Object</code> &#124; <code>undefined</code>
    * [.toObject()](#Relationships+toObject) ⇒ <code>Object</code>

<a name="new_Relationships_new"></a>

#### new Relationships(node)
Creates a new instance of _Relationships.


| Param | Type | Description |
| --- | --- | --- |
| node | <code>Object</code> | The node. |

<a name="Relationships+add"></a>

#### relationships.add(type, target) ⇒ <code>Object</code>
Add a new relationship.

**Kind**: instance method of <code>[Relationships](#Relationships)</code>  
**Returns**: <code>Object</code> - The new relationship.  

| Param | Type | Description |
| --- | --- | --- |
| type | <code>string</code> | The type of relationship. |
| target | <code>string</code> | The target of the relationship. |

<a name="Relationships+findByType"></a>

#### relationships.findByType(type) ⇒ <code>Object</code> &#124; <code>undefined</code>
Find a relationship by type.

**Kind**: instance method of <code>[Relationships](#Relationships)</code>  
**Returns**: <code>Object</code> &#124; <code>undefined</code> - The matching relationship or undefined if not found.  

| Param | Type | Description |
| --- | --- | --- |
| type | <code>string</code> | The type to search for. |

<a name="Relationships+toObject"></a>

#### relationships.toObject() ⇒ <code>Object</code>
Convert the collection to an object.

**Kind**: instance method of <code>[Relationships](#Relationships)</code>  
**Returns**: <code>Object</code> - The object.  
<a name="Row"></a>

### Row
A row.

**Kind**: global class  
<a name="SharedStrings"></a>

### SharedStrings
The shared strings table.

**Kind**: global class  

* [SharedStrings](#SharedStrings)
    * [new SharedStrings(node)](#new_SharedStrings_new)
    * [.getIndexForString(string)](#SharedStrings+getIndexForString) ⇒ <code>number</code>
    * [.getStringByIndex(index)](#SharedStrings+getStringByIndex) ⇒ <code>string</code>
    * [.toObject()](#SharedStrings+toObject) ⇒ <code>Object</code>

<a name="new_SharedStrings_new"></a>

#### new SharedStrings(node)
Constructs a new instance of _SharedStrings.


| Param | Type | Description |
| --- | --- | --- |
| node | <code>Object</code> | The node. |

<a name="SharedStrings+getIndexForString"></a>

#### sharedStrings.getIndexForString(string) ⇒ <code>number</code>
Gets the index for a string

**Kind**: instance method of <code>[SharedStrings](#SharedStrings)</code>  
**Returns**: <code>number</code> - The index  

| Param | Type | Description |
| --- | --- | --- |
| string | <code>string</code> | The string |

<a name="SharedStrings+getStringByIndex"></a>

#### sharedStrings.getStringByIndex(index) ⇒ <code>string</code>
Get the string for a given index

**Kind**: instance method of <code>[SharedStrings](#SharedStrings)</code>  
**Returns**: <code>string</code> - The string  

| Param | Type | Description |
| --- | --- | --- |
| index | <code>number</code> | The index |

<a name="SharedStrings+toObject"></a>

#### sharedStrings.toObject() ⇒ <code>Object</code>
Convert the collection to an object.

**Kind**: instance method of <code>[SharedStrings](#SharedStrings)</code>  
**Returns**: <code>Object</code> - The object.  
<a name="Style"></a>

### Style
A style.

**Kind**: global class  

* [Style](#Style)
    * [new Style(styleSheet, id, xfNode, fontNode, fillNode, borderNode)](#new_Style_new)
    * [.style(name, [value])](#Style+style) ⇒ <code>\*</code> &#124; <code>[Style](#Style)</code>

<a name="new_Style_new"></a>

#### new Style(styleSheet, id, xfNode, fontNode, fillNode, borderNode)
Creates a new instance of _Style.


| Param | Type | Description |
| --- | --- | --- |
| styleSheet | <code>StyleSheet</code> | The styleSheet. |
| id | <code>number</code> | The style ID. |
| xfNode | <code>Object</code> | The xf node. |
| fontNode | <code>Object</code> | The font node. |
| fillNode | <code>Object</code> | The fill node. |
| borderNode | <code>Object</code> | The border node. |

<a name="Style+style"></a>

#### style.style(name, [value]) ⇒ <code>\*</code> &#124; <code>[Style](#Style)</code>
Gets or sets a style.

**Kind**: instance method of <code>[Style](#Style)</code>  
**Returns**: <code>\*</code> &#124; <code>[Style](#Style)</code> - The value if getting or the style if setting.  

| Param | Type | Description |
| --- | --- | --- |
| name | <code>string</code> | The style name. |
| [value] | <code>\*</code> | The value to set. |

<a name="_StyleSheet"></a>

### _StyleSheet
A style sheet.

**Kind**: global class  

* [_StyleSheet](#_StyleSheet)
    * [new _StyleSheet(node)](#new__StyleSheet_new)
    * [.createStyle([sourceId])](#_StyleSheet+createStyle) ⇒ <code>[Style](#Style)</code>
    * [.getNumberFormatCode(id)](#_StyleSheet+getNumberFormatCode) ⇒ <code>string</code>
    * [.getNumberFormatId(code)](#_StyleSheet+getNumberFormatId) ⇒ <code>number</code>
    * [.toObject()](#_StyleSheet+toObject) ⇒ <code>string</code>

<a name="new__StyleSheet_new"></a>

#### new _StyleSheet(node)
Creates an instance of _StyleSheet.


| Param | Type | Description |
| --- | --- | --- |
| node | <code>string</code> | The style sheet node |

<a name="_StyleSheet+createStyle"></a>

#### _StyleSheet.createStyle([sourceId]) ⇒ <code>[Style](#Style)</code>
Create a style.

**Kind**: instance method of <code>[_StyleSheet](#_StyleSheet)</code>  
**Returns**: <code>[Style](#Style)</code> - The style.  

| Param | Type | Description |
| --- | --- | --- |
| [sourceId] | <code>number</code> | The source style ID to copy, if provided. |

<a name="_StyleSheet+getNumberFormatCode"></a>

#### _StyleSheet.getNumberFormatCode(id) ⇒ <code>string</code>
Get the number format code for a given ID.

**Kind**: instance method of <code>[_StyleSheet](#_StyleSheet)</code>  
**Returns**: <code>string</code> - The format code.  

| Param | Type | Description |
| --- | --- | --- |
| id | <code>number</code> | The number format ID. |

<a name="_StyleSheet+getNumberFormatId"></a>

#### _StyleSheet.getNumberFormatId(code) ⇒ <code>number</code>
Get the nuumber format ID for a given code.

**Kind**: instance method of <code>[_StyleSheet](#_StyleSheet)</code>  
**Returns**: <code>number</code> - The number format ID.  

| Param | Type | Description |
| --- | --- | --- |
| code | <code>string</code> | The format code. |

<a name="_StyleSheet+toObject"></a>

#### _StyleSheet.toObject() ⇒ <code>string</code>
Convert the style sheet to an XML string.

**Kind**: instance method of <code>[_StyleSheet](#_StyleSheet)</code>  
**Returns**: <code>string</code> - The XML string.  
<a name="STANDARD_CODES"></a>

### STANDARD_CODES
Standard number format codesTaken from http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/

**Kind**: global constant  
<a name="STARTING_CUSTOM_NUMBER_FORMAT_ID"></a>

### STARTING_CUSTOM_NUMBER_FORMAT_ID
The starting ID for custom number formats. The first 163 indexes are reserved.

**Kind**: global constant  

