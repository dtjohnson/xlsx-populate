[![view on npm](http://img.shields.io/npm/v/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![npm module downloads per month](http://img.shields.io/npm/dm/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![Build Status](https://travis-ci.org/dtjohnson/xlsx-populate.svg?branch=master)](https://travis-ci.org/dtjohnson/xlsx-populate)
[![Dependency Status](https://david-dm.org/dtjohnson/xlsx-populate.svg)](https://david-dm.org/dtjohnson/xlsx-populate)

# xlsx-populate
Excel XLSX parser/generator written in JavaScript with Node.js and browser support, jQuery/d3-style method chaining, encryption, and a focus on keeping existing workbook features and styles in tact.

## Table of Contents
<!-- toc -->

## Installation

### Node.js
```bash
npm install xlsx-populate
```
Note that xlsx-populate uses ES6 features so only Node.js v4+ is supported.

### Browser

A functional browser example can be found in [examples/browser/index.html](https://gitcdn.xyz/repo/dtjohnson/xlsx-populate/master/examples/browser/index.html).

xlsx-populate is written first for Node.js. We use [browserify](http://browserify.org/) and [babelify](https://github.com/babel/babelify) to transpile and pack up the module for use in the browser.

You have a number of options to include the code in the browser. You can download the combined, minified code from the browser directory in this repository or you can install with bower:
```bash
bower install xlsx-populate
```
After including the module in the browser, it is available globally as `XlsxPopulate`.

Alternatively, you can require this module using [browserify](http://browserify.org/). Since xlsx-populate uses ES6 features, you will also need to use [babelify](https://github.com/babel/babelify) with [babel-preset-env](https://www.npmjs.com/package/babel-preset-env).

## Usage

xlsx-populate has an [extensive API](#api-reference) for working with Excel workbooks. This section reviews the most common functions and use cases. Examples can also be found in the examples directory of the source code.

### Populating Data

To populate data in a workbook, you first load one (either blank, from data, or from file). Then you can access sheets and
 cells within the workbook to manipulate them.
```js
const XlsxPopulate = require('xlsx-populate');

// Load a new blank workbook
XlsxPopulate.fromBlankAsync()
    .then(workbook => {
        // Modify the workbook.
        workbook.sheet("Sheet1").cell("A1").value("This is neat!");

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    });
```

### Parsing Data

You can pull data out of existing workbooks using [Cell.value](#Cell+value) as a getter without any arguments:
```js
const XlsxPopulate = require('xlsx-populate');

// Load an existing workbook
XlsxPopulate.fromFileAsync("./Book1.xlsx")
    .then(workbook => {
        // Modify the workbook.
        const value = workbook.sheet("Sheet1").cell("A1").value();

        // Log the value.
        console.log(value);
    });
```
__Note__: in cells that contain values calculated by formulas, Excel will store the calculated value in the workbook. The [value](#Cell+value) method will return the value of the cells at the time the workbook was saved. xlsx-populate will _not_ recalculate the values as you manipulate the workbook and will _not_ write the values to the output.

### Ranges
xlsx-populate also supports ranges of cells to allow parsing/manipulation of multiple cells at once.
```js
const r = workbook.sheet(0).range("A1:C3");

// Set all cell values to the same value:
r.value(5);

// Set the values using a 2D array:
r.value([
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
]);

// Set the values using a callback function:
r.value((cell, ri, ci, range) => Math.random());
```

A common use case is to simply pull all of the values out all at once. You can easily do that with the [Sheet.usedRange](#Sheet+usedRange) method.
```js
// Get 2D array of all values in the worksheet.
const values = workbook.sheet("Sheet1").usedRange().value();
```

Alternatively, you can set the values in a range with only the top-left cell in the range:
```js
workbook.sheet(0).cell("A1").value([
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
]);
```
The set range is returned.

### Rows and Columns

You can access rows and columns in order to change size, hide/show, or access cells within:
```js
// Get the B column, set its width and unhide it (assuming it was hidden).
sheet.column("B").width(25).hidden(false);

const cell = sheet.row(5).cell(3); // Returns the cell at C5.
```

### Managing Sheets
xlsx-populate supports a number of options for managing sheets.

You can get a sheet by name or index or get all of the sheets as an array:
```js
// Get sheet by index
const sheet1 = workbook.sheet(0);

// Get sheet by name
const sheet2 = workbook.sheet("Sheet2");

// Get all sheets as an array
const sheets = workbook.sheets();
```

You can add new sheets:
```js
// Add a new sheet named 'New 1' at the end of the workbook
const newSheet1 = workbook.addSheet('New 1');

// Add a new sheet named 'New 2' at index 1 (0-based)
const newSheet2 = workbook.addSheet('New 2', 1);

// Add a new sheet named 'New 3' before the sheet named 'Sheet1'
const newSheet3 = workbook.addSheet('New 3', 'Sheet1');

// Add a new sheet named 'New 4' before the sheet named 'Sheet1' using a Sheet reference.
const sheet = workbook.sheet('Sheet1');
const newSheet4 = workbook.addSheet('New 4', sheet);
```
*Note: the sheet rename method does not rename references to the sheet so formulas, etc. can be broken. Use with caution!*

You can rename sheets:
```js
// Rename the first sheet.
const sheet = workbook.sheet(0).name("new sheet name");
```

You can move sheets:
```js
// Move 'Sheet1' to the end
workbook.moveSheet("Sheet1");

// Move 'Sheet1' to index 2
workbook.moveSheet("Sheet1", 2);

// Move 'Sheet1' before 'Sheet2'
workbook.moveSheet("Sheet1", "Sheet2");
```
The above methods can all use sheet references instead of names as well. And you can also move a sheet using a method on the sheet:
```js
// Move the sheet before 'Sheet2'
sheet.move("Sheet2");
```

You can delete sheets:
```js
// Delete 'Sheet1'
workbook.deleteSheet("Sheet1");

// Delete sheet with index 2
workbook.deleteSheet(2);

// Delete from sheet reference
workbook.sheet(0).delete();
```

You can get/set the active sheet:
```js
// Get the active sheet
const sheet = workbook.activeSheet();

// Check if the current sheet is active
sheet.active() // returns true or false

// Activate the sheet
sheet.active(true);

// Or from the workbook
workbook.activeSheet("Sheet2");
```

### Defined Names
Excel supports creating defined names that refer to addresses, formulas, or constants. These defined names can be scoped
to the entire workbook or just individual sheets. xlsx-populate supports looking up defined names that refer to cells or
ranges. (Dereferencing other names will result in an error.) Defined names are particularly useful if you are populating
data into a known template. Then you do not need to know the exact location.

```js
// Look up workbook-scoped name and set the value to 5.
workbook.definedName("some name").value(5);

// Look of a name scoped to the first sheet and set the value to "foo".
workbook.sheet(0).definedName("some other name").value("foo");
```

You can also create, modify, or delete defined names:
```js
// Create/modify a workbook-scope defined name
workbook.definedName("some name", "TRUE");

// Delete a sheet-scoped defined name:
workbook.sheet(0).definedName("some name", null);
```

### Find and Replace
You can search for occurrences of text in cells within the workbook or sheets and optionally replace them.
```js
// Find all occurrences of the text "foo" in the workbook and replace with "bar".
workbook.find("foo", "bar"); // Returns array of matched cells

// Find the matches but don't replace.
workbook.find("foo");

// Just look in the first sheet.
workbook.sheet(0).find("foo");

// Check if a particular cell matches the value.
workbook.sheet("Sheet1").cell("A1").find("foo"); // Returns true or false
```

Like [String.replace](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/replace), the find method can also take a RegExp search pattern and replace can take a function callback:
```js
// Use a RegExp to replace all lowercase letters with uppercase
workbook.find(/[a-z]+/g, match => match.toUpperCase());
```

### Styles
xlsx-populate supports a wide range of cell formatting. See the [Style Reference](#style-reference) for the various options.

To get/set a cell style:
```js
// Get a single style
const bold = cell.style("bold"); // true

// Get multiple styles
const styles = cell.style(["bold", "italic"]); // { bold: true, italic: true }

// Set a single style
cell.style("bold", true);

// Set multiple styles
cell.style({ bold: true, italic: true });
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

If you are setting styles for many cells, performance is far better if you set for an entire row or column:
```js
// Set a single style
sheet.row(1).style("bold", true);

// Set multiple styles
sheet.column("A").style({ bold: true, italic: true });

// Get a single style
const bold = sheet.column(3).style("bold");

// Get multiple styles
const styles = sheet.row(5).style(["bold", "italic"]);
```
Note that the row/column style behavior mirrors Excel. Setting a style on a column will apply that style to all existing cells and any new cells that are populated. Getting the row/column style will return only the styles that have been applied to the entire row/column, not the styles of every cell in the row or column.

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

Number formats are one of the most common styles. They can be set using the `numberFormat` style.
```js
cell.style("numberFormat", "0.00");
```

Information on how number format codes work can be found [here](https://support.office.com/en-us/article/Number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=en-US&rs=en-US&ad=US).
You can also look up the desired format code in Excel:
* Right-click on a cell in Excel with the number format you want.
* Click on "Format Cells..."
* Switch the category to "Custom" if it is not already.
* The code in the "Type" box is the format you should copy.

### Rich Texts
You can read/write rich texts to cells.

#### Supported styles
`bold`, `italic`, `underline`, `strikethrough`, `subscript`, `fontSize`,
`fontFamily`, `fontGenericFamily`, `fontScheme`, `fontColor`.
See the [Style Reference](#style-reference) for the various options.

#### Usage
You can read and modify rich texts on an existing rich text cell:
```js
// assume A1 is a rich text cell
const RichText = require('xlsx-Populate').RichText;
const cell = workbook.sheet(0).cell('A1');
cell.value() instanceof RichText // returns true
const richtext = cell.value();
// get the concatenate text
richtext.text();

// loop through each rich text fragment
for (let i = 0; i < richtext.length; i++) {
    const fragment = richtext.get(i);
    // Get the style
    fragment.style('bold');
    // Get many styles
    fragment.style(['bold', 'italic']);
    // Set one style
    fragment.style('bold', true);
    // Set many styles
    fragment.style({ 'bold': true, 'italic': true });
    // Get the value
    fragment.value();
    // Set the value
    fragment.value('hello');
}

// remove the first rich text fragment
richtext.remove(0);

// clear this rich texts
richtext.clear();
```

How to set a cell to rich texts:
```js
const RichText = require('xlsx-Populate').RichText;
const cell = workbook.sheet(0).cell('A1');
// set a cell value to rich text
cell.value(new RichText());

// add two rich text fragments
cell.value()
    .add('hello ', { italic: true, bold: true })
    .add('world!', { fontColor: 'FF0000' });
````

You can specify the index when adding rich text fragment.
```js
// add before the first fragment
cell.value().add('text', { bold: true }, 0);
// add before the second fragment
cell.value().add('text', { bold: true }, 1);
// add after the last fragment
cell.value().add('text', { bold: true });
```
#### Notes
We make a deep copy of the richtext instance when assign it to a cell, which
means you can only modify the content of the richtext before calling `cell.value(richtext)`. 
Any modification to the richtext instance after calling `cell.value(richtext)` will not
save to the cell. i.e.
```js
const richtext = new RichText();
richtext.add('hello');
cell.value(richtext);
cell.value().text(); // returns 'hello'

richtext.add(' world')
richtext.text(); // returns 'hello world' 
cell.value().text(); // returns 'hello'
cell.value() === richtext; // returns false

cell.value().add(' world');
cell.value().text(); // returns 'hello world'
```

This means you can create a rich text instance and assign it to any cells! Each cell does
not share the same instance but creates a deep copy of the instance.
```js
const sheet = workbook.sheet(0);
const richtext = new RichText();
richtext.add('hello');
const range = sheet.range("A1:C3");
range.value(richtext);
// they do not share the same instance
sheet.cell('A1').value() === sheet.cell('C1').value() // returns false
```

You can get the rich text from a cell and set it to anoher cell.
```js
const richtext = cell1.value();
cell2.value(richtext);
cell1.value() === cell2.value() // returns false
```

Whenever you call `richtext.add(text, styles, index)`, we will detect if the given `text`
contains line separators (`\n`, `\r`, `\r\n`), if it does, we will call
`cell.style('wrapText', true)` for you. MS Excel needs wrapText to be true
to have the new lines displayed, otherwise you will see the texts in one line.
You may also need to set row height to have all lines displayed.
```js
cell.value()
    // it support all line separators
    .add('123\n456\r789\r\n10', { italic: true, fontColor: '123456' })
// remember to set height to show the whole row
workbook.sheet(0).row(1).height(100);
```

### Dates

Excel stores date/times as the number of days since 1/1/1900 ([sort of](https://en.wikipedia.org/wiki/Leap_year_bug)). It just applies a number formatting to make the number appear as a date. So to set a date value, you will need to also set a number format for a date if one doesn't already exist in the cell:
```js
cell.value(new Date(2017, 1, 22)).style("numberFormat", "dddd, mmmm dd, yyyy");
```
When fetching the value of the cell, it will be returned as a number. To convert it to a date use [XlsxPopulate.numberToDate](#XlsxPopulate.numberToDate):
```js
const num = cell.value(); // 42788
const date = XlsxPopulate.numberToDate(num); // Wed Feb 22 2017 00:00:00 GMT-0500 (Eastern Standard Time)
```

### Data Validation
Data validation is also supported. To set/get/remove a cell data validation:
```js
// Set the data validation
cell.dataValidation({
    type: 'list',
    allowBlank: false,
    showInputMessage: false,
    prompt: false,
    promptTitle: 'String',
    showErrorMessage: false,
    error: 'String',
    errorTitle: 'String',
    operator: 'String',
    formula1: '$A:$A',//Required
    formula2: 'String'
});

//Here is a short version of the one above.
cell.dataValidation('$A:$A');

// Get the data validation
const obj = cell.dataValidation(); // Returns an object

// Remove the data validation
cell.dataValidation(null); //Returns the cell
```

Similarly for ranges:
```js

// Set all cells in range with a single shared data validation
range.dataValidation({
    type: 'list',
    allowBlank: false,
    showInputMessage: false,
    prompt: false,
    promptTitle: 'String',
    showErrorMessage: false,
    error: 'String',
    errorTitle: 'String',
    operator: 'String',
    formula1: 'Item1,Item2,Item3,Item4',//Required
    formula2: 'String'
});

//Here is a short version of the one above.
range.dataValidation('Item1,Item2,Item3,Item4');

// Get the data validation
const obj = range.dataValidation(); // Returns an object

// Remove the data validation
range.dataValidation(null); //Returns the Range
```
Please note, the data validation gets applied to the entire range, *not* each Cell in the range.

### Method Chaining

xlsx-populate uses method-chaining similar to that found in [jQuery](https://jquery.com/) and [d3](https://d3js.org/). This lets you construct large chains of setters as desired:
```js
workbook
    .sheet(0)
        .cell("A1")
            .value("foo")
            .style("bold", true)
        .relativeCell(1, 0)
            .formula("A1")
            .style("italic", true)
.workbook()
    .sheet(1)
        .range("A1:B3")
            .value(5)
        .cell(0, 0)
            .style("underline", "double");

```

### Hyperlinks
Hyperlinks are also supported on cells using the [Cell.hyperlink](#Cell+hyperlink) method. The method will _not_ style the content to look like a hyperlink. You must do that yourself:
```js
// Set a hyperlink
cell.value("Link Text")
    .style({ fontColor: "0563c1", underline: true })
    .hyperlink("http://example.com");

// Set a hyperlink with tooltip
cell.value("Link Text")
    .style({ fontColor: "0563c1", underline: true })
    .hyperlink({ hyperlink: "http://example.com", tooltip: "example.com" });

// Get the hyperlink
const value = cell.hyperlink(); // Returns 'http://example.com'

// Set a hyperlink to email
cell.value("Click to Email Jeff Bezos")
    .hyperlink({ email: "jeff@amazon.com", emailSubject: "I know you're a busy man Jeff, but..." });

// Set a hyperlink to an internal cell using an address string.
cell.value("Click to go to an internal cell")
    .hyperlink("Sheet2!A1");

// Set a hyperlink to an internal cell using a cell object.
cell.value("Click to go to an internal cell")
    .hyperlink(workbook.sheet(0).cell("A1"));
```

### Print Options
Print options are accessed using the [Sheet.printOptions](#Sheet+printOptions) method. Defaults are all assumed to be false, so if the attribute is missing, then the method returns false. A method [Sheet.printGridLines](#Sheet+printGridLines) is provided to offer the convenience of setting both gridLines and gridLinesSet.
```js
// Print row and column headings
sheet.printOptions('headings', true);

// Get the headings flag
const headings = sheet.printOptions('headings'); // Returns true

// Clear flag for center on page vertically when printing
sheet.printOptions('verticalCentered', undefined);

// Get the verticalCentered flag
const verticalCentered = sheet.printOptions('verticalCentered'); // Returns false

// Enable grid lines in print
sheet.printGridLines(true);

// Now both gridLines and gridLinesSet print options are set
sheet.printOptions('gridLines') === sheet.printOptions('gridLinesSet') === true; // Returns true

// To disable, just disable one of gridLines or gridLinesSet
sheet.printOptions('gridLineSets', false);

const isPrintGridLinesEnabled = sheet.printGridLines(); // Returns false
```

### Page Margins
Excel requires that all page margins are defined or none at all. To ensure this, please choose an existing or custom preset. See [Sheet.pageMarginsPreset](#Sheet+pageMarginsPreset).

```js
// Get the current preset
sheet.pageMarginsPreset(); // Returns undefined

// Switch to an existing preset
sheet.pageMarginsPreset('normal');
```

Page margins are accessed using the [Sheet.pageMargins](#Sheet+pageMargins) method. If a page margin is not set, the preset will fill in the gaps.

```js
// Get top margin in inches, note that the current preset is currently set to normal (see above)
sheet.pageMargins('top'); // Returns 0.75

// Set top page margin in inches
sheet.pageMargins('top', 1.1);

// Get top page margin in inches.
const topPageMarginInInches = sheet.pageMargins('top'); // Returns 1.1
```

### SheetView Panes
SheetView Panes are accessed using the [Sheet.panes](#Sheet+panes) method.
For convenience, we have [Sheet.freezePanes](#Sheet+freezePanes),
[Sheet.splitPanes](#Sheet+splitPanes), [Sheet.resetPanes](#Sheet+resetPanes),
and type [PaneOptions](#paneoptions--object).
```js
// access Pane options
sheet.panes(); // return PaneOptions Object

// manually Set Pane options, WARNING: setting wrong options may result in excel fails to open.
const paneOptions = { state: 'frozen', topLeftCell: 'B2', xSplit: 1, ySplit: 1, activePane: 'bottomRight' }
sheet.panes(paneOptions); // return PaneOptions Object

// freeze panes (freeze first column and first two rows)
sheet.freezePanes(1, 2);
// OR
sheet.freezePanes('B3');

// split panes (Horizontal Split Position: 1000 / 20 pt, Vertical Split Position: 2000 / 20 pt)
sheet.splitPanes(1000, 2000);

// reset to normal panes (no freeze panes and split panes)
sheet.resetPanes();
```

### Serving from Express
You can serve the workbook from [express](http://expressjs.com/) or other web servers with something like this:
```js
router.get("/download", function (req, res, next) {
    // Open the workbook.
    XlsxPopulate.fromFileAsync("input.xlsx")
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
Usage in the browser is almost the same. A functional example can be found in [examples/browser/index.html](https://gitcdn.xyz/repo/dtjohnson/xlsx-populate/master/examples/browser/index.html). The library is exposed globally as `XlsxPopulate`. Existing workbooks can be loaded from a file:
```js
// Assuming there is a file input in the page with the id 'file-input'
var file = document.getElementById("file-input").files[0];

// A File object is a special kind of blob.
XlsxPopulate.fromDataAsync(file)
    .then(function (workbook) {
        // ...
    });
```

You can also load from AJAX if you set the responseType to 'arraybuffer':
```js
var req = new XMLHttpRequest();
req.open("GET", "http://...", true);
req.responseType = "arraybuffer";
req.onreadystatechange = function () {
    if (req.readyState === 4 && req.status === 200){
        XlsxPopulate.fromDataAsync(req.response)
            .then(function (workbook) {
                // ...
            });
    }
};

req.send();
```

To download the workbook, you can either export as a blob (default behavior) or as a base64 string. You can then insert a link into the DOM and click it:
```js
workbook.outputAsync()
    .then(function (blob) {
        if (window.navigator && window.navigator.msSaveOrOpenBlob) {
            // If IE, you must uses a different method.
            window.navigator.msSaveOrOpenBlob(blob, "out.xlsx");
        } else {
            var url = window.URL.createObjectURL(blob);
            var a = document.createElement("a");
            document.body.appendChild(a);
            a.href = url;
            a.download = "out.xlsx";
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        }
    });
```

Alternatively, you can download via a data URI, but this is not supported by IE:
```js
workbook.outputAsync("base64")
    .then(function (base64) {
        location.href = "data:" + XlsxPopulate.MIME_TYPE + ";base64," + base64;
    });
```

### Promises
xlsx-populate uses [promises](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise) to manage async input/output. By default it uses the `Promise` defined in the browser or Node.js. In browsers that don't support promises (IE) a [polyfill is used via JSZip](https://stuk.github.io/jszip/documentation/api_jszip/external.html).
```js
// Get the current promise library in use.
// Helpful for getting a usable Promise library in IE.
var Promise = XlsxPopulate.Promise;
```

If you prefer, you can override the default `Promise` library used with another ES6 compliant library like [bluebird](http://bluebirdjs.com/).
```js
const Promise = require("bluebird");
const XlsxPopulate = require("xlsx-populate");
XlsxPopulate.Promise = Promise;
```

### Encryption
XLSX Agile encryption and descryption are supported so you can read and write password-protected workbooks. To read a protected workbook, pass the password in as an option:
```js
XlsxPopulate.fromFileAsync("./Book1.xlsx", { password: "S3cret!" })
    .then(workbook => {
        // ...
    });
```

Similarly, to write a password encrypted workbook:
```js
workbook.toFileAsync("./out.xlsx", { password: "S3cret!" });
```
The password option is supported in all output methods. N.B. Workbooks will only be encrypted if you supply a password when outputting even if they had a password when reading.

Encryption support is also available in the browser, but take care! Any password you put in browser code can be read by anyone with access to your code. You should only use passwords that are supplied by the end-user. Also, the performance of encryption/decryption in the browser is far worse than with Node.js. IE, in particular, is extremely slow. xlsx-populate is bundled for browsers with and without encryption support as the encryption libraries increase the size of the bundle a lot.

## Missing Features
There are many, many features of the XLSX format that are not yet supported. If your use case needs something that isn't supported
please open an issue to show your support. Better still, feel free to [contribute](#contributing) a pull request!

## Submitting an Issue
If you happen to run into a bug or an issue, please feel free to [submit an issue](https://github.com/dtjohnson/xlsx-populate/issues). I only ask that you please include sample JavaScript code that demonstrates the issue.
If the problem lies with modifying some template, it is incredibly difficult to debug the issue without the template. So please attach the template if possible. If you have confidentiality concerns, please attach a different workbook that exhibits the issue or you can send your workbook directly to [dtjohnson](https://github.com/dtjohnson) after creating the issue.

## Contributing

Pull requests are very much welcome! If you'd like to contribute, please make sure to read this section carefully first.

### How xlsx-populate Works
An XLSX workbook is essentially a zip of a bunch of XML files. xlsx-populate uses [JSZip](https://stuk.github.io/jszip/)
to unzip the workbook and [sax-js](https://github.com/isaacs/sax-js) to parse the XML documents into corresponding objects.
As you call methods, xlsx-populate manipulates the content of those objects. When you generate the output, xlsx-populate
uses [xmlbuilder-js](https://github.com/oozcitak/xmlbuilder-js) to convert the objects back to XML and then uses JSZip to
rezip them back into a workbook.

The way in which xlsx-populate manipulates objects that are essentially the XML data is very different from the usual way
parser/generator libraries work. Most other libraries will deserialize the XML into a rich object model. That model is then
manipulated and serialized back into XML upon generation. The challenge with this approach is that the Office Open XML spec is [HUGE](http://www.ecma-international.org/publications/standards/Ecma-376.htm).
It is extremely difficult for libraries to be able to support the entire specification. So these other libraries will deserialize
only the portion of the spec they support and any other content/styles in the workbook they don't support are lost. Since
xlsx-populate just manipulates the XML data, it is able to preserve styles and other content while still only supporting
a fraction of the spec.

### Setting up your Environment
You'll need to make sure [Node.js](https://nodejs.org/en/) v4+ is installed (as xlsx-populate uses ES6 syntax). You'll also
need to install [gulp](https://github.com/gulpjs/gulp):
```bash
npm install -g gulp
```

Make sure you have [git](https://git-scm.com/) installed. Then follow [this guide](https://git-scm.com/book/en/v2/GitHub-Contributing-to-a-Project) to see how to check out code, branch, and
then submit your code as a pull request. When you check out the code, you'll first need to install the npm dependencies.
From the project root, run:
```bash
npm install
```

The default gulp task is set up to watch the source files for updates and retest while you edit. From the project root just run:
```bash
gulp
```

You should see the test output in your console window. As you edit files the tests will run again and show you if you've
broken anything. (Note that if you've added new files you'll need to restart gulp for the new files to be watched.)

Now write your code and make sure to add [Jasmine](https://jasmine.github.io/) unit tests. When you are finished, you need
to build the code for the browser. Do that by running the gulp build command:
```bash
gulp build
```

Verify all is working, check in your code, and submit a pull request.

### Pull Request Checklist
To make sure your code is consistent and high quality, please make sure to follow this checklist before submitting a pull request:
 * Your code must follow the getter/setter pattern using a single function for both. Check `arguments.length` or use `ArgHandler` to distinguish.
 * You must use valid [JSDoc](http://usejsdoc.org/) comments on *all* methods and classes. Use `@private` for private methods and `@ignore` for any public methods that are internal to xlsx-populate and should not be included in the public API docs.
 * You must adhere to the configured [ESLint](http://eslint.org/) linting rules. You can configure your IDE to display rule violations live or you can run `gulp lint` to see them.
 * Use [ES6](http://es6-features.org/#Constants) syntax. (This should be enforced by ESLint.)
 * Make sure to have full [Jasmine](https://jasmine.github.io/) unit test coverage for your code.
 * Make sure all tests pass successfully.
 * Whenever possible, do not modify/break existing API behavior. This module adheres to the [semantic versioning standard](https://docs.npmjs.com/getting-started/semantic-versioning). So any breaking changes will require a major release.
 * If your feature needs more documentation than just the JSDoc output, please add to the docs/template.md README file.


### Gulp Tasks

xlsx-populate uses [gulp](https://github.com/gulpjs/gulp) as a build tool. There are a number of tasks:

* __browser__ - Transpile and build client-side JavaScript project bundle using [browserify](http://browserify.org/) and [babelify](https://github.com/babel/babelify).
* __lint__ - Check project source code style using [ESLint](http://eslint.org/).
* __unit__ - Run [Jasmine](https://jasmine.github.io/) unit tests.
* __unit-browser__ - Run the unit tests in real browsers using [Karma](https://karma-runner.github.io/1.0/index.html).
* __e2e-parse__ - End-to-end tests of parsing data out of sample workbooks that were created in Microsoft Excel.
* __e2e-generate__ - End-to-end tests of generating workbooks using xlsx-populate. To verify the workbooks were truly generated correctly they need to be opened in Microsoft Excel and verified. This task automates this verification using the .NET Excel Interop library with [Edge.js](https://github.com/tjanczuk/edge) acting as a bridge between Node.js and C#. Note that these tests will _only_ run on Windows with Microsoft Excel and the [Primary Interop Assemblies installed](https://msdn.microsoft.com/en-us/library/kh3965hw.aspx).
* __e2e-browser__ - End-to-end tests of usage of the browserify bundle in real browsers using Karma.
* __blank__ - Convert a blank XLSX template into a JS buffer module to support [fromBlankAsync](#XlsxPopulate.fromBlankAsync).
* __docs__ - Build this README doc by combining docs/template.md, API docs generated with [jsdoc-to-markdown](https://github.com/jsdoc2md/jsdoc-to-markdown), and a table of contents generated with [markdown-toc](https://github.com/jonschlinkert/markdown-toc).
* __watch__ - Watch files for changes and then run associated gulp task. (Used by the default task.)
* __build__ - Run all gulp tasks, including linting and tests, and build the docs and browser bundle.
* __default__ - Run blank, unit, and docs tasks and watch the source files for those tasks for changes.

## Style Reference

### NOTOC-Styles
|Style Name|Type|Description|
| ------------- | ------------- | ----- |
|bold|`boolean`|`true` for bold, `false` for not bold|
|italic|`boolean`|`true` for italic, `false` for not italic|
|underline|<code>boolean&#124;string</code>|`true` for single underline, `false` for no underline, `'double'` for double-underline|
|strikethrough|`boolean`|`true` for strikethrough `false` for not strikethrough|
|subscript|`boolean`|`true` for subscript, `false` for not subscript (cannot be combined with superscript)|
|superscript|`boolean`|`true` for superscript, `false` for not superscript (cannot be combined with subscript)|
|fontSize|`number`|Font size in points. Must be greater than 0.|
|fontFamily|`string`|Name of font family.|
|fontGenericFamily|`number`|1: Serif, 2: Sans Serif, 3: Monospace, |
|fontScheme|`string`|`'minor'`\|`'major'`\|`'none'` |
|fontColor|<code>Color&#124;string&#124;number</code>|Color of the font. If string, will set an RGB color. If number, will set a theme color.|
|horizontalAlignment|`string`|Horizontal alignment. Allowed values: `'left'`, `'center'`, `'right'`, `'fill'`, `'justify'`, `'centerContinuous'`, `'distributed'`|
|justifyLastLine|`boolean`|a.k.a Justified Distributed. Only applies when horizontalAlignment === `'distributed'`. A boolean value indicating if the cells justified or distributed alignment should be used on the last line of text. (This is typical for East Asian alignments but not typical in other contexts.)|
|indent|`number`|Number of indents. Must be greater than or equal to 0.|
|verticalAlignment|`string`|Vertical alignment. Allowed values: `'top'`, `'center'`, `'bottom'`, `'justify'`, `'distributed'`|
|wrapText|`boolean`|`true` to wrap the text in the cell, `false` to not wrap.|
|shrinkToFit|`boolean`|`true` to shrink the text in the cell to fit, `false` to not shrink.|
|textDirection|`string`|Direction of the text. Allowed values: `'left-to-right'`, `'right-to-left'`|
|textRotation|`number`|Counter-clockwise angle of rotation in degrees. Must be [-90, 90] where negative numbers indicate clockwise rotation.|
|angleTextCounterclockwise|`boolean`|Shortcut for textRotation of 45 degrees.|
|angleTextClockwise|`boolean`|Shortcut for textRotation of -45 degrees.|
|rotateTextUp|`boolean`|Shortcut for textRotation of 90 degrees.|
|rotateTextDown|`boolean`|Shortcut for textRotation of -90 degrees.|
|verticalText|`boolean`|Special rotation that shows text vertical but individual letters are oriented normally. `true` to rotate, `false` to not rotate.|
|fill|<code>SolidFill&#124;PatternFill&#124;GradientFill&#124;Color&#124;string&#124;number</code>|The cell fill. If Color, will set a solid fill with the color. If string, will set a solid RGB fill. If number, will set a solid theme color fill.|
|border|<code>Borders&#124;Border&#124;string&#124;boolean</code>|The border settings. If string, will set outside borders to given border style. If true, will set outside border style to `'thin'`.|
|borderColor|<code>Color&#124;string&#124;number</code>|Color of the borders. If string, will set an RGB color. If number, will set a theme color.|
|borderStyle|`string`|Style of the outside borders. Allowed values: `'hair'`, `'dotted'`, `'dashDotDot'`, `'dashed'`, `'mediumDashDotDot'`, `'thin'`, `'slantDashDot'`, `'mediumDashDot'`, `'mediumDashed'`, `'medium'`, `'thick'`, `'double'`|
|leftBorder, rightBorder, topBorder, bottomBorder, diagonalBorder|<code>Border&#124;string&#124;boolean</code>|The border settings for the given side. If string, will set border to the given border style. If true, will set border style to `'thin'`.|
|leftBorderColor, rightBorderColor, topBorderColor, bottomBorderColor, diagonalBorderColor|<code>Color&#124;string&#124;number</code>|Color of the given border. If string, will set an RGB color. If number, will set a theme color.|
|leftBorderStyle, rightBorderStyle, topBorderStyle, bottomBorderStyle, diagonalBorderStyle|`string`|Style of the given side.|
|diagonalBorderDirection|`string`|Direction of the diagonal border(s) from left to right. Allowed values: `'up'`, `'down'`, `'both'`|
|numberFormat|`string`|Number format code. See docs [here](https://support.office.com/en-us/article/Number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=en-US&rs=en-US&ad=US).|

### NOTOC-Color
An object representing a color.

|Property|Type|Description|
| ------------- | ------------- | ----- |
|[rgb]|`string`|RGB color code (e.g. `'ff0000'`). Either rgb or theme is required.|
|[theme]|`number`|Index of a theme color. Either rgb or theme is required.|
|[tint]|`number`|Optional tint value of the color from -1 to 1. Particularly useful for theme colors. 0.0 means no tint, -1.0 means 100% darken, and 1.0 means 100% lighten.|

### NOTOC-Borders
An object representing all of the borders.

|Property|Type|Description|
| ------------- | ------------- | ----- |
|[left]|<code>Border&#124;string&#124;boolean</code>|The border settings for the left side. If string, will set border to the given border style. If true, will set border style to `'thin'`.|
|[right]|<code>Border&#124;string&#124;boolean</code>|The border settings for the right side. If string, will set border to the given border style. If true, will set border style to `'thin'`.|
|[top]|<code>Border&#124;string&#124;boolean</code>|The border settings for the top side. If string, will set border to the given border style. If true, will set border style to `'thin'`.|
|[bottom]|<code>Border&#124;string&#124;boolean</code>|The border settings for the bottom side. If string, will set border to the given border style. If true, will set border style to `'thin'`.|
|[diagonal]|<code>Border&#124;string&#124;boolean</code>|The border settings for the diagonal side. If string, will set border to the given border style. If true, will set border style to `'thin'`.|

### NOTOC-Border
An object representing an individual border.

|Property|Type|Description|
| ------------- | ------------- | ----- |
|style|`string`|Style of the given border.|
|color|<code>Color&#124;string&#124;number</code>|Color of the given border. If string, will set an RGB color. If number, will set a theme color.|
|[direction]|`string`|For diagonal border, the direction of the border(s) from left to right. Allowed values: `'up'`, `'down'`, `'both'`|

### NOTOC-SolidFill
An object representing a solid fill.

|Property|Type|Description|
| ------------- | ------------- | ----- |
|type|`'solid'`||
|color|<code>Color&#124;string&#124;number</code>|Color of the fill. If string, will set an RGB color. If number, will set a theme color.|

### NOTOC-PatternFill
An object representing a pattern fill.

|Property|Type|Description|
| ------------- | ------------- | ----- |
|type|`'pattern'`||
|pattern|`string`|Name of the pattern. Allowed values: `'gray125'`, `'darkGray'`, `'mediumGray'`, `'lightGray'`, `'gray0625'`, `'darkHorizontal'`, `'darkVertical'`, `'darkDown'`, `'darkUp'`, `'darkGrid'`, `'darkTrellis'`, `'lightHorizontal'`, `'lightVertical'`, `'lightDown'`, `'lightUp'`, `'lightGrid'`, `'lightTrellis'`.|
|foreground|<code>Color&#124;string&#124;number</code>|Color of the foreground. If string, will set an RGB color. If number, will set a theme color.|
|background|<code>Color&#124;string&#124;number</code>|Color of the background. If string, will set an RGB color. If number, will set a theme color.|

### NOTOC-GradientFill
An object representing a gradient fill.

|Property|Type|Description|
| ------------- | ------------- | ----- |
|type|`'gradient'`||
|[gradientType]|`string`|Type of gradient. Allowed values: `'linear'` (default), `'path'`. With a path gradient, a path is drawn between the top, left, right, and bottom values and a graident is draw from that path to the outside of the cell.|
|stops|`Array.<{}>`||
|stops[].position|`number`|The position of the stop from 0 to 1.|
|stops[].color|<code>Color&#124;string&#124;number</code>|Color of the stop. If string, will set an RGB color. If number, will set a theme color.|
|[angle]|`number`|If linear gradient, the angle of clockwise rotation of the gradient.|
|[left]|`number`|If path gradient, the left position of the path as a percentage from 0 to 1.|
|[right]|`number`|If path gradient, the right position of the path as a percentage from 0 to 1.|
|[top]|`number`|If path gradient, the top position of the path as a percentage from 0 to 1.|
|[bottom]|`number`|If path gradient, the bottom position of the path as a percentage from 0 to 1.|

## API Reference
<!-- api -->
