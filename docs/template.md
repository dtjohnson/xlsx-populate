[![view on npm](http://img.shields.io/npm/v/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![npm module downloads per month](http://img.shields.io/npm/dm/xlsx-populate.svg)](https://www.npmjs.org/package/xlsx-populate)
[![Build Status](https://travis-ci.org/dtjohnson/xlsx-populate.svg?branch=master)](https://travis-ci.org/dtjohnson/xlsx-populate)
[![Dependency Status](https://david-dm.org/dtjohnson/xlsx-populate.svg)](https://david-dm.org/dtjohnson/xlsx-populate)

# xlsx-populate
TODO

## Table of Contents
<!-- toc -->

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
<!-- api -->