Backlog of features to implement (in no particular order):

* Selected cells in sheet (worksheet->sheetViews->sheetView->selection->sqref)
* Support rich text (in cell values and shared strings)
* Sheet: Rename, delete. Will require manipulating/clearing formulas.
* When clearing a shared formula ref cell, we should move the shared formula ref to another.
* Returning a shared formula in a not ref cell returns "SHARED". We should return a translated formula.
* ColumnRange, RowRange, Group
* Column/Row styles
* Conditional formatting
* Print settings
* Autofilters
* Data validation
* Formula parsing
* Charts
* Cell comments. Will require rich text parsing. A comments relationship is creating in the sheet rels file that points to a commentsN.xml file.
* Cell protection
* Copy style
* Built-in styles
* Named styles
* Insert images
* Frozen rows/columns
* Workbook metadata (like author)
* Create defined name
* Enum of standard number formats?