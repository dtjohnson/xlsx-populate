Backlog of features to implement (in no particular order):

* Selected cells in sheet (worksheet->sheetViews->sheetView->selection->sqref). Should wait for Group.
* Support rich text (in cell values and shared strings)
* Sheet: Rename, delete. Will require manipulating/clearing formulas.
* When clearing a shared formula ref cell, we should move the shared formula ref to another.
* Returning a shared formula in a not ref cell returns "SHARED". We should return a translated formula.
* ColumnRange, RowRange, Group
* Support sheet, row, and column styles. That will require iterating over all contained styles and updating.
* Conditional formatting
* Print settings
* Formula parsing
* Charts
* Cell comments. Will require rich text parsing. A comments relationship is creating in the sheet rels file that points to a commentsN.xml file.
* Cell protection
* Copy style
* Built-in styles
* Named styles
* Insert images
* Frozen rows/columns
* Create defined name
* Enum of standard number formats?
* Drawings
* Array formulas - May want to introduce a new Formula class.
