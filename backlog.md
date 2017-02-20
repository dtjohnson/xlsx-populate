Backlog of features to implement (in no particular order):

* Selected/active cell in sheet

```
<worksheet>
    ...
	<sheetViews>
		<sheetView tabSelected="1" workbookViewId="0">
			<selection activeCell="D4" sqref="D4"/>
		</sheetView>
	</sheetViews>
	...
</worksheet>
```

* Active sheet in workbook

```
<workbook>
    ...
    <bookViews>
		<workbookView xWindow="4650" yWindow="0" windowWidth="27870" windowHeight="12795" activeTab="1"/>
	</bookViews>
    ...
</workbook>
```

* Support rich text (in cell values and shared strings)
* Sheet: Rename, move, delete, activate. May require manipulating/clearing formulas.
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
* Cell comments
* Cell protection
* Copy style
* Built-in styles
* Named styles
* Insert images
* Frozen rows/columns
* Workbook metadata (like author)
* Create defined name
* Enum of standard number formats?
