* Cell.value getter
* Shared string support?
* Search for value
* Active cell in sheet

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