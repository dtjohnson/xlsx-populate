static class Parser
{
    public static object Parse(Workbook workbook)
    {
        return new {
            editedValues = ParseEditedValues(workbook),
            values = ParseValues(workbook),
            richText = ParseRichText(workbook),
            hyperlinks = ParseHyperlinks(workbook),
            comments = ParseComments(workbook),
            mergeCells = ParseMergeCells(workbook),
            formulas = ParseFormulas(workbook),
            definedNames = ParseDefinedNames(workbook),
            styles = ParseStyles(workbook),
            conditionalFormatting = ParseConditionalFormatting(workbook),
            dataValidation = ParseDataValidation(workbook),
            filter = ParseFilter(workbook),
            table = ParseTable(workbook),
            pivotTable = ParsePivotTable(workbook),
            image = ParseImage(workbook),
            chart = ParseChart(workbook),
        };
    }

    private static object ParseEditedValues(Workbook workbook)
    {
        var edits = new Dictionary<string, object>();

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var range = sheet.Range["A1"];
            edits.Add(sheet.Name, range.Value2);
        }

        return edits;
    }

    private static object ParseValues(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Values"];

        var editedRange = sheet.Range["A1"];
        var textRange = sheet.Range["A2"];
        var numberRange = sheet.Range["A3"];
        var booleanRange = sheet.Range["A4"];
        var dateRange = sheet.Range["A5"];

        return new {
            edited = editedRange.Value2,
            text = new {
                value = textRange.Value2,
                text = textRange.Text,
                format = textRange.NumberFormat,
            },
            number = new {
                value = numberRange.Value2,
                text = numberRange.Text,
                format = numberRange.NumberFormat,
            },
            boolean = new {
                value = booleanRange.Value2,
                text = booleanRange.Text,
                format = booleanRange.NumberFormat,
            },
            date = new {
                value = dateRange.Value2,
                text = dateRange.Text,
                format = dateRange.NumberFormat,
            },
        };
    }

    private static object ParseRichText(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Rich Text"];
        var range = sheet.Range["A2"];

        var boldChars = range.get_Characters(1, 3);
        var italicChars = range.get_Characters(5, 3);

        return new {
            bold = new {
                text = boldChars.Text,
                bold = boldChars.Font.Bold,
                italic = boldChars.Font.Italic,
            },
            italic = new {
                text = italicChars.Text,
                bold = italicChars.Font.Bold,
                italic = italicChars.Font.Italic,
            },
        };
    }

    private static object ParseHyperlinks(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Hyperlinks"];
        var range = sheet.Range["A2"];

        return new {
            text = range.Value2,
            address = range.Hyperlinks[1].Address,
        };
    }

    private static object ParseComments(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Comments"];
        var hiddenComment = sheet.Range["A2"].Comment;
        var visibleComment = sheet.Range["A3"].Comment;

        return new {
            hidden = new {
                text = hiddenComment.Text(),
                visible = hiddenComment.Visible,
                topLeftCell = hiddenComment.Shape.TopLeftCell.Address[false, false],
                top = hiddenComment.Shape.Top,
                left = hiddenComment.Shape.Left,
                height = hiddenComment.Shape.Height,
                width = hiddenComment.Shape.Width,
            },
            visible = new {
                text = visibleComment.Text(),
                visible = visibleComment.Visible,
                topLeftCell = visibleComment.Shape.TopLeftCell.Address[false, false],
                top = visibleComment.Shape.Top,
                left = visibleComment.Shape.Left,
                height = visibleComment.Shape.Height,
                width = visibleComment.Shape.Width,
            },
        };
    }

    private static object ParseMergeCells(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Merge Cells"];

        return new[] {
            sheet.Range["A2"].MergeArea.Address[false, false],
            sheet.Range["A4"].MergeArea.Address[false, false],
            sheet.Range["A7"].MergeArea.Address[false, false],
        };
    }

    private static object ParseFormulas(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Formulas"];

        var normalRange = sheet.Range["A2"];
        var arrayRange = sheet.Range["A3"];

        // There doesn't seem to be a means of distinguishing an array formula from a normal one.
        // Both Range.Formula and Range.FormulaArray return the same value in both cases.
        return new {
            normal = new {
                value = normalRange.Value2,
                formula = normalRange.Formula,
            },
            array = new {
                value = arrayRange.Value2,
                formula = arrayRange.Formula,
            },
        };
    }

    private static object ParseDefinedNames(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Defined Names"];
        var otherSheet = (Worksheet)workbook.Sheets["Hyperlinks"];

        return new {
            workbookNameFromWorkbook = GetDefinedNameValue(workbook, "WORKBOOK_NAME"),
            workbookNameFromSheet = GetDefinedNameValue(sheet, "WORKBOOK_NAME"),
            workbookNameFromOtherSheet = GetDefinedNameValue(otherSheet, "WORKBOOK_NAME"),
            sheetNameFromWorkbook = GetDefinedNameValue(workbook, "SHEET_NAME"),
            sheetNameFromSheet = GetDefinedNameValue(sheet, "SHEET_NAME"),
            sheetNameFromOtherSheet = GetDefinedNameValue(otherSheet, "SHEET_NAME"),
        };
    }

    private static object GetDefinedNameValue(dynamic parent, string name)
    {
        object value;

        try
        {
        value = "foo";
            var definedName = parent.Names[name];
            var range = definedName.RefersToRange;
            value = range.Value2;
        }
        catch
        {
            value = null;
        }

        return value;
    }

    private static object ParseStyles(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Styles"];

        var populatedCell = sheet.Range["B2"];
        var emptyCell = sheet.Range["C3"];
        var goodCell = sheet.Range["A4"];
        var badCell = sheet.Range["B4"];
        var numberFormatCell = sheet.Range["A6"];
        var row = (Range)sheet.Rows[3];
        var column = (Range)sheet.Columns["C"];

        return new {
            populatedCell = new {
                bold = populatedCell.Font.Bold,
                italic = populatedCell.Font.Italic,
                value = populatedCell.Value2,
            },
            emptyCell = new {
                bold = emptyCell.Font.Bold,
                italic = emptyCell.Font.Italic,
                value = emptyCell.Value2,
            },
            row = new {
                bold = row.Font.Bold,
                italic = row.Font.Italic
            },
            column = new {
                bold = column.Font.Bold,
                italic = column.Font.Italic
            },
            good = ((Style)goodCell.Style).Name,
            bad = ((Style)badCell.Style).Name,
            numberFormat = numberFormatCell.NumberFormat,
        };
    }

    private static object ParseConditionalFormatting(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Conditional Formatting"];

        // This will throw if the conditional format doesn't exist or is not a color scale.
        var range = sheet.Range["A2", "F2"];
        var colorScale = (ColorScale)range.FormatConditions[1];
        return "color scale";
    }

    private static object ParseDataValidation(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Data Validation"];

        var range = sheet.Range["A2"];
        var validation = range.Validation;

        return new {
            type = (XlDVType)validation.Type,
            source = validation.Formula1,
            ignoreBlank = validation.IgnoreBlank,
            inCellDropdown = validation.InCellDropdown,
        };
    }

    private static object ParseFilter(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Filter"];

        var autoFilter = sheet.AutoFilter;
        var filter = autoFilter.Filters[1];

        return new {
            range = autoFilter.Range.Address[false, false],
            count = filter.Count,
            on = filter.On,
            op = (XlAutoFilterOperator)filter.Operator,
            criteria = filter.Criteria1,
        };
    }

    private static object ParseTable(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Table"];

        var table = sheet.ListObjects["Table1"];
        var tableStyle = (TableStyle)table.TableStyle;

        return new {
            range = table.Range.Address[false, false],
            style = tableStyle.Name,
        };
    }

    private static object ParsePivotTable(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Pivot Table"];

        var pivotTable = (PivotTable)sheet.PivotTables("PivotTable1");

        return new {
            sourceData = pivotTable.SourceData,
            tableRange = pivotTable.TableRange1.Address[false, false],
        };
    }

    private static object ParseImage(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Image"];

        var picture = (Picture)sheet.Pictures(1);

        return new {
            topLeftCell = picture.TopLeftCell.Address[false, false],
            top = picture.Top,
            left = picture.Left,
            height = picture.Height,
            width = picture.Width,
        };
    }

    private static object ParseChart(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Sheets["Chart"];

        var chartObj = (ChartObject)sheet.ChartObjects(1);
        var chart = chartObj.Chart;

        return new {
            type = chart.ChartType,
            topLeftCell = chartObj.TopLeftCell.Address[false, false],
            top = chartObj.Top,
            left = chartObj.Left,
            height = chartObj.Height,
            width = chartObj.Width,
        };
    }
}
