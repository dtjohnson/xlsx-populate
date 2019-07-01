static class Parser
{
    public static object Parse(Workbook workbook)
    {
        var sheet1 = (Worksheet)workbook.Worksheets[1];
        var sheet2 = (Worksheet)workbook.Worksheets[2];
        var sheet3 = (Worksheet)workbook.Worksheets[3];
        var sheet4 = (Worksheet)workbook.Worksheets[4];
        var sheet5 = (Worksheet)workbook.Worksheets[5];
        var sheet6 = (Worksheet)workbook.Worksheets[6];
        var sheet7 = (Worksheet)workbook.Worksheets[7];

        var activeWindow = sheet1.Application.ActiveWindow;

        sheet1.Activate();
        var result1 = new {
            SplitRow = activeWindow.SplitRow,
            SplitColumn = activeWindow.SplitColumn,
            FreezePanes = activeWindow.FreezePanes,
            Split = activeWindow.Split,
        };
        sheet2.Activate();
        var result2 = new {
            SplitRow = activeWindow.SplitRow,
            SplitColumn = activeWindow.SplitColumn,
            FreezePanes = activeWindow.FreezePanes,
            Split = activeWindow.Split,
        };
        sheet3.Activate();
        var result3 = new {
            SplitRow = activeWindow.SplitRow,
            SplitColumn = activeWindow.SplitColumn,
            FreezePanes = activeWindow.FreezePanes,
            Split = activeWindow.Split,
        };
        sheet4.Activate();
        var result4 = new {
            SplitRow = activeWindow.SplitRow,
            SplitColumn = activeWindow.SplitColumn,
            FreezePanes = activeWindow.FreezePanes,
            Split = activeWindow.Split,
        };
        sheet5.Activate();
        var result5 = new {
            SplitRow = activeWindow.SplitRow,
            SplitColumn = activeWindow.SplitColumn,
            FreezePanes = activeWindow.FreezePanes,
            Split = activeWindow.Split,
        };
        sheet6.Activate();
        var result6 = new {
            SplitRow = activeWindow.SplitRow,
            SplitColumn = activeWindow.SplitColumn,
            FreezePanes = activeWindow.FreezePanes,
            Split = activeWindow.Split,
        };
        sheet7.Activate();
        var result7 = new {
            SplitRow = activeWindow.SplitRow,
            SplitColumn = activeWindow.SplitColumn,
            FreezePanes = activeWindow.FreezePanes,
            Split = activeWindow.Split,
        };
        return new {
            sheet1 = result1,
            sheet2 = result2,
            sheet3 = result3,
            sheet4 = result4,
            sheet5 = result5,
            sheet6 = result6,
            sheet7 = result7
        };
    }
}
