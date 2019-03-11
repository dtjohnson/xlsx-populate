static class Parser
{
    public static object Parse(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Worksheets[1];

        return new {
            SplitRow = sheet.Application.ActiveWindow.SplitRow,
            SplitColumn = sheet.Application.ActiveWindow.SplitColumn,
            FreezePanes = sheet.Application.ActiveWindow.FreezePanes,
        };
    }
}
