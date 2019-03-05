static class Parser
{
    public static object Parse(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Worksheets[1];
        // https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.pagesetup?view=excel-pia
        var pageSetup = sheet.PageSetup;

        return new {
            // 1 inch = 72 pt
            left = pageSetup.LeftMargin / 72,
            right = pageSetup.RightMargin / 72,
            top = pageSetup.TopMargin / 72,
            bottom = pageSetup.BottomMargin / 72,
            header = pageSetup.HeaderMargin / 72,
            footer = pageSetup.FooterMargin / 72,
        };
    }
}
