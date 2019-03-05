static class Parser
{
    public static object Parse(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Worksheets[1];
        // https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.pagesetup?view=excel-pia
        var pageSetup = sheet.PageSetup;

        return new {
            // 1 inch = 72 pt
            PrintGridlines = pageSetup.PrintGridlines,
            PrintHeadings = pageSetup.PrintHeadings,
            CenterHorizontally = pageSetup.CenterHorizontally,
            CenterVertically = pageSetup.CenterVertically,
        };
    }
}
