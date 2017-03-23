static class Parser
{
    public static object Parse(Workbook workbook)
    {
        var results = new List<object>();

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            results.Add(new {
                name = sheet.Name,
                active = sheet == workbook.ActiveSheet,
                color = sheet.Tab.ThemeColor,
            });
        }

        return results;
    }
}
