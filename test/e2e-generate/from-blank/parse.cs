static class Parser
{
    public static object Parse(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Worksheets[1];
        var valueCell = sheet.Range["A1"];
        var formulaCell = sheet.Range["A2"];

        return new {
            value = valueCell.Value2,
            formula = formulaCell.Formula,
            formulaValue = formulaCell.Value2,
        };
    }
}
