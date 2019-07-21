static class Parser {
    public static object Parse(Workbook workbook) {
        var sheet = (Worksheet)workbook.Worksheets[2];
        var range = sheet.Range["A1", "A3"];
        var values = range.Value2;

        return new {
            values = values
        };
    }
}
