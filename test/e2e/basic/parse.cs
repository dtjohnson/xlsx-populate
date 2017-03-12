static class Parser {
    public static object Parse(Workbook workbook) {
        var sheet = (Worksheet)workbook.Worksheets[1];
        var cell = (Range)sheet.Cells[1, 1];
        var value = cell.Value2;

        return new {
            value = value
        };
    }
}
