using System;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

public class Startup
{
    public async Task<object> Invoke(object input)
    {
        int v = (int)input;
        return CreateExcelWorksheet.Foo();
    }
}

static class CreateExcelWorksheet
{
    public static object Foo()
    {
        var app = new Application();
        app.Visible = true;

        var workbook = app.Workbooks.Open(@"C:\Users\djohnson\Code\GitHub\xlsx-populate\test\e2e\out.xlsx");
        var sheet = (Worksheet)workbook.Worksheets[1];
        var cell = (Range)sheet.Cells[1, 1];
        var value = cell.Value2;

        app.DisplayAlerts = false;
        app.Quit();

        return new {
            value = value
        };
    }
}