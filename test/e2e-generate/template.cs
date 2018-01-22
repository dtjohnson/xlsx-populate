using System;
using System.Drawing;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

public class Startup
{
    public async Task<object> Invoke(dynamic input)
    {
        var app = new Application();

        try
        {
            app.Visible = false;
            var workbook = app.Workbooks.Open(input.path, Password: input.password);
            return Parser.Parse(workbook);
        }
        finally
        {
            app.DisplayAlerts = false;
            app.Quit();
        }
    }
}
