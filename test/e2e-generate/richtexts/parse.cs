static class Parser
{
    public static object Parse(Workbook workbook)
    {
        var sheet = (Worksheet)workbook.Worksheets[1];
        var rtCell = sheet.Range["A1"];
        var rt1 = rtCell.Characters[1, 5]; // hello
        var rt2 = rtCell.Characters[6, 4]; // test
        var rt3 = rtCell.Characters[10, 4]; // 123\n
        var rt4 = rtCell.Characters[14, 4]; // 456\n
        var rt5 = rtCell.Characters[18, 4]; // 789\n
        var rt6 = rtCell.Characters[22, 8]; // 10\n11\n12

        return new {
            text = rtCell.Text,
            length = rtCell.Characters.Count,
            rt1Text = rt1.Text,
            rt2Text = rt2.Text,
            rt3Text = rt3.Text,
            rt4Text = rt4.Text,
            rt5Text = rt5.Text,
            rt6Text = rt6.Text,
            rt2Bold = rt2.Font.Bold,
            rt2FontFamily = rt2.Font.Name,
            rt3Italic = rt3.Font.Italic,
            // RGB = Red + (Green*256) + (Blue*256*256), which is 255+256+256*256 = 66047
            rt3FontColor = rt3.Font.Color,
            rt4Underline = rt4.Font.Underline,
            rt5Strikethrough = rt5.Font.Strikethrough,
            rt6Subscript = rt6.Font.Subscript,

            // https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlunderlinestyle?view=excel-pia
            rt6Underline = rt6.Font.Underline,
        };
    }
}
