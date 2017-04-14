static class Parser
{
    public static object Parse(Workbook workbook)
    {
        var results = new Dictionary<object, object>();
        var sheet = (Worksheet)workbook.Worksheets[1];

        var cell = sheet.Range["A1"];
        results[cell.Value2] = cell.Font.Bold;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Font.Italic;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = (XlUnderlineStyle)cell.Font.Underline;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = (XlUnderlineStyle)cell.Font.Underline;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Font.Strikethrough;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Font.Superscript;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Font.Subscript;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Font.Size;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Font.Size;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Font.Name;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = OleToHex(cell.Font.Color);

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Font.ThemeColor;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = new {
            theme = cell.Font.ThemeColor,
            tint = Math.Round((double)cell.Font.TintAndShade, 4), // Round to avoid floating point comparison issues.
        };

        cell = cell.Offset[1, 0];
        results[cell.Value2] = (XlHAlign)cell.HorizontalAlignment;

        // Justify Last Line / Justify Distributed doesn't seem accessible in the Interop API...
        cell = cell.Offset[1, 0];
        results[cell.Value2] = (XlHAlign)cell.HorizontalAlignment;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.IndentLevel;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = (XlVAlign)cell.VerticalAlignment;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.WrapText;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.ShrinkToFit;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = (Constants)cell.ReadingOrder;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Orientation;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Orientation;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Orientation;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = (XlOrientation)cell.Orientation;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = (XlOrientation)cell.Orientation;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = (XlOrientation)cell.Orientation;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = new {
            pattern = (XlPattern)cell.Interior.Pattern,
            rgb = OleToHex(cell.Interior.Color),
        };

        cell = cell.Offset[1, 0];
        results[cell.Value2] = new {
            pattern = (XlPattern)cell.Interior.Pattern,
            theme = cell.Interior.ThemeColor,
        };

        cell = cell.Offset[1, 0];
        results[cell.Value2] = new {
            pattern = (XlPattern)cell.Interior.Pattern,
            theme = cell.Interior.ThemeColor,
            tint = Math.Round((double)cell.Interior.TintAndShade, 4),
        };

        cell = cell.Offset[1, 0];
        results[cell.Value2] = new {
            pattern = (XlPattern)cell.Interior.Pattern,
            fgRgb = OleToHex(cell.Interior.PatternColor),
            bgTheme = cell.Interior.ThemeColor,
            bgTint = Math.Round((double)cell.Interior.TintAndShade, 4),
        };

        cell = cell.Offset[1, 0];
        var linearGradient = (LinearGradient)cell.Interior.Gradient;
        results[cell.Value2] = new {
            pattern = (XlPattern)cell.Interior.Pattern,
            degree = linearGradient.Degree,
            stops = new object[] {
                new {
                    position = linearGradient.ColorStops[1].Position,
                    rgb = OleToHex(linearGradient.ColorStops[1].Color),
                },
                new {
                    position = linearGradient.ColorStops[2].Position,
                    rgb = OleToHex(linearGradient.ColorStops[2].Color),
                },
                new {
                    position = linearGradient.ColorStops[3].Position,
                    rgb = OleToHex(linearGradient.ColorStops[3].Color),
                },
            }
        };

        cell = cell.Offset[1, 0];
        var rectangularGradient = (RectangularGradient)cell.Interior.Gradient;
        results[cell.Value2] = new {
            pattern = (XlPattern)cell.Interior.Pattern,
            left = rectangularGradient.RectangleLeft,
            right = rectangularGradient.RectangleRight,
            top = rectangularGradient.RectangleTop,
            bottom = rectangularGradient.RectangleBottom,
            stops = new object[] {
                new {
                    position = rectangularGradient.ColorStops[1].Position,
                    rgb = OleToHex(rectangularGradient.ColorStops[1].Color),
                },
                new {
                    position = rectangularGradient.ColorStops[2].Position,
                    rgb = OleToHex(rectangularGradient.ColorStops[2].Color),
                },
            }
        };

        cell = cell.Offset[2, 0];
        results[cell.Value2] = GetBorders(cell);

        cell = cell.Offset[2, 0];
        results[cell.Value2] = GetBorders(cell);

        cell = cell.Offset[2, 0];
        results[cell.Value2] = GetBorders(cell);

        cell = cell.Offset[2, 0];
        results[cell.Value2] = GetBorders(cell);

        cell = cell.Offset[2, 0];
        results[cell.Value2] = GetBorders(cell);

        cell = cell.Offset[2, 0];
        results[cell.Value2] = GetBorders(cell);

        cell = cell.Offset[2, 0];
        results[cell.Value2] = GetBorders(cell);

        cell = cell.Offset[2, 0];
        results[cell.Value2] = GetBorders(cell);

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Offset[0, 1].NumberFormat;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Offset[0, 1].NumberFormat;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Offset[0, 1].NumberFormat;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Offset[0, 1].NumberFormat;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Offset[0, 1].NumberFormat;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Offset[0, 1].NumberFormat;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Offset[0, 1].NumberFormat;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Offset[0, 1].NumberFormat;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Offset[0, 1].NumberFormat;

        cell = cell.Offset[1, 0];
        results[cell.Value2] = cell.Offset[0, 1].NumberFormat;

        results["column-row"] = GetRowColumnStyleData(workbook, 2);
        results["row-column-value"] = GetRowColumnStyleData(workbook, 3);
        results["value-column-row"] = GetRowColumnStyleData(workbook, 4);

        return results;
    }

    private static object GetRowColumnStyleData(Workbook workbook, int sheetIndex)
    {
        var sheet = (Worksheet)workbook.Worksheets[sheetIndex];
        var column = (Range)sheet.Columns[1];
        var row = (Range)sheet.Rows[1];

        return new {
            column = new {
                bold = column.Font.Bold,
                italic = column.Font.Italic,
            },
            row = new {
                bold = row.Font.Bold,
                italic = row.Font.Italic,
            },
            A1 = GetCellStyleData(sheet, "A1"),
            B1 = GetCellStyleData(sheet, "B1"),
            A2 = GetCellStyleData(sheet, "A2"),
            B2 = GetCellStyleData(sheet, "B2"),
        };
    }

    private static object GetCellStyleData(Worksheet sheet, string address)
    {
        var cell = (Range)sheet.Range[address];
        return new {
            bold = cell.Font.Bold,
            italic = cell.Font.Italic,
            value = cell.Value2,
        };
    }

    private static string OleToHex(object ole)
    {
        var c = ColorTranslator.FromOle(Convert.ToInt32(ole));
        return c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
    }

    private static object GetBorders(Range cell)
    {
        return new {
            left = GetBorder(cell, XlBordersIndex.xlEdgeLeft),
            right = GetBorder(cell, XlBordersIndex.xlEdgeRight),
            top = GetBorder(cell, XlBordersIndex.xlEdgeTop),
            bottom = GetBorder(cell, XlBordersIndex.xlEdgeBottom),
            diagonalUp = GetBorder(cell, XlBordersIndex.xlDiagonalUp),
            diagonalDown = GetBorder(cell, XlBordersIndex.xlDiagonalDown),
        };
    }

    private static object GetBorder(Range cell, XlBordersIndex index)
    {
        var border = cell.Borders[index];
        var result = new Dictionary<string, object>();

        var style = (XlLineStyle)border.LineStyle;
        result["style"] = style;
        if (style == XlLineStyle.xlLineStyleNone) return result;

        result["tint"] = Math.Round((double)border.TintAndShade, 4);
        result["weight"] = (XlBorderWeight)border.Weight;

        try
        {
            result["theme"] = border.ThemeColor;
        }
        catch
        {
            result["rgb"] = OleToHex(border.Color);
        }

        return result;
    }
}
