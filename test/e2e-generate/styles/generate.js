"use strict";

module.exports = workbook => {
    const sheet = workbook.sheet("Sheet1");

    sheet.row(1).height(30);
    sheet.column('A').width(15);
    sheet.column('B').width(25);

    sheet.cell("A1").value("bold").style("bold", true)
        .relativeCell(1, 0).value("italic").style("italic", true)
        .relativeCell(1, 0).value("underline").style("underline", true)
        .relativeCell(1, 0).value("double underline").style("underline", "double")
        .relativeCell(1, 0).value("strikethrough").style("strikethrough", true)
        .relativeCell(1, 0).value("superscript").style("superscript", true)
        .relativeCell(1, 0).value("subscript").style("subscript", true)
        .relativeCell(1, 0).value("larger").style("fontSize", 14)
        .relativeCell(1, 0).value("smaller").style("fontSize", 8)
        .relativeCell(1, 0).value("comic sans").style("fontFamily", "Comic Sans MS")
        .relativeCell(1, 0).value("rgb font color").style("fontColor", "ff0000")
        .relativeCell(1, 0).value("theme font color").style("fontColor", 4)
        .relativeCell(1, 0).value("tinted theme font color").style("fontColor", { theme: 4, tint: -0.5 })
        .relativeCell(1, 0).value("horizontal center").style("horizontalAlignment", "center")
        .relativeCell(1, 0).value("this text is justified distributed").style({ horizontalAlignment: "distributed", justifyLastLine: true })
        .relativeCell(1, 0).value("indent").style("indent", 2)
        .relativeCell(1, 0).value("vertical center").style("verticalAlignment", "center")
        .relativeCell(1, 0).value("this text is wrapped text").style("wrapText", true)
        .relativeCell(1, 0).value("this text is shrink to fit").style("shrinkToFit", true)
        .relativeCell(1, 0).value("right-to-left").style("textDirection", "right-to-left")
        .relativeCell(1, 0).value("text rotation").style("textRotation", -10)
        .relativeCell(1, 0).value("angle counterclockwise").style("angleTextCounterclockwise", true)
        .relativeCell(1, 0).value("angle clockwise").style("angleTextClockwise", true)
        .relativeCell(1, 0).value("rotate text up").style("rotateTextUp", true)
        .relativeCell(1, 0).value("rotate text down").style("rotateTextDown", true)
        .relativeCell(1, 0).value("vertical text").style("verticalText", true)
        .relativeCell(1, 0).value("rgb solid fill").style("fill", "ff0000")
        .relativeCell(1, 0).value("theme solid fill").style("fill", 5)
        .relativeCell(1, 0).value("tinted theme solid fill").style("fill", { theme: 5, tint: 0.25 })
        .relativeCell(1, 0).value("pattern fill").style("fill", {
            type: "pattern",
            pattern: "darkDown",
            foreground: "ff0000",
            background: {
                theme: 3,
                tint: 0.4
            }
        })
        .relativeCell(1, 0).value("linear gradient fill").style("fill", {
            type: "gradient",
            angle: 10,
            stops: [
                { position: 0, color: "ff0000" },
                { position: 0.5, color: "00ff00" },
                { position: 1, color: "0000ff" }
            ]
        })
        .relativeCell(1, 0).value("path gradient fill").style("fill", {
            type: "gradient",
            gradientType: "path",
            left: 0.1,
            right: 0.3,
            top: 0.5,
            bottom: 0.7,
            stops: [
                { position: 0, color: "ff0000" },
                { position: 1, color: "0000ff" }
            ]
        })
        .relativeCell(2, 0).value("thin border").style("border", true)
        .relativeCell(2, 0).value("thick border").style("border", "thick")
        .relativeCell(2, 0).value("double blue border").style("border", { style: "double", color: "0000ff" })
        .relativeCell(2, 0).value("right red border").style("rightBorder", true).style("borderColor", "ff0000")
        .relativeCell(2, 0).value("theme diagonal up border").style("diagonalBorder", { style: "dashed", color: { theme: 6, tint: -0.1 }, direction: "up" })
        .relativeCell(2, 0).value("various styles border").style("borderStyle", { top: "hair", right: "thin", bottom: "medium", left: "thick" })
        .relativeCell(2, 0).value("various colors border").style("border", "thick").style("borderColor", { top: "ff0000", right: "00ff00", bottom: "0000ff", left: "ffff00" })
        .relativeCell(2, 0).value("complex border").style("border", {
            top: true,
            right: "thick",
            bottom: { style: "dotted", color: "ff0000" },
            left: { style: "mediumDashed", color: 5 },
            diagonal: { style: "thick", color: "0000ff", direction: "both" }
        })
        .relativeCell(1, 0).value("number").relativeCell(0, 1).value(1.2).style("numberFormat", "0.00")
        .relativeCell(1, -1).value("currency").relativeCell(0, 1).value(1.2).style("numberFormat", `$#,##0.00`)
        .relativeCell(1, -1).value("accounting").relativeCell(0, 1).value(1.2).style("numberFormat", `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)`)
        .relativeCell(1, -1).value("short date").relativeCell(0, 1).value(1.2).style("numberFormat", "m/d/yyyy")
        .relativeCell(1, -1).value("long date").relativeCell(0, 1).value(1.2).style("numberFormat", `[$-x-sysdate]dddd, mmmm dd, yyyy`)
        .relativeCell(1, -1).value("time").relativeCell(0, 1).value(1.2).style("numberFormat", `[$-x-systime]h:mm:ss AM/PM`)
        .relativeCell(1, -1).value("percentage").relativeCell(0, 1).value(1.2).style("numberFormat", "0.00%")
        .relativeCell(1, -1).value("fraction").relativeCell(0, 1).value(1.2).style("numberFormat", "# ?/?")
        .relativeCell(1, -1).value("scientific").relativeCell(0, 1).value(1.2).style("numberFormat", "0.00E+00")
        .relativeCell(1, -1).value("text").relativeCell(0, 1).value(1.2).style("numberFormat", "@");

    const sheet2 = workbook.addSheet("Sheet2");
    sheet2.column("A").style("bold", true);
    sheet2.row(1).style("italic", true);

    const sheet3 = workbook.addSheet("Sheet3");
    sheet3.row(1).style("italic", true);
    sheet3.column("A").style("bold", true);
    sheet3.range("A1:B2").value("foo");

    const sheet4 = workbook.addSheet("Sheet4");
    sheet4.range("A1:B2").value("foo");
    sheet4.column("A").style("bold", true);
    sheet4.row(1).style("italic", true);
};
