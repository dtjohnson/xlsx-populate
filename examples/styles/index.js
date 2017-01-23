"use strict";

const Workbook = require('../../lib/Workbook');

// Load the input workbook from file.
Workbook.fromBlankAsync("./row col size.xlsx")
    .then(workbook => {
        // // Modify the workbook.
        const sheet = workbook.sheet("Sheet1");

        console.log(JSON.stringify(sheet.column('A').width(15).width()));

        console.log(JSON.stringify(sheet.cell("A1").value("bold").style("bold", true).style("bold")));
        console.log(JSON.stringify(sheet.cell("A2").value("italic").style("italic", true).style("italic")));
        console.log(JSON.stringify(sheet.cell("A3").value("underline").style("underline", true).style("underline")));
        console.log(JSON.stringify(sheet.cell("A4").value("strikethrough").style("strikethrough", true).style("strikethrough")));
        console.log(JSON.stringify(sheet.cell("A5").value("font vertical alignment").style("fontVerticalAlignment", "superscript").style("fontVerticalAlignment")));
        console.log(JSON.stringify(sheet.cell("A6").value("superscript").style("superscript", true).style("superscript")));
        console.log(JSON.stringify(sheet.cell("A7").value("subscript").style("subscript", true).style("subscript")));
        console.log(JSON.stringify(sheet.cell("A8").value("larger").style("fontSize", 14).style("fontSize")));
        console.log(JSON.stringify(sheet.cell("A9").value("smaller").style("fontSize", 8).style("fontSize")));
        console.log(JSON.stringify(sheet.cell("A10").value("comic sans").style("fontFamily", "Comic Sans MS").style("fontFamily")));
        console.log(JSON.stringify(sheet.cell("A11").value("left").style("horizontalAlignment", "left").style("horizontalAlignment", null).style("horizontalAlignment")));
        console.log(JSON.stringify(sheet.cell("A12").value("center").style("horizontalAlignment", "center").style("horizontalAlignment")));
        console.log(JSON.stringify(sheet.cell("A13").value("right").style("horizontalAlignment", "right").style("horizontalAlignment")));
        console.log(JSON.stringify(sheet.cell("A14").value("top").style("verticalAlignment", "top").style("verticalAlignment")));
        console.log(JSON.stringify(sheet.cell("A15").value("middle").style("verticalAlignment", "center").style("verticalAlignment")));
        console.log(JSON.stringify(sheet.cell("A16").value("bottom").style("verticalAlignment", "bottom").style("verticalAlignment")));
        console.log(JSON.stringify(sheet.cell("A17").value("this is wrapped text").style("wrappedText", true).style("wrappedText")));
        // // sheet.cell("A18").value("background color").style().fillBackgroundColor("ff0000");
        console.log(JSON.stringify(sheet.cell("A19").value("rgb font color").style("fontColor", "ff0000").style("fontColor")));
        console.log(JSON.stringify(sheet.cell("A20").value("indexed font color").style("fontColor", 4).style("fontColor")));
        // console.log(JSON.stringify(sheet.cell("A21").value("top border").style().topBorderStyle("thin").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A22").value("left border").style().leftBorderStyle("thin").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A23").value("right border").style().rightBorderStyle("thin").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A24").value("bottom border").style().bottomBorderStyle("thin").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A25").value("double bottom border").style().bottomBorderStyle("double").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A26").value("medium bottom border").style().bottomBorderStyle("medium").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A27").value("thick bottom border").style().bottomBorderStyle("thick").topBorderStyle()));
        console.log(JSON.stringify(sheet.cell("A28").value("indent").style("indent", 2).style("indent")));
        console.log(JSON.stringify(sheet.cell("A29").value("text rotation").style("textRotation", 20).style("textRotation")));
        console.log(JSON.stringify(sheet.cell("A30").value("angle counterclockwise").style("angleTextCounterclockwise", true).style("angleTextCounterclockwise")));
        console.log(JSON.stringify(sheet.cell("A31").value("angle clockwise").style("angleTextClockwise", true).style("angleTextClockwise")));
        console.log(JSON.stringify(sheet.cell("A32").value("verticalText").style("verticalText", true).style("verticalText")));
        console.log(JSON.stringify(sheet.cell("A33").value("rotate text up").style("rotateTextUp", true).style("rotateTextUp")));
        console.log(JSON.stringify(sheet.cell("A34").value("rotate text down").style("rotateTextDown", true).style("rotateTextDown")));
        // sheet.cell("A35").value("number").relativeCell(0, 1).value(1.2).style().numberFormat(2);
        // sheet.cell("A36").value("currency").relativeCell(0, 1).value(1.2).style().numberFormat(`$#,##0.00`);
        // sheet.cell("A37").value("accounting").relativeCell(0, 1).value(1.2).style().numberFormat(`_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)`);
        // sheet.cell("A38").value("short date").relativeCell(0, 1).value(1.2).style().numberFormat(14);
        // sheet.cell("A39").value("long date").relativeCell(0, 1).value(1.2).style().numberFormat(`[$-x-sysdate]dddd, mmmm dd, yyyy`);
        // sheet.cell("A40").value("time").relativeCell(0, 1).value(1.2).style().numberFormat(`[$-x-systime]h:mm:ss AM/PM`);
        // sheet.cell("A41").value("percentage").relativeCell(0, 1).value(1.2).style().numberFormat(10);
        // sheet.cell("A42").value("fraction").relativeCell(0, 1).value(1.2).style().numberFormat(12);
        // sheet.cell("A43").value("scientific").relativeCell(0, 1).value(1.2).style().numberFormat(11);
        // sheet.cell("A44").value("text").relativeCell(0, 1).value(1.2).style().numberFormat(49);

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    });
