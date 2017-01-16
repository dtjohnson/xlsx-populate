"use strict";

const Workbook = require('../../lib/Workbook');

// Load the input workbook from file.
Workbook.fromBlankAsync("./foo.xlsx")
    .then(workbook => {
        // // Modify the workbook.
        const sheet = workbook.sheet("Sheet1");

        sheet.cell("A1").value("bold").style().bold(true);
        sheet.cell("A2").value("italic").style().italic(true);
        sheet.cell("A3").value("underline").style().underline(true);
        sheet.cell("A4").value("strikethrough").style().strikethrough(true);
        sheet.cell("A5").value("superscript").style().superscript(true);
        sheet.cell("A6").value("subscript").style().subscript(true);
        sheet.cell("A7").value("larger").style().fontSize(14);
        sheet.cell("A8").value("smaller").style().fontSize(8);
        sheet.cell("A9").value("comic sans").style().fontFamily("Comic Sans MS");
        // sheet.cell("A10").value("left").style().horizontalAlignment("left");
        // sheet.cell("A11").value("center").style().horizontalAlignment("center");
        // sheet.cell("A12").value("right").style().horizontalAlignment("right");
        // sheet.cell("A13").value("top").style().verticalAlignment("top");
        // sheet.cell("A14").value("middle").style().verticalAlignment("center");
        // sheet.cell("A15").value("bottom").style().verticalAlignment("bottom");
        // sheet.cell("A16").value("wrapped text").style().wrappedText(true);
        // sheet.cell("A17").value("background color").style().fillBackgroundColor("ff0000");
        // sheet.cell("A18").value("font color").style().fontColor("ff0000");
        // sheet.cell("A19").value("top border").style().topBorderStyle("thin");
        // sheet.cell("A20").value("left border").style().leftBorderStyle("thin");
        // sheet.cell("A21").value("right border").style().rightBorderStyle("thin");
        // sheet.cell("A22").value("bottom border").style().bottomBorderStyle("thin");
        // sheet.cell("A23").value("double bottom border").style().bottomBorderStyle("double");
        // sheet.cell("A24").value("medium bottom border").style().bottomBorderStyle("medium");
        // sheet.cell("A25").value("thick bottom border").style().bottomBorderStyle("thick");
        // sheet.cell("A26").value("indent").style().indent(2);
        // sheet.cell("A27").value("text orientation").style().textOrientation(20);
        // sheet.cell("A28").value("angle counterclockwise").style().angleTextCounterclockwise(true);
        // sheet.cell("A29").value("angle clockwise").style().angleTextClockwise(true);
        // sheet.cell("A30").value("verticalText").style().verticalText(true);
        // sheet.cell("A31").value("angle counterclockwise").style().angleTextCounterclockwise(true);
        // sheet.cell("A32").value("rotate text up").style().rotateTextUp(true);
        // sheet.cell("A33").value("rotate text down").style().rotateTextDown(true);
        // sheet.cell("A34").value("number").relativeCell(0, 1).value(1.2).style().numberFormat(2);
        // sheet.cell("A35").value("currency").relativeCell(0, 1).value(1.2).style().numberFormat(`$#,##0.00`);
        // sheet.cell("A36").value("accounting").relativeCell(0, 1).value(1.2).style().numberFormat(`_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)`);
        // sheet.cell("A37").value("short date").relativeCell(0, 1).value(1.2).style().numberFormat(14);
        // sheet.cell("A38").value("long date").relativeCell(0, 1).value(1.2).style().numberFormat(`[$-x-sysdate]dddd, mmmm dd, yyyy`);
        // sheet.cell("A39").value("time").relativeCell(0, 1).value(1.2).style().numberFormat(`[$-x-systime]h:mm:ss AM/PM`);
        // sheet.cell("A40").value("percentage").relativeCell(0, 1).value(1.2).style().numberFormat(10);
        // sheet.cell("A41").value("fraction").relativeCell(0, 1).value(1.2).style().numberFormat(12);
        // sheet.cell("A42").value("scientific").relativeCell(0, 1).value(1.2).style().numberFormat(11);
        // sheet.cell("A43").value("text").relativeCell(0, 1).value(1.2).style().numberFormat(49);

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    });
