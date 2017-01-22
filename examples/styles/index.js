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
        // console.log(JSON.stringify(sheet.cell("A5").value("superscript").style().superscript(true).superscript()));
        // console.log(JSON.stringify(sheet.cell("A6").value("subscript").style().subscript(true).subscript()));
        // console.log(JSON.stringify(sheet.cell("A7").value("larger").style().fontSize(14).fontSize()));
        // console.log(JSON.stringify(sheet.cell("A8").value("smaller").style().fontSize(8).fontSize()));
        // console.log(JSON.stringify(sheet.cell("A9").value("comic sans").style().fontFamily("Comic Sans MS").fontFamily()));
        // console.log(JSON.stringify(sheet.cell("A10").value("left").style().horizontalAlignment("left").horizontalAlignment()));
        // console.log(JSON.stringify(sheet.cell("A11").value("center").style().horizontalAlignment("center").horizontalAlignment()));
        // console.log(JSON.stringify(sheet.cell("A12").value("right").style().horizontalAlignment("right").horizontalAlignment()));
        // console.log(JSON.stringify(sheet.cell("A13").value("top").style().verticalAlignment("top").verticalAlignment()));
        // console.log(JSON.stringify(sheet.cell("A14").value("middle").style().verticalAlignment("center").verticalAlignment()));
        // console.log(JSON.stringify(sheet.cell("A15").value("bottom").style().verticalAlignment("bottom").verticalAlignment()));
        // console.log(JSON.stringify(sheet.cell("A16").value("this is wrapped text").style().wrappedText(true).wrappedText()));
        // // sheet.cell("A17").value("background color").style().fillBackgroundColor("ff0000");
        // console.log(JSON.stringify(sheet.cell("A18").value("rgb font color").style().fontColor("ff0000").fontColor()));
        // console.log(JSON.stringify(sheet.cell("A19").value("indexed font color").style().fontColor(1).fontColor()));
        // console.log(JSON.stringify(sheet.cell("A20").value("top border").style().topBorderStyle("thin").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A21").value("left border").style().leftBorderStyle("thin").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A22").value("right border").style().rightBorderStyle("thin").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A23").value("bottom border").style().bottomBorderStyle("thin").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A24").value("double bottom border").style().bottomBorderStyle("double").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A25").value("medium bottom border").style().bottomBorderStyle("medium").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A26").value("thick bottom border").style().bottomBorderStyle("thick").topBorderStyle()));
        // console.log(JSON.stringify(sheet.cell("A27").value("indent").style().indent(2).indent()));
        // console.log(JSON.stringify(sheet.cell("A28").value("text rotation").style().textRotation(20).textRotation()));
        // console.log(JSON.stringify(sheet.cell("A29").value("angle counterclockwise").style().angleTextCounterclockwise(true).angleTextCounterclockwise()));
        // console.log(JSON.stringify(sheet.cell("A30").value("angle clockwise").style().angleTextClockwise(true).angleTextClockwise()));
        // console.log(JSON.stringify(sheet.cell("A31").value("verticalText").style().verticalText(true).verticalText()));
        // console.log(JSON.stringify(sheet.cell("A32").value("rotate text up").style().rotateTextUp(true).rotateTextUp()));
        // console.log(JSON.stringify(sheet.cell("A33").value("rotate text down").style().rotateTextDown(true).rotateTextDown()));
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
