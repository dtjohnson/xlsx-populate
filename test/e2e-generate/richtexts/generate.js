"use strict";
const RichText = require("../../../lib/XlsxPopulate").RichText;

module.exports = workbook => {
    const sheet = workbook.sheet(0);
    const cell = sheet.cell('A1');
    const rt = new RichText();
    rt.add('test', { bold: true, fontFamily: 'Arial' })
        .add('123\n', { italic: true, fontColor: 'FF0101' })
        .add('456\r', { underline: true })
        .add('789\r\n', { strikethrough: true })
        .add('10\n11\r12', { subscript: true, underline: 'double' });
    rt.add('hello');
    rt.remove(5);
    rt.add('hello', {}, 0);
    cell.value(rt);
};
