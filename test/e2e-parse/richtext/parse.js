"use strict";

module.exports = workbook => {
    const cell = workbook.sheet(0).cell('A2');
    cell.value().add('123', { fontColor: '0000ff' });

    // create new rich text
    const b1 = workbook.sheet(0).cell('B1');
    b1.clear().richText()
        .add('test', { bold: true, italic: true })

        // support all line separators works
        .add('123\n', { italic: true, fontColor: '123456' })
        .add('456\r', { italic: true, fontColor: '654321' })
        .add('789\r\n', { italic: true, fontColor: 'ff0000' })
        .add('10\n11\r12', { italic: true, fontColor: '00ff00' });

    // remember to set height to show whole rows
    workbook.sheet(0).row(1).height(100);
    // workbook.toFileAsync('./out.xlsx');
    return cell.value().get(1).style('fontFamily');
};
