"use strict";

module.exports = workbook => {
    const cell = workbook.sheet(0).cell('A2');
    cell.value().add('123', { fontColor: '0000ff' });

    // create new rich text
    workbook.sheet(0).cell('B1').toRichText()
        .add('test', { bold: true, italic: true })
        .add('123', { italic: true, fontColor: '123456' });

    return cell.value().get(1).style('fontFamily');
};
