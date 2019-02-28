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
        .add('456\r', { underline: true, fontColor: '654321' })
        .add('789\r\n', { strikethrough: true, fontColor: 'ff0000' })
        .add('10\n11\r12', { subscript: true, fontColor: '00ff00' })
        .add('10\n11\r12', {
            superscript: true, fontSize: 20, fontFamily: 'Arial', fontGenericFamily: 1,
            fontScheme: 'major'
        });

    // remember to set height to show whole rows
    workbook.sheet(0).row(1).height(100);
    return [
        cell.value().get(1).style(['fontFamily', 'fontColor']),
        b1.richText().get(0).value(),
        b1.richText().get(0).style(['bold', 'italic']),
        b1.richText().get(1).value(),
        b1.richText().get(1).style(['italic', 'fontColor']),
        b1.richText().get(2).value(),
        b1.richText().get(2).style(['underline', 'fontColor']),
        b1.richText().get(3).value(),
        b1.richText().get(3).style(['strikethrough', 'fontColor']),
        b1.richText().get(4).value(),
        b1.richText().get(4).style(['subscript', 'fontColor']),
        b1.richText().get(5).style(['superscript', 'fontSize', 'fontFamily', 'fontGenericFamily', 'fontScheme'])
    ];
};
