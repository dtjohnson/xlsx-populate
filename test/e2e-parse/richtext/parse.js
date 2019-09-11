"use strict";

module.exports = workbook => {
    const cell = workbook.sheet(0).cell('A2');
    const b1 = workbook.sheet(0).cell('B1');

    return [
        cell.value().get(1).style(['fontFamily', 'fontColor']),
        b1.value().get(0).value(),
        b1.value().get(0).style(['bold', 'italic']),
        b1.value().get(1).value(),
        b1.value().get(1).style(['italic', 'fontColor']),
        b1.value().get(2).value(),
        b1.value().get(2).style(['underline', 'fontColor']),
        b1.value().get(3).value(),
        b1.value().get(3).style(['strikethrough', 'fontColor']),
        b1.value().get(4).value(),
        b1.value().get(4).style(['subscript', 'fontColor']),
        b1.value().get(5).style(['superscript', 'fontSize', 'fontFamily', 'fontGenericFamily', 'fontScheme']),
        b1.value().get(6).style('fontFamily')
    ];
};
