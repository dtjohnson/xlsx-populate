"use strict";

module.exports = workbook => {
    const sheet = workbook.sheet(0);
    return {
        left: sheet.pageMargins('left'),
        right: sheet.pageMargins('right'),
        top: sheet.pageMargins('top'),
        bottom: sheet.pageMargins('bottom'),
        header: sheet.pageMargins('header'),
        footer: sheet.pageMargins('footer')
    };
};
