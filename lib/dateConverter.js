"use strict";

// The base date = 0.
const dateBase = new Date(1900, 0, 0);

// The date conversion has a bug that assumes 1900 was a leap year. So we need to add one for dates after this.
const incorrectLeapDate = new Date(1900, 1, 28);

// Number of milliseconds in a day.
const millisecondsInDay = 1000 * 60 * 60 * 24;

/**
 * Date converter.
 * @private
 */
module.exports = {
    /**
     * Convert a date to a number for Excel.
     * @param {Date} date - The date.
     * @returns {number} The number.
     */
    dateToNumber(date) {
        let number = (date - dateBase) / millisecondsInDay;

        // "Bug" in Excel that treats 1900 as a leap year.
        if (date > incorrectLeapDate) number += 1;

        return number;
    },

    /**
     * Convert a number to a date.
     * @param {number} number - The number.
     * @returns {Date} The date.
     */
    numberToDate(number) {
        if (number > this.dateToNumber(incorrectLeapDate)) number--;
        const milliseconds = number * millisecondsInDay + dateBase.getTime();
        return new Date(milliseconds);
    }
};
