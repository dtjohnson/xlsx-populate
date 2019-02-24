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
        // Clone the date and strip the time off.
        const dateOnly = new Date(date.getTime());
        dateOnly.setHours(0, 0, 0, 0);

        // Set the number to be the number of days between the date and the base date.
        // We need to round as daylight savings will cause fractional days, which we don't want.
        let number = Math.round((dateOnly - dateBase) / millisecondsInDay);
        
        // Add the true fractional days from just the milliseconds left in the current day.
        number += (date - dateOnly) / millisecondsInDay;

        // Adjust for the "bug" in Excel that treats 1900 as a leap year.
        if (date > incorrectLeapDate) number += 1;

        return number;
    },

    /**
     * Convert a number to a date.
     * @param {number} number - The number.
     * @returns {Date} The date.
     */
    numberToDate(number) {
        // If the number is greater than the incorrect leap date, we should subtract one.
        if (number > this.dateToNumber(incorrectLeapDate)) number--;
        
        // Break the number of full days and the remaining milliseconds in the current day.
        const fullDays = Math.floor(number);
        const partialMilliseconds = Math.round((number - fullDays) * millisecondsInDay);

        // Create a new date from the base date plus the time in the current day.
        const date = new Date(dateBase.getTime() + partialMilliseconds);

        // Now add the number of full days. JS will properly handle the month/year changes.
        date.setDate(date.getDate() + fullDays);

        return date;
    }
};
