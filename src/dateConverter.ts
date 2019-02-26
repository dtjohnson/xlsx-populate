// The base date = 0.
const dateBase = new Date(1900, 0, 0);

// The date conversion has a bug that assumes 1900 was a leap year. So we need to add one for dates after this.
const incorrectLeapDate = new Date(1900, 1, 28);

// Number of milliseconds in a day.
const millisecondsInDay = 1000 * 60 * 60 * 24;

/**
 * Convert a date to a number for Excel.
 * @param date - The date.
 * @returns The number.
 */
export function dateToNumber(date: Date): number {
    // Clone the date and strip the time off.
    const dateOnly = new Date(date.getTime());
    dateOnly.setHours(0, 0, 0, 0);

    // Set the number to be the number of days between the date and the base date.
    // We need to round as daylight savings will cause fractional days, which we don't want.
    let num = Math.round((dateOnly.getTime() - dateBase.getTime()) / millisecondsInDay);

    // Add the true fractional days from just the milliseconds left in the current day.
    num += (date.getTime() - dateOnly.getTime()) / millisecondsInDay;

    // Adjust for the "bug" in Excel that treats 1900 as a leap year.
    if (date > incorrectLeapDate) num += 1;

    return num;
}

/**
 * Convert a number to a date.
 * @param num - The number.
 * @returns The date.
 */
export function numberToDate(num: number): Date {
    // If the number is greater than the incorrect leap date, we should subtract one.
    if (num > dateToNumber(incorrectLeapDate)) num--;

    // Break the number of full days and the remaining milliseconds in the current day.
    const fullDays = Math.floor(num);
    const partialMilliseconds = Math.round((num - fullDays) * millisecondsInDay);

    // Create a new date from the base date plus the time in the current day.
    const date = new Date(dateBase.getTime() + partialMilliseconds);

    // Now add the number of full days. JS will properly handle the month/year changes.
    date.setDate(date.getDate() + fullDays);

    return date;
}
