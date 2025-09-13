/**
 * Converts a JavaScript Date object to Excel serial number format.
 * Excel stores dates as serial numbers where 1 represents January 1, 1900.
 *
 * @param date - The JavaScript Date object to convert
 * 
 * @returns The Excel serial number as a floating-point number
 */
export function dateToExcelSerial(date: Date): number {
	const msPerDay = 24 * 60 * 60 * 1000;
	const excelEpoch = Date.UTC(1899, 11, 30); // Excel "day 0"

	return (date.getTime() - excelEpoch) / msPerDay;
}
