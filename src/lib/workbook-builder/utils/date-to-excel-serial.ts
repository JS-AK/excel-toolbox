export function dateToExcelSerial(date: Date) {
	const msPerDay = 24 * 60 * 60 * 1000;
	const excelEpoch = Date.UTC(1899, 11, 30); // Excel "day 0"

	return (date.getTime() - excelEpoch) / msPerDay;
}
