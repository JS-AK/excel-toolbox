/**
 * Finds the maximum row number in a list of <row> elements.
 * @param {string[]} rows - An array of strings, each representing a <row> element.
 * @returns {number} - The maximum row number.
 */
export function getMaxRowNumber(rows: string[]): number {
	let max = 0;
	for (const row of rows) {
		const match = row.match(/<row[^>]* r="(\d+)"/);
		if (match) {
			if (!match[1]) continue;

			const num = parseInt(match[1], 10);

			if (num > max) max = num;
		}
	}
	return max;
}
