/**
 * Filters out all rows in the given map that have a row number
 * lower than or equal to the given minRow, and returns the
 * filtered rows as a single string. Useful for removing rows
 * from a template that are positioned above a certain row
 * number.
 *
 * @param {Map<number, string>} map - The map of row numbers to row content
 * @param {number} minRow - The minimum row number to include in the output
 * @returns {string} The filtered rows, concatenated into a single string
 */
export function getRowsAbove(map: Map<number, string>, minRow: number): string {
	const filteredRows: string[] = [];

	for (const [key, value] of map.entries()) {
		if (key > minRow) {
			filteredRows.push(value);
		}
	}

	return filteredRows.join("");
}
