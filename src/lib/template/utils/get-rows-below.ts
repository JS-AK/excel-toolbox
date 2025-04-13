/**
 * Filters out all rows in the given map that have a row number
 * greater than or equal to the given maxRow, and returns the
 * filtered rows as a single string. Useful for removing rows
 * from a template that are positioned below a certain row
 * number.
 *
 * @param {Map<number, string>} map - The map of row numbers to row content
 * @param {number} maxRow - The maximum row number to include in the output
 * @returns {string} The filtered rows, concatenated into a single string
 */
export function getRowsBelow(map: Map<number, string>, maxRow: number): string {
	const filteredRows: string[] = [];

	for (const [key, value] of map.entries()) {
		if (key < maxRow) {
			filteredRows.push(value);
		}
	}

	return filteredRows.join("");
}
