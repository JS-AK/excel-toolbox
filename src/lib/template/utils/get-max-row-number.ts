/**
 * Finds the maximum row number in a list of <row> elements
 * and returns the maximum row number + 1.
 * @param {string} line - The line of XML to parse.
 * @returns {number} - The maximum row number found + 1.
 */
export function getMaxRowNumber(line : string): number {
	let result = 1;

	const rowMatches = [...line.matchAll(/<row[^>]+r="(\d+)"[^>]*>/g)];

	for (const match of rowMatches) {
		const rowNum = parseInt(match[1] as string, 10);

		if (rowNum >= result) {
			result = rowNum + 1;
		}
	}

	return result;
}
