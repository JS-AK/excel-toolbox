/**
 * Compares two column strings and returns a number indicating their relative order.
 *
 * @param a - The first column string to compare.
 * @param b - The second column string to compare.
 * @returns 0 if the columns are equal, -1 if the first column is less than the second, or 1 if the first column is greater than the second.
 */
export function compareColumns(a: string, b: string): number {
	if (a === b) {
		return 0;
	}

	return a.length === b.length ? (a < b ? -1 : 1) : (a.length < b.length ? -1 : 1);
}
