
/**
 * Shifts the row number in a cell reference by the specified number of rows.
 * The function takes a cell reference string in the format "A1" and a row shift value.
 * It returns the shifted cell reference string.
 *
 * @example
 * // Shifts the cell reference "A1" down by 2 rows, resulting in "A3"
 * shiftCellRef('A1', 2);
 * @param {string} cellRef - The cell reference string to be shifted
 * @param {number} rowShift - The number of rows to shift the reference by
 * @returns {string} - The shifted cell reference string
 */
export function shiftCellRef(cellRef: string, rowShift: number): string {
	const match = cellRef.match(/^([A-Z]+)(\d+)$/);

	if (!match) return cellRef;

	const col = match[1];

	if (!match[2]) return cellRef;

	const row = parseInt(match[2], 10);

	return `${col}${row + rowShift}`;
}
