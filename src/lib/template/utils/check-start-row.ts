/**
 * Checks if startRow is a positive integer.
 *
 * @param startRow The start row to check.
 *
 * @throws {Error} If startRow is not a positive integer.
 */
export function checkStartRow(startRow?: number): void {
	if (startRow === undefined) {
		return;
	}

	if (!Number.isInteger(startRow) || startRow < 1) {
		throw new Error(`Invalid startRow "${startRow}". Must be a positive integer.`);
	}
}
