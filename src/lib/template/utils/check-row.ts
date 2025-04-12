/**
 * Validates that each key in the given row object is a valid cell reference.
 *
 * This function checks that all keys in the provided row object are composed
 * only of column letters (A-Z, case insensitive). If a key is found that does
 * not match this pattern, an error is thrown with a message indicating the
 * invalid cell reference.
 *
 * @param row - An object representing a row of data, where keys are cell
 *              references and values are strings.
 *
 * @throws {Error} If any key in the row is not a valid column letter.
 */
export function checkRow(row: Record<string, string>): void {
	for (const key of Object.keys(row)) {
		if (!/^[A-Z]+$/i.test(key)) {
			throw new Error(
				`Invalid cell reference "${key}" in row. Only column letters (like "A", "B", "C") are allowed.`,
			);
		}
	}
}
