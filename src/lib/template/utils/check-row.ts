/**
 * Validates an object representing a single row of data to ensure that its keys
 * are valid Excel column references. Throws an error if any of the keys are
 * invalid.
 *
 * @param row An object with string keys that represent the cell references and
 * string values that represent the values of those cells.
 */
export function checkRow(row: Record<string, string>): void {
	for (const key of Object.keys(row)) {
		if (!/^[A-Z]+$/i.test(key) || !/^[A-Z]$|^[A-Z][A-Z]$|^[A-Z][A-Z][A-Z]$/i.test(key)) {
			throw new Error(
				`Invalid cell reference "${key}" in row. Only column letters (like "A", "B", "C") are allowed.`,
			);
		}
	}
}
