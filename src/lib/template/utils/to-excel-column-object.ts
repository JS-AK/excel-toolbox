import { columnIndexToLetter } from "./column-index-to-letter.js";

/**
 * Converts an array of values into a Record<string, string> with Excel column names as keys.
 *
 * The column names are generated in the standard Excel column naming convention (A, B, ..., Z, AA, AB, ...).
 * The corresponding values are converted to strings using the String() function.
 *
 * @param values - The array of values to convert
 * @returns The resulting Record<string, string>
 */
export function toExcelColumnObject(values: unknown[]): Record<string, string> {
	const result: Record<string, string> = {};

	for (let i = 0; i < values.length; i++) {
		const key = columnIndexToLetter(i);
		result[key] = String(values[i]);
	}

	return result;
}
