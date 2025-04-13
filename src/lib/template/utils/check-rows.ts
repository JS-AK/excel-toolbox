import { checkRow } from "./check-row.js";

/**
 * Validates an array of row objects to ensure each cell reference is valid.
 * Each row object is checked to ensure that its keys (cell references) are
 * composed of valid column letters (e.g., "A", "B", "C").
 *
 * @param rows An array of row objects, where each object represents a row
 * of data with cell references as keys and cell values as strings.
 *
 * @throws {Error} If any cell reference in the rows is invalid.
 */
export function checkRows(rows: Record<string, string>[]): void {
	for (const row of rows) {
		checkRow(row);
	}
}
