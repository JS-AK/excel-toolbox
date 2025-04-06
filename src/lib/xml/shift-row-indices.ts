/**
 * Adjusts row indices in Excel XML row elements by a specified offset.
 * Handles both row element attributes and cell references within rows.
 *
 * This function is particularly useful when merging sheets or rearranging
 * worksheet content while maintaining proper Excel XML structure.
 *
 * @param {string[]} rows - Array of XML <row> elements as strings
 * @param {number} offset - Numeric value to adjust row indices by:
 *                         - Positive values shift rows down
 *                         - Negative values shift rows up
 * @returns {string[]} - New array with modified row elements containing updated indices
 *
 * @example
 * // Shifts rows down by 2 positions
 * shiftRowIndices([`<row r="1"><c r="A1"/></row>`], 2);
 * // Returns: [`<row r="3"><c r="A3"/></row>`]
 */
export function shiftRowIndices(rows: string[], offset: number): string[] {
	return rows.map(row => {
		// Process each row element through two replacement phases:

		// 1. Update the row's own index (r="N" attribute)
		let adjustedRow = row.replace(
			/(<row[^>]*\br=")(\d+)(")/,
			(_, prefix, rowIndex, suffix) => {
				return `${prefix}${parseInt(rowIndex) + offset}${suffix}`;
			},
		);

		// 2. Update all cell references within the row (r="AN" attributes)
		adjustedRow = adjustedRow.replace(
			/(<c[^>]*\br=")([A-Z]+)(\d+)(")/g,
			(_, prefix, columnLetter, cellRowIndex, suffix) => {
				return `${prefix}${columnLetter}${parseInt(cellRowIndex) + offset}${suffix}`;
			},
		);

		return adjustedRow;
	});
}
