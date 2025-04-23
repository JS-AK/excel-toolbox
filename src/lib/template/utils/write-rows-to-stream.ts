import { prepareRowToCells } from "./prepare-row-to-cells.js";

interface WritableLike {
	write(chunk: string | Buffer): boolean;
	end?: () => void;
}

/**
 * Writes an async iterable of rows to an Excel XML file.
 *
 * Each row is expected to be an array of values, where each value is
 * converted to a string using the `String()` function. Empty values are
 * replaced with an empty string.
 *
 * The `startRowNumber` parameter is used as the starting row number
 * for the first row written to the file. Subsequent rows are written
 * with incrementing row numbers.
 *
 * @param output - A file write stream to write the Excel XML to.
 * @param rows - An async iterable of rows, where each row is an array
 *               of values.
 * @param startRowNumber - The starting row number to use for the first
 *                         row written to the file.
 *
 * @returns An object with a single property `rowNumber`, which is the
 *          last row number written to the file (i.e., the `startRowNumber`
 *          plus the number of rows written).
 */
export async function writeRowsToStream(
	output: WritableLike,
	rows: AsyncIterable<unknown[] | unknown[][]>,
	startRowNumber: number,
): Promise<{ rowNumber: number }> {
	let rowNumber = startRowNumber;

	for await (const row of rows) {
		// Transform the row into XML
		if (Array.isArray(row[0])) {
			for (const subRow of row as unknown[][]) {
				const cells = prepareRowToCells(subRow, rowNumber);

				// Write the row to the file
				output.write(`<row r="${rowNumber}">${cells.join("")}</row>`);

				rowNumber++;
			}
		} else {
			const cells = prepareRowToCells(row, rowNumber);

			// Write the row to the file
			output.write(`<row r="${rowNumber}">${cells.join("")}</row>`);

			rowNumber++;
		}
	}

	return { rowNumber };
}
