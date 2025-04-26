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
 * @param {WritableLike} output - A file write stream to write the Excel XML to.
 * @param {AsyncIterable<unknown[] | unknown[][]>} rows - An async iterable of rows, where each row is an array
 *                                                        of values or an array of arrays of values.
 * @param {number} startRowNumber - The starting row number to use for the first
 *                                 row written to the file.
 * @returns {Promise<{
 *   dimension: {
 *     maxColumn: string;
 *     maxRow: number;
 *     minColumn: string;
 *     minRow: number;
 *   };
 *   rowNumber: number;
 * }>} An object containing:
 *   - dimension: The boundaries of the written data (min/max columns and rows)
 *   - rowNumber: The last row number written to the file
 */
export async function writeRowsToStream(
	output: WritableLike,
	rows: AsyncIterable<unknown[] | unknown[][]>,
	startRowNumber: number,
): Promise<{
	dimension: { maxColumn: string; maxRow: number; minColumn: string; minRow: number };
	rowNumber: number;
}> {
	let rowNumber = startRowNumber;

	const dimension = {
		maxColumn: "A",
		maxRow: startRowNumber,
		minColumn: "A",
		minRow: startRowNumber,
	};

	// Функция для сравнения колонок (A < B, AA > Z и т.д.)
	const compareColumns = (a: string, b: string): number => {
		if (a === b) return 0;
		return a.length === b.length ? (a < b ? -1 : 1) : (a.length < b.length ? -1 : 1);
	};

	const processRow = (row: unknown[], currentRowNumber: number) => {
		const cells = prepareRowToCells(row, currentRowNumber);
		if (cells.length === 0) return;

		output.write(`<row r="${currentRowNumber}">${cells.map(cell => cell.cellXml).join("")}</row>`);

		// Обновление границ
		const firstCellRef = cells[0]?.cellRef;
		const lastCellRef = cells[cells.length - 1]?.cellRef;

		if (firstCellRef) {
			const colLetters = firstCellRef.match(/[A-Z]+/)?.[0] || "";
			if (compareColumns(colLetters, dimension.minColumn) < 0) {
				dimension.minColumn = colLetters;
			}
		}

		if (lastCellRef) {
			const colLetters = lastCellRef.match(/[A-Z]+/)?.[0] || "";
			if (compareColumns(colLetters, dimension.maxColumn) > 0) {
				dimension.maxColumn = colLetters;
			}
		}

		dimension.maxRow = currentRowNumber;
	};

	for await (const row of rows) {
		if (!row.length) continue;

		if (Array.isArray(row[0])) {
			for (const subRow of row as unknown[][]) {
				if (!subRow.length) continue;

				processRow(subRow, rowNumber);

				rowNumber++;
			}
		} else {
			processRow(row, rowNumber);

			rowNumber++;
		}
	}

	return { dimension, rowNumber };
}
