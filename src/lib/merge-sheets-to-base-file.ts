import { extractRowsFromSheet } from './xml/extract-rows-from-sheet.js';
import { shiftRowIndices } from './xml/shift-row-indices.js';
import { buildMergedSheet } from './xml/build-merged-sheet.js';

import * as Zip from './zip/index.js';

export async function mergeSheetsToBaseFile(data: {
	baseFile: Buffer,
	baseSheetIndex?: number,
	additions: { file: Buffer, sheetIndex: number, }[],
	gap?: number
}) {
	const {
		additions = [],
		baseFile,
		baseSheetIndex = 1,
		gap = 1,
	} = data;
	const baseFiles = Zip.read(baseFile);
	const basePath = `xl/worksheets/sheet${baseSheetIndex}.xml`;

	if (!baseFiles[basePath]) {
		throw new Error(`Base file does not contain ${basePath}`);
	}

	const {
		lastRowNumber,
		mergeCells: baseMergeCells,
		rows: baseRows,
	} = extractRowsFromSheet(baseFiles[basePath] as string);

	const allRows = [...baseRows];
	const allMergeCells = [...(baseMergeCells || [])];
	let currentRowOffset = lastRowNumber + gap;

	for (const { file, sheetIndex } of additions) {
		const files = isSameBuffer(file, baseFile) ? baseFiles : Zip.read(file);
		const sheetPath = `xl/worksheets/sheet${sheetIndex}.xml`;

		if (!files[sheetPath]) {
			throw new Error(`File does not contain ${sheetPath}`);
		}

		const { mergeCells, rows } = extractRowsFromSheet(files[sheetPath] as string);

		const shiftedRows = shiftRowIndices(rows, currentRowOffset);

		const shiftedMergeCells = (mergeCells || []).map(cell => {
			const [start, end] = cell.ref.split(':');

			if (!start || !end) {
				return cell;
			}

			const shiftedStart = shiftCellRef(start, currentRowOffset);
			const shiftedEnd = shiftCellRef(end, currentRowOffset);
			return { ...cell, ref: `${shiftedStart}:${shiftedEnd}` };
		});

		allRows.push(...shiftedRows);
		allMergeCells.push(...shiftedMergeCells);
		currentRowOffset += getMaxRowNumber(rows) + gap;
	}

	const mergedXml = buildMergedSheet(
		baseFiles[basePath] as string,
		allRows,
		allMergeCells,
	);

	baseFiles[basePath] = mergedXml;

	const zip = Zip.create(baseFiles);

	return zip;
}

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
function shiftCellRef(cellRef: string, rowShift: number): string {
	const match = cellRef.match(/^([A-Z]+)(\d+)$/);

	if (!match) return cellRef;

	const col = match[1];

	if (!match[2]) return cellRef;

	const row = parseInt(match[2], 10);

	return `${col}${row + rowShift}`;
}

/**
 * Checks if two Buffers are the same
 * @param {Buffer} buf1 - the first Buffer
 * @param {Buffer} buf2 - the second Buffer
 * @returns {boolean} - true if the Buffers are the same, false otherwise
 */
function isSameBuffer(buf1: Buffer, buf2: Buffer): boolean {
	return buf1.equals(buf2);
}

/**
 * Finds the maximum row number in a list of <row> elements.
 * @param {string[]} rows - An array of strings, each representing a <row> element.
 * @returns {number} - The maximum row number.
 */
function getMaxRowNumber(rows: string[]): number {
	let max = 0;
	for (const row of rows) {
		const match = row.match(/<row[^>]* r="(\d+)"/);
		if (match) {
			if (!match[1]) continue;

			const num = parseInt(match[1], 10);

			if (num > max) max = num;
		}
	}
	return max;
}
