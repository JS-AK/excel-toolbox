import { buildMergedSheet } from "./xml/build-merged-sheet.js";
import { extractRowsFromSheet } from "./xml/extract-rows-from-sheet.js";
import { shiftRowIndices } from "./xml/shift-row-indices.js";

import * as Zip from "./zip/index.js";

/**
 * Merge rows from other Excel files into a base Excel file.
 * The output is a new Excel file with the merged content.
 *
 * @param {Object} data
 * @param {Object[]} data.additions
 * @param {Buffer} data.additions.file - The file to extract rows from
 * @param {number} data.additions.sheetIndex - The 1-based index of the sheet to extract rows from
 * @param {Buffer} data.baseFile - The base file to add rows to
 * @param {number} [data.baseSheetIndex=1] - The 1-based index of the sheet in the base file to add rows to
 * @param {number} [data.gap=1] - The number of empty rows to insert between each added section
 * @param {string[]} [data.sheetNamesToRemove=[]] - The names of sheets to remove from the output file
 * @param {number[]} [data.sheetsToRemove=[]] - The 1-based indices of sheets to remove from the output file
 * @returns {Buffer} - The merged Excel file
 */
export function mergeSheetsToBaseFile(data: {
	additions: { file: Buffer; sheetIndex: number }[];
	baseFile: Buffer;
	baseSheetIndex?: number;
	gap?: number;
	sheetNamesToRemove?: string[];
	sheetsToRemove?: number[];
}): Buffer {
	const {
		additions = [],
		baseFile,
		baseSheetIndex = 1,
		gap = 1,
		sheetNamesToRemove = [],
		sheetsToRemove = [],
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
			const [start, end] = cell.ref.split(":");

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

	for (const sheetIndex of sheetsToRemove) {
		const sheetPath = `xl/worksheets/sheet${sheetIndex}.xml`;
		delete baseFiles[sheetPath];

		if (baseFiles["xl/workbook.xml"]) {
			baseFiles["xl/workbook.xml"] = removeSheetFromWorkbook(
				baseFiles["xl/workbook.xml"],
				sheetIndex,
			);
		}

		if (baseFiles["xl/_rels/workbook.xml.rels"]) {
			baseFiles["xl/_rels/workbook.xml.rels"] = removeSheetFromRels(
				baseFiles["xl/_rels/workbook.xml.rels"],
				sheetIndex,
			);
		}

		if (baseFiles["[Content_Types].xml"]) {
			baseFiles["[Content_Types].xml"] = removeSheetFromContentTypes(
				baseFiles["[Content_Types].xml"],
				sheetIndex,
			);
		}
	}

	for (const sheetName of sheetNamesToRemove) {
		removeSheetByName(baseFiles, sheetName);
	}

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

/**
 * Removes the specified sheet from the workbook (xl/workbook.xml).
 * @param {string} xml - The workbook file contents as a string
 * @param {number} sheetIndex - The 1-based index of the sheet to remove
 * @returns {string} - The modified workbook file contents
 */
function removeSheetFromWorkbook(xml: string, sheetIndex: number): string {
	return xml.replace(
		new RegExp(`<sheet[^>]+sheetId=["']${sheetIndex}["'][^>]*/>`, "g"),
		"",
	);
}

/**
 * Removes the specified sheet from the workbook relationships file (xl/_rels/workbook.xml.rels).
 * @param {string} xml - The workbook relationships file contents as a string
 * @param {number} sheetIndex - The 1-based index of the sheet to remove
 * @returns {string} - The modified workbook relationships file contents
 */
function removeSheetFromRels(xml: string, sheetIndex: number): string {
	return xml.replace(
		new RegExp(
			`<Relationship[^>]+Target=["']worksheets/sheet${sheetIndex}\\.xml["'][^>]*/>`,
			"g",
		),
		"",
	);
}

/**
 * Removes the specified sheet from the Content_Types.xml file.
 * @param {string} xml - The Content_Types.xml file contents as a string
 * @param {number} sheetIndex - The 1-based index of the sheet to remove
 * @returns {string} - The modified Content_Types.xml file contents
 */
function removeSheetFromContentTypes(xml: string, sheetIndex: number): string {
	return xml.replace(
		new RegExp(
			`<Override[^>]+PartName=["']/xl/worksheets/sheet${sheetIndex}\\.xml["'][^>]*/>`,
			"g",
		),
		"",
	);
}

/**
 * Removes a sheet from the Excel workbook by name.
 * @param {Object.<string, string | Buffer>} files - The dictionary of files in the workbook.
 * @param {string} sheetName - The name of the sheet to remove.
 * @returns {void}
 */
function removeSheetByName(files: Record<string, string>, sheetName: string): void {
	const workbookXml = files["xl/workbook.xml"];
	const relsXml = files["xl/_rels/workbook.xml.rels"];

	if (!workbookXml || !relsXml) {
		return;
	}

	const sheetMatch = workbookXml.match(new RegExp(`<sheet[^>]+name=["']${sheetName}["'][^>]*/>`));

	if (!sheetMatch) {
		return;
	}

	const sheetTag = sheetMatch[0];
	const sheetIdMatch = sheetTag.match(/sheetId=["'](\d+)["']/);
	const ridMatch = sheetTag.match(/r:id=["'](rId\d+)["']/);

	if (!sheetIdMatch || !ridMatch) {
		return;
	}

	const relId = ridMatch[1];

	const relMatch = relsXml.match(new RegExp(`<Relationship[^>]+Id=["']${relId}["'][^>]+Target=["']([^"']+)["'][^>]*/>`));

	if (!relMatch) {
		return;
	}

	const relTag = relMatch[0];
	const targetMatch = relTag.match(/Target=["']([^"']+)["']/);

	if (!targetMatch) {
		return;
	}

	const targetPath = `xl/${targetMatch[1]}`.replace(/\\/g, "/");

	delete files[targetPath];

	files["xl/workbook.xml"] = workbookXml.replace(sheetTag, "");
	files["xl/_rels/workbook.xml.rels"] = relsXml.replace(relTag, "");

	const contentTypes = files["[Content_Types].xml"];

	if (contentTypes) {
		files["[Content_Types].xml"] = contentTypes.replace(
			new RegExp(`<Override[^>]+PartName=["']/${targetPath}["'][^>]*/>`, "g"),
			"",
		);
	}
}
