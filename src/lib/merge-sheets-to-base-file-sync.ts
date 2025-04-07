import * as Utils from "./utils/index.js";
import * as Zip from "./zip/index.js";

import { mergeSheetsToBaseFileProcess } from "./merge-sheets-to-base-file-process.js";

/**
 * Merge rows from other Excel files into a base Excel file.
 * The output is a new Excel file with the merged content.
 *
 * @param {Object} data
 * @param {Object[]} data.additions
 * @param {Buffer} data.additions.file - The file to extract rows from
 * @param {number[]} data.additions.sheetIndexes - The 1-based indexes of the sheet to extract rows from
 * @param {Buffer} data.baseFile - The base file to add rows to
 * @param {number} [data.baseSheetIndex=1] - The 1-based index of the sheet in the base file to add rows to
 * @param {number} [data.gap=0] - The number of empty rows to insert between each added section
 * @param {string[]} [data.sheetNamesToRemove=[]] - The names of sheets to remove from the output file
 * @param {number[]} [data.sheetsToRemove=[]] - The 1-based indices of sheets to remove from the output file
 * @returns {Buffer} - The merged Excel file
 */
export function mergeSheetsToBaseFileSync(data: {
	additions: { file: Buffer; isBaseFile?: boolean; sheetIndexes: number[] }[];
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
		gap = 0,
		sheetNamesToRemove = [],
		sheetsToRemove = [],
	} = data;
	const baseFiles = Zip.readSync(baseFile);

	const additionsUpdated: { files: Record<string, string>; sheetIndexes: number[] }[] = [];

	for (const { file, isBaseFile, sheetIndexes } of additions) {
		const files = (isBaseFile || Utils.isSameBuffer(file, baseFile))
			? baseFiles
			: Zip.readSync(file);

		additionsUpdated.push({
			files,
			sheetIndexes,
		});
	}

	mergeSheetsToBaseFileProcess({
		additions: additionsUpdated,
		baseFiles,
		baseSheetIndex,
		gap,
		sheetNamesToRemove,
		sheetsToRemove,
	});

	return Zip.createSync(baseFiles);
}
