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
 * @returns {Promise<Buffer>} - The merged Excel file
 */
export async function mergeSheetsToBaseFile(data: {
	additions: { file: Buffer; isBaseFile?: boolean; sheetIndexes: number[] }[];
	baseFile: Buffer;
	baseSheetIndex?: number;
	gap?: number;
	sheetNamesToRemove?: string[];
	sheetsToRemove?: number[];
}): Promise<Buffer> {
	const {
		additions = [],
		baseFile,
		baseSheetIndex = 1,
		gap = 0,
		sheetNamesToRemove = [],
		sheetsToRemove = [],
	} = data;
	const baseFiles = await Zip.read(baseFile);

	const additionsUpdated: { files: Record<string, Buffer>; sheetIndexes: number[] }[] = [];

	for (const { file, isBaseFile, sheetIndexes } of additions) {
		const files = (isBaseFile || Utils.isSameBuffer(file, baseFile))
			? baseFiles
			: await Zip.read(file);

		additionsUpdated.push({
			files,
			sheetIndexes,
		});
	}

	await mergeSheetsToBaseFileProcess({
		additions: additionsUpdated,
		baseFiles,
		baseSheetIndex,
		gap,
		sheetNamesToRemove,
		sheetsToRemove,
	});

	return Zip.create(baseFiles);
}
