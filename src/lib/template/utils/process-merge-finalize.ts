import { updateDimension } from "./update-dimension.js";

export type ProcessMergeFinalizeData = {
	initialMergeCells: string[];
	lastIndex: number;
	resultRows: string[];
	sharedStrings: string[];
	sharedStringsHeader: string | null;
	sheetMergeCells: string[];
	sheetXml: string;
};

export type ProcessMergeFinalizeResult = {
	shared: string;
	sheet: string;
};

/**
 * Finalizes the processing of the merged sheet by updating the merge cells and
 * inserting them into the sheet XML. It also returns the modified sheet XML and
 * shared strings.
 *
 * @param {object} data - An object containing the following properties:
 *   - `initialMergeCells`: The initial merge cells from the original sheet.
 *   - `lastIndex`: The last processed position in the sheet XML.
 *     describing the merge cells.
 *   - `resultRows`: An array of processed XML rows.
 *   - `sharedStrings`: An array of shared strings.
 *   - `sharedStringsHeader`: The XML declaration of the shared strings.
 *   - `sheetMergeCells`: An array of merge cell XML strings.
 *   - `sheetXml`: The original sheet XML string.
 *
 * @returns An object with two properties:
 *   - `shared`: The modified shared strings XML string.
 *   - `sheet`: The modified sheet XML string with updated merge cells.
 */
export function processMergeFinalize(data: ProcessMergeFinalizeData): ProcessMergeFinalizeResult {
	const {
		initialMergeCells,
		lastIndex,
		resultRows,
		sharedStrings,
		sharedStringsHeader,
		sheetMergeCells,
		sheetXml,
	} = data;

	resultRows.push(sheetXml.slice(lastIndex));

	// Form XML for mergeCells if there are any
	const mergeXml = sheetMergeCells.length
		? `<mergeCells count="${sheetMergeCells.length}">${sheetMergeCells.join("")}</mergeCells>`
		: initialMergeCells;

	// Insert mergeCells before the closing sheetData tag
	const sheetWithMerge = resultRows.join("").replace(/<\/sheetData>/, `</sheetData>${mergeXml}`);

	// Return modified sheet XML and shared strings
	return {
		shared: `${sharedStringsHeader}\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">${sharedStrings.join("")}</sst>`,
		sheet: updateDimension(sheetWithMerge),
	};
}
