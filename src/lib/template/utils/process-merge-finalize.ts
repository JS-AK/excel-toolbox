import { updateDimension } from "./update-dimension.js";

/**
 * Finalizes the processing of the merged sheet by updating the merge cells and
 * inserting them into the sheet XML. It also returns the modified sheet XML and
 * shared strings.
 *
 * @param {object} data - An object containing the following properties:
 *   - `initialMergeCells`: The initial merge cells from the original sheet.
 *   - `lastIndex`: The last processed position in the sheet XML.
 *   - `mergeCellMatches`: An array of objects with `from` and `to` properties,
 *     describing the merge cells.
 *   - `resultRows`: An array of processed XML rows.
 *   - `rowShift`: The total row shift.
 *   - `sharedStrings`: An array of shared strings.
 *   - `sharedStringsHeader`: The XML declaration of the shared strings.
 *   - `sheetMergeCells`: An array of merge cell XML strings.
 *   - `sheetXml`: The original sheet XML string.
 *
 * @returns An object with two properties:
 *   - `shared`: The modified shared strings XML string.
 *   - `sheet`: The modified sheet XML string with updated merge cells.
 */
export function processMergeFinalize(data: {
	initialMergeCells: string[];
	lastIndex: number;
	mergeCellMatches: { from: string; to: string }[];
	resultRows: string[];
	rowShift: number;
	sharedStrings: string[];
	sharedStringsHeader: string | null;
	sheetMergeCells: string[];
	sheetXml: string;
}) {
	const {
		initialMergeCells,
		lastIndex,
		mergeCellMatches,
		resultRows,
		rowShift,
		sharedStrings,
		sharedStringsHeader,
		sheetMergeCells,
		sheetXml,
	} = data;

	for (const { from, to } of mergeCellMatches) {
		const [, fromCol, fromRow] = from.match(/^([A-Z]+)(\d+)$/)!;
		const [, toCol, toRow] = to.match(/^([A-Z]+)(\d+)$/)!;

		const fromRowNum = Number(fromRow);
		// These rows have already been processed, don't add duplicates
		if (fromRowNum <= lastIndex) continue;

		const newFrom = `${fromCol}${fromRowNum + rowShift}`;
		const newTo = `${toCol}${Number(toRow) + rowShift}`;

		sheetMergeCells.push(`<mergeCell ref="${newFrom}:${newTo}"/>`);
	}

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
