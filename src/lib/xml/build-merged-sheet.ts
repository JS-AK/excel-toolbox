/**
 * Builds a new XML string for a merged Excel sheet by combining the original XML
 * with merged rows and optional cell merge information.
 *
 * This function replaces the sheet data content in the original XML with the merged rows
 * and optionally adds merge cell definitions at the end of the sheet data.
 *
 * @param {string} originalXml - The original XML string of the Excel worksheet
 * @param {string[]} mergedRows - Array of XML strings representing each row in the merged sheet
 * @param {Object[]} [mergeCells] - Optional array of merge cell definitions
 *        Each object should have a 'ref' property specifying the merge range (e.g., "A1:B2")
 * @returns {string} - The reconstructed XML string with merged content
 */
export function buildMergedSheet(
	originalXml: string,
	mergedRows: string[],
	mergeCells: { ref: string }[] = [],
): string {
	// Replace the entire sheetData section in the original XML with our merged rows
	// The regex matches:
	// - Opening <sheetData> tag with any attributes
	// - Any content between opening and closing tags (including line breaks)
	// - Closing </sheetData> tag
	let xmlData = originalXml.replace(
		/<sheetData[^>]*>[\s\S]*?<\/sheetData>/,
		`<sheetData>\n${mergedRows.join("\n")}\n</sheetData>`,
	);

	// If merge cells were specified, add them after the sheetData section
	if (mergeCells.length > 0) {
		// Create mergeCells XML section:
		// - Includes count attribute with total number of merges
		// - Contains one mergeCell element for each merge definition
		const mergeCellsXml = `<mergeCells count="${mergeCells.length}">${mergeCells.map(mc => `<mergeCell ref="${mc.ref}"/>`).join("")}</mergeCells>`;

		// Insert the mergeCells section immediately after the sheetData closing tag
		xmlData = xmlData.replace("</sheetData>", `</sheetData>${mergeCellsXml}`);
	}

	// Return the fully reconstructed XML
	return xmlData;
}
