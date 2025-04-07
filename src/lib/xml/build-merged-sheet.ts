/**
 * Builds a new XML string for a merged Excel sheet by combining the original XML
 * with merged rows and optional cell merge information.
 *
 * This function replaces the <sheetData> section in the original XML with the provided merged rows,
 * and optionally adds <mergeCells> definitions after the </sheetData> tag.
 *
 * @param {string} originalXml - The original XML string of the Excel worksheet.
 * @param {string[]} mergedRows - Array of XML strings representing each row in the merged sheet.
 * @param {Object[]} [mergeCells] - Optional array of merge cell definitions.
 *        Each object should have a 'ref' property specifying the merge range (e.g., "A1:B2").
 * @returns {string} - The reconstructed XML string with merged content.
 */
export function buildMergedSheet(
	originalXml: string,
	mergedRows: string[],
	mergeCells: { ref: string }[] = [],
): string {
	// Remove any existing <mergeCells> section from the XML
	let xmlData = originalXml.replace(/<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/g, "");

	// Construct a new <sheetData> section with the provided rows
	const sheetDataXml = `<sheetData>${mergedRows.join("")}</sheetData>`;

	// Replace the existing <sheetData> section with the new one
	xmlData = xmlData.replace(/<sheetData[^>]*>[\s\S]*?<\/sheetData>/, sheetDataXml);

	if (mergeCells.length > 0) {
		// Construct a new <mergeCells> section with the provided merge references
		const mergeCellsXml = `<mergeCells count="${mergeCells.length}">${mergeCells.map(mc => `<mergeCell ref="${mc.ref}"/>`).join("")}</mergeCells>`;

		// Insert <mergeCells> after </sheetData> and before the next XML tag
		xmlData = xmlData.replace(
			/(<\/sheetData>)(\s*<)/,
			`$1\n${mergeCellsXml}\n$2`,
		);
	}

	return xmlData;
}
