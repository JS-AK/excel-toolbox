import { extractXmlFromSheet } from "./extract-xml-from-sheet.js";

/**
 * Parses a worksheet (either as Buffer or string) to extract row data,
 * last row number, and merge cell information from Excel XML format.
 *
 * This function is particularly useful for processing Excel files in
 * Open XML Spreadsheet format (.xlsx).
 *
 * @param {Buffer|string} sheet - The worksheet content to parse, either as:
 *                               - Buffer (binary Excel sheet)
 *                               - string (raw XML content)
 * @returns {{
 *   rows: string[],
 *   lastRowNumber: number,
 *   mergeCells: {ref: string}[]
 * }} An object containing:
 *   - rows: Array of raw XML strings for each <row> element
 *   - lastRowNumber: Highest row number found in the sheet (1-based)
 *   - mergeCells: Array of merged cell ranges (e.g., [{ref: "A1:B2"}])
 * @throws {Error} If the sheetData section is not found in the XML
 */
export function extractRowsFromSheet(sheet: Buffer | string): {
	rows: string[];
	lastRowNumber: number;
	mergeCells: { ref: string }[];
	xml: string;
} {
	// Convert Buffer input to XML string if needed
	const xml = typeof sheet === "string" ? sheet : extractXmlFromSheet(sheet);

	// Extract the sheetData section containing all rows
	const sheetDataMatch = xml.match(/<sheetData[^>]*>([\s\S]*?)<\/sheetData>/);
	if (!sheetDataMatch) {
		throw new Error("sheetData not found in worksheet XML");
	}

	const sheetDataContent = sheetDataMatch[1] || "";

	// Extract all <row> elements using regex
	const rowMatches = [...sheetDataContent.matchAll(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g)];
	const rows = rowMatches.map(match => match[0]);

	// Calculate the highest row number present in the sheet
	const lastRowNumber = rowMatches
		.map(match => {
			// Extract row number from r="..." attribute (1-based)
			const rowNumMatch = match[0].match(/r="(\d+)"/);
			return rowNumMatch?.[1] ? parseInt(rowNumMatch[1], 10) : null;
		})
		.filter((row): row is number => row !== null) // Type guard to filter out nulls
		.reduce((max, current) => Math.max(max, current), 0); // Find maximum row number

	// Extract all merged cell ranges from the worksheet
	const mergeCells: { ref: string }[] = [];
	const mergeCellsMatch = xml.match(/<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/);

	if (mergeCellsMatch) {
		// Find all mergeCell entries with ref attributes
		const mergeCellMatches = mergeCellsMatch[1]?.match(/<mergeCell[^>]+ref="([^"]+)"[^>]*>/g) || [];

		mergeCellMatches.forEach(match => {
			const refMatch = match.match(/ref="([^"]+)"/);
			if (refMatch?.[1]) {
				mergeCells.push({ ref: refMatch[1] }); // Store the cell range (e.g., "A1:B2")
			}
		});
	}

	return {
		lastRowNumber,
		mergeCells,
		rows,
		xml,
	};
}
