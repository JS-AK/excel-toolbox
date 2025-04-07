/**
 * Removes the specified sheet from the workbook (xl/workbook.xml).
 * @param {string} xml - The workbook file contents as a string
 * @param {number} sheetIndex - The 1-based index of the sheet to remove
 * @returns {string} - The modified workbook file contents
 */
export function removeSheetFromWorkbook(xml: string, sheetIndex: number): string {
	return xml.replace(
		new RegExp(`<sheet[^>]+sheetId=["']${sheetIndex}["'][^>]*/>`, "g"),
		"",
	);
}
