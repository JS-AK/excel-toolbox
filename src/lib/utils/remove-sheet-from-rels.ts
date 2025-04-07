/**
 * Removes the specified sheet from the workbook relationships file (xl/_rels/workbook.xml.rels).
 * @param {string} xml - The workbook relationships file contents as a string
 * @param {number} sheetIndex - The 1-based index of the sheet to remove
 * @returns {string} - The modified workbook relationships file contents
 */
export function removeSheetFromRels(xml: string, sheetIndex: number): string {
	return xml.replace(
		new RegExp(
			`<Relationship[^>]+Target=["']worksheets/sheet${sheetIndex}\\.xml["'][^>]*/>`,
			"g",
		),
		"",
	);
}
