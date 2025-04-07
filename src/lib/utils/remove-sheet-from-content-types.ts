/**
 * Removes the specified sheet from the Content_Types.xml file.
 * @param {string} xml - The Content_Types.xml file contents as a string
 * @param {number} sheetIndex - The 1-based index of the sheet to remove
 * @returns {string} - The modified Content_Types.xml file contents
 */
export function removeSheetFromContentTypes(xml: string, sheetIndex: number): string {
	return xml.replace(
		new RegExp(
			`<Override[^>]+PartName=["']/xl/worksheets/sheet${sheetIndex}\\.xml["'][^>]*/>`,
			"g",
		),
		"",
	);
}
