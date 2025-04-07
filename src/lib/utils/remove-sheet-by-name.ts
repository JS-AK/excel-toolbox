/**
 * Removes a sheet from the Excel workbook by name.
 * @param {Object.<string, string | Buffer>} files - The dictionary of files in the workbook.
 * @param {string} sheetName - The name of the sheet to remove.
 * @returns {void}
 */
export function removeSheetByName(files: Record<string, string>, sheetName: string): void {
	const workbookXml = files["xl/workbook.xml"];
	const relsXml = files["xl/_rels/workbook.xml.rels"];

	if (!workbookXml || !relsXml) {
		return;
	}

	const sheetMatch = workbookXml.match(new RegExp(`<sheet[^>]+name=["']${sheetName}["'][^>]*/>`));

	if (!sheetMatch) {
		return;
	}

	const sheetTag = sheetMatch[0];
	const sheetIdMatch = sheetTag.match(/sheetId=["'](\d+)["']/);
	const ridMatch = sheetTag.match(/r:id=["'](rId\d+)["']/);

	if (!sheetIdMatch || !ridMatch) {
		return;
	}

	const relId = ridMatch[1];

	const relMatch = relsXml.match(new RegExp(`<Relationship[^>]+Id=["']${relId}["'][^>]+Target=["']([^"']+)["'][^>]*/>`));

	if (!relMatch) {
		return;
	}

	const relTag = relMatch[0];
	const targetMatch = relTag.match(/Target=["']([^"']+)["']/);

	if (!targetMatch) {
		return;
	}

	const targetPath = `xl/${targetMatch[1]}`.replace(/\\/g, "/");

	delete files[targetPath];

	files["xl/workbook.xml"] = workbookXml.replace(sheetTag, "");
	files["xl/_rels/workbook.xml.rels"] = relsXml.replace(relTag, "");

	const contentTypes = files["[Content_Types].xml"];

	if (contentTypes) {
		files["[Content_Types].xml"] = contentTypes.replace(
			new RegExp(`<Override[^>]+PartName=["']/${targetPath}["'][^>]*/>`, "g"),
			"",
		);
	}
}
