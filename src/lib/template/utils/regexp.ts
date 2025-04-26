/**
 * Creates a regular expression to match a relationship element with a specific ID.
 *
 * @param {string} id - The relationship ID to match (e.g. "rId1")
 * @returns {RegExp} A regular expression that matches a Relationship XML element with the given ID and captures the Target attribute value
 * @example
 * const regex = relationshipMatch("rId1");
 * const xml = '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>';
 * const match = xml.match(regex);
 * // match[1] === "worksheets/sheet1.xml"
 */
export function relationshipMatch(id: string): RegExp {
	return new RegExp(`<Relationship[^>]+Id="${id}"[^>]+Target="([^"]+)"[^>]*/>`);
}

/**
 * Creates a regular expression to match a sheet element with a specific name.
 *
 * @param {string} sheetName - The name of the sheet to match
 * @returns {RegExp} A regular expression that matches a sheet XML element with the given name and captures the r:id attribute value
 * @example
 * const regex = sheetMatch("Sheet1");
 * const xml = '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>';
 * const match = xml.match(regex);
 * // match[1] === "rId1"
 */
export function sheetMatch(sheetName: string): RegExp {
	return new RegExp(`<sheet[^>]+name="${sheetName}"[^>]+r:id="([^"]+)"[^>]*/>`);
}
