/**
 * Processes the sheet XML by extracting the initial <mergeCells> block and
 * extracting all merge cell references. The function returns an object with
 * three properties:
 * - `initialMergeCells`: The initial <mergeCells> block as a string array.
 * - `mergeCellMatches`: An array of objects with `from` and `to` properties,
 *   representing the merge cell references.
 * - `modifiedXml`: The modified sheet XML with the <mergeCells> block removed.
 *
 * @param sheetXml - The sheet XML string.
 * @returns An object with the above three properties.
 */
export function processMergeCells(sheetXml: string) {
	// Regular expression for finding <mergeCells> block
	const mergeCellsBlockRegex = /<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/;

	// Find the first <mergeCells> block (if there are multiple, in xlsx usually there is only one)
	const mergeCellsBlockMatch = sheetXml.match(mergeCellsBlockRegex);

	const initialMergeCells: string[] = [];
	const mergeCellMatches: { from: string; to: string }[] = [];

	if (mergeCellsBlockMatch) {
		const mergeCellsBlock = mergeCellsBlockMatch[0];
		initialMergeCells.push(mergeCellsBlock);

		// Extract <mergeCell ref="A1:B2"/> from this block
		const mergeCellRegex = /<mergeCell ref="([A-Z]+\d+):([A-Z]+\d+)"\/>/g;
		for (const match of mergeCellsBlock.matchAll(mergeCellRegex)) {
			mergeCellMatches.push({ from: match[1]!, to: match[2]! });
		}
	}

	// Remove the <mergeCells> block from the XML
	const modifiedXml = sheetXml.replace(mergeCellsBlockRegex, "");

	return {
		initialMergeCells,
		mergeCellMatches,
		modifiedXml,
	};
};
