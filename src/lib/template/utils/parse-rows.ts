export function parseRows(innerRows: string): Map<number, string> {
	const rowsMap = new Map<number, string>();

	// Regex to match all <row> elements
	// 1. `(<row[^>]+r="(\d+)"[^>]*\/>)` - for self-closing tags (<row ... />)
	// 2. `|(<row[^>]+r="(\d+)"[^>]*>[\s\S]*?<\/row>)` — for self-closing tags (<row ... />)
	const rowRegex = /(<row[^>]+r="(\d+)"[^>]*\/>)|(<row[^>]+r="(\d+)"[^>]*>[\s\S]*?<\/row>)/g;

	let match;
	while ((match = rowRegex.exec(innerRows)) !== null) {
		// if this is a self-closing tag (<row ... />)
		if (match[1]) {
			const fullRow = match[1];
			const rowNumber = match[2];
			if (!rowNumber) throw new Error("Row number not found");

			rowsMap.set(Number(rowNumber), fullRow);
		}
		// if this is a regular tag (<row>...</row>)
		else if (match[3]) {
			const fullRow = match[3];
			const rowNumber = match[4];

			if (!rowNumber) throw new Error("Row number not found");

			rowsMap.set(Number(rowNumber), fullRow);
		}
	}

	return rowsMap;
}
