export function updateDimension(xml: string): string {
	const cellRefs = [...xml.matchAll(/<c r="([A-Z]+)(\d+)"/g)];
	if (cellRefs.length === 0) return xml;

	let minCol = Infinity, maxCol = -Infinity;
	let minRow = Infinity, maxRow = -Infinity;

	for (const [, colStr, rowStr] of cellRefs) {
		const col = columnLetterToNumber(colStr!);
		const row = parseInt(rowStr!, 10);
		if (col < minCol) minCol = col;
		if (col > maxCol) maxCol = col;
		if (row < minRow) minRow = row;
		if (row > maxRow) maxRow = row;
	}

	const newRef = `${columnNumberToLetter(minCol)}${minRow}:${columnNumberToLetter(maxCol)}${maxRow}`;
	return xml.replace(/<dimension ref="[^"]*"/, `<dimension ref="${newRef}"`);
}

function columnLetterToNumber(letters: string): number {
	let num = 0;
	for (let i = 0; i < letters.length; i++) {
		num = num * 26 + (letters.charCodeAt(i) - 64);
	}
	return num;
}

function columnNumberToLetter(num: number): string {
	let letters = "";

	while (num > 0) {
		const rem = (num - 1) % 26;
		letters = String.fromCharCode(65 + rem) + letters;
		num = Math.floor((num - 1) / 26);
	}

	return letters;
}
