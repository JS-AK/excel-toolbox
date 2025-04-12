/**
 * Converts a 0-based column index to an Excel-style letter (A, B, ..., Z, AA, AB, ...).
 *
 * @param index - The 0-based column index.
 * @returns The Excel-style letter for the given column index.
 */
export function columnIndexToLetter(index: number): string {
	let letters = "";
	while (index >= 0) {
		letters = String.fromCharCode((index % 26) + 65) + letters;
		index = Math.floor(index / 26) - 1;
	}
	return letters;
}
