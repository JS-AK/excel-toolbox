/**
 * Converts a zero-based column index to its corresponding Excel column letter.
 *
 * @throws Will throw an error if the input is not a positive integer.
 * @param {number} index - The zero-based index of the column to convert.
 * @returns {string} The corresponding Excel column letter.
 *
 * @example
 * columnIndexToLetter(0); // returns "A"
 * columnIndexToLetter(25); // returns "Z"
 * columnIndexToLetter(26); // returns "AA"
 * columnIndexToLetter(51); // returns "AZ"
 * columnIndexToLetter(52); // returns "BA"
 */
export function columnIndexToLetter(index: number): string {
	if (!Number.isInteger(index) || index < 0) {
		throw new Error(`Invalid column index: ${index}. Must be a positive integer.`);
	}

	let letters = "";

	while (index >= 0) {
		letters = String.fromCharCode((index % 26) + 65) + letters;
		index = Math.floor(index / 26) - 1;
	}

	return letters;
}
