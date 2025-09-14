/**
 * Trims whitespace from multiline strings and joins them with a specified separator.
 *
 * This function processes multiline text by:
 * - Splitting the input string by line breaks (handles both \n and \r\n)
 * - Trimming whitespace from each line
 * - Optionally normalizing multiple spaces to single spaces
 * - Optionally filtering out empty lines
 * - Joining the processed lines with a custom separator
 *
 * @param options - Configuration object for processing the multiline string
 * @param options.inputString - The multiline string to process
 * @param options.keepEmptyLines - Whether to preserve empty lines in the output (default: false)
 * @param options.normalizeSpaces - Whether to normalize multiple consecutive spaces to single spaces (default: true)
 * @param options.separator - The string to use when joining lines (default: " ")
 * @returns The processed string with lines joined by the separator
 */
export function trimAndJoinMultiline(options: {
	inputString: string;
	keepEmptyLines?: boolean;
	normalizeSpaces?: boolean;
	separator?: string;
}): string {
	const {
		inputString,
		keepEmptyLines = false,
		normalizeSpaces = true,
		separator = " ",
	} = options;

	const lines = inputString.split(/\r?\n/);
	let trimmedLines = lines.map(line => line.trim());

	if (normalizeSpaces) {
		trimmedLines = trimmedLines.map(line => line.replace(/\s+/g, " "));
	}

	const filteredLines = keepEmptyLines
		? trimmedLines
		: trimmedLines.filter(line => line.length > 0);

	return filteredLines.join(separator);
}
