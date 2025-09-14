/**
 * Adds a custom number format to the collection or returns the existing id.
 *
 * Excel reserves built-in ids below 164; custom formats start from 164.
 *
 * @param payload - Input arguments
 * @param payload.formatCode - Format code, e.g., "0.00" or "dd/mm/yyyy"
 * @param payload.numFmts - Mutable list of custom formats to search/extend
 * @returns Numeric id of the number format
 */
export const addNumFmt = (payload: {
	formatCode: string;
	numFmts: { formatCode: string; id: number }[];
}): number => {
	const { formatCode, numFmts } = payload;

	// 164+ is reserved for custom formats
	const existing = numFmts.find(nf => nf.formatCode === formatCode);

	if (existing) {
		return existing.id;
	}

	const id = 164 + numFmts.length;

	numFmts.push({ formatCode, id });

	return id;
};
