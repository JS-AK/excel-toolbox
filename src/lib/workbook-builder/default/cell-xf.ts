export type CellXf = {
	borderId: number;
	fillId: number;
	fontId: number;
	numFmtId: number;
};

/**
 * Creates a default cell formatting object.
 *
 * @returns Cell formatting object with default IDs
 */
export const cellXf = (): CellXf => ({
	borderId: 0,
	fillId: 0,
	fontId: 0,
	numFmtId: 0,
});
