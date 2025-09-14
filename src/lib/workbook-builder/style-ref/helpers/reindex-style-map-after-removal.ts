/**
 * Reindexes a style map in-place after removing a style at a given index.
 *
 * - Deletes entries that point to the removed index
 * - Decrements by 1 all indices greater than the removed index
 *
 * @param payload - Input arguments
 * @param payload.removedIndex - The style index that was removed
 * @param payload.styleMap - Map of serialized style key to style index
 * @returns void
 */
export const reindexStyleMapAfterRemoval = (payload: {
	removedIndex: number;
	styleMap: Map<string, number>;
}): void => {
	const { removedIndex, styleMap } = payload;

	const updates: Array<[string, number]> = [];

	for (const [key, idx] of styleMap.entries()) {
		if (idx === removedIndex) {
			styleMap.delete(key);
		} else if (idx > removedIndex) {
			updates.push([key, idx - 1]);
		}
	}

	for (const [key, newIdx] of updates) {
		styleMap.set(key, newIdx);
	}
};
