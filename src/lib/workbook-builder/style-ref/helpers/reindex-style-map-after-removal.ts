export const reindexStyleMapAfterRemoval = (payload: {
	removedIndex: number;
	styleMap: Map<string, number>;
}) => {
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
