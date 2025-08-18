export const addUnique = (arr: unknown[], item: unknown): number => {
	const idx = arr.findIndex(x => JSON.stringify(x) === JSON.stringify(item));

	if (idx !== -1) return idx;

	arr.push(item);

	return arr.length - 1;
};
