/**
 * Adds an item to the array if it does not already exist and returns its index.
 *
 * Equality is determined using JSON.stringify on items.
 *
 * Note: This works well for simple JSON-serializable values. For complex
 * structures with functions or circular references, provide a custom equality
 * strategy instead of using this helper.
 *
 * @param arr - Target array to search or extend
 * @param item - Item to ensure uniqueness for
 * @returns Index of the existing or newly appended item
 */
export const addUnique = <T>(
	arr: T[],
	item: T,
): number => {
	const idx = arr.findIndex(x => JSON.stringify(x) === JSON.stringify(item));

	if (idx !== -1) {
		return idx;
	}

	arr.push(item);

	return arr.length - 1;
};
