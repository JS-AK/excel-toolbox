/**
 * Recursively checks if an object contains any arrays in its values.
 *
 * @param {Record<string, unknown>} replacements - The object to check for arrays.
 * @returns {boolean} True if any arrays are found, false otherwise.
 */

export function foundArraysInReplacements(replacements: Record<string, unknown>): boolean {
	let isFound = false;

	for (const key in replacements) {
		const value = replacements[key];

		if (Array.isArray(value)) {
			isFound = true;

			return isFound;
		}

		if (typeof value === "object" && value !== null) {
			isFound = foundArraysInReplacements(value as Record<string, unknown>);

			if (isFound) {
				return isFound;
			}
		}
	}

	return isFound;
}
