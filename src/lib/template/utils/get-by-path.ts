/**
 * Gets a value from an object by a given path.
 *
 * @param obj - The object to search.
 * @param path - The path to the value, separated by dots.
 * @returns The value at the given path, or undefined if not found.
 */
export function getByPath(obj: unknown, path: string): unknown {
	return path.split(".").reduce((acc, key) => {
		if (acc && typeof acc === "object" && key in acc) {
			return (acc as Record<string, unknown>)[key];
		}
		return undefined;
	}, obj);
}
