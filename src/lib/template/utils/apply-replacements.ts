import { getByPath } from "./get-by-path.js";

/**
 * Replaces placeholders in the given content string with values from the replacements map.
 *
 * The function searches for placeholders in the format `${key}` within the content
 * string, where `key` corresponds to a path in the replacements object.
 * If a value is found for the key, it replaces the placeholder with the value.
 * If no value is found, the original placeholder remains unchanged.
 *
 * @param content - The string containing placeholders to be replaced.
 * @param replacements - An object where keys represent placeholder paths and values are the replacements.
 * @returns A new string with placeholders replaced by corresponding values from the replacements object.
 */

export const applyReplacements = (content: string, replacements: Record<string, unknown>): string => {
	if (!content) {
		return "";
	}

	return content.replace(/\$\{([^}]+)\}/g, (match, path) => {
		const value = getByPath(replacements, path);

		return value !== undefined ? String(value) : match;
	});
};
