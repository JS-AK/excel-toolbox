/**
 * Escapes special characters in a string for use in an XML document.
 *
 * Replaces:
 * - `&` with `&amp;`
 * - `<` with `&lt;`
 * - `>` with `&gt;`
 * - `"` with `&quot;`
 * - `'` with `&apos;`
 *
 * @param str - The string to escape.
 * @returns The escaped string.
 */
export function escapeXml(str: string): string {
	return str
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;")
		.replace(/"/g, "&quot;")
		.replace(/'/g, "&apos;");
}
