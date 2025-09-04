/**
 * Creates a default font XML node with Calibri font.
 *
 * @returns XML node representing a default font configuration
 */
export const font = () => ({
	children: [
		{ attrs: { val: "11" }, tag: "sz" },
		{ attrs: { theme: "1" }, tag: "color" },
		{ attrs: { val: "Calibri" }, tag: "name" },
		{ attrs: { val: "2" }, tag: "family" },
		{ attrs: { val: "minor" }, tag: "scheme" },
	],
	tag: "font",
});
