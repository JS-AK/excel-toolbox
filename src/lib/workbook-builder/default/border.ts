/**
 * Creates a basic empty border XML node.
 *
 * @returns XML node representing an empty border with left, right, top, and bottom elements
 */
export const border = () => ({
	children: [
		{ tag: "left" },
		{ tag: "right" },
		{ tag: "top" },
		{ tag: "bottom" },
	],
	tag: "border",
});
