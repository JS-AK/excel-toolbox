import { XmlNode } from "../types/index.js";

/**
 * Creates a basic empty border XML node.
 *
 * @returns XML node representing an empty border with left, right, top, and bottom elements
 */
export const border = (): XmlNode => ({
	children: [
		{ tag: "left" },
		{ tag: "right" },
		{ tag: "top" },
		{ tag: "bottom" },
	],
	tag: "border",
});
