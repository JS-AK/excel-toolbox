import { XmlNode } from "../types/index.js";

/**
 * Creates a basic fill XML node with no pattern.
 *
 * @returns XML node representing an empty pattern fill
 */
export const fill = (): XmlNode => ({
	children: [
		{
			attrs: { patternType: "none" },
			tag: "patternFill",
		},
	],
	tag: "fill",
});
