import type { CellStyle, XmlNode } from "../../types/index.js";

/**
 * Builds an XmlNode representing a <border> element based on the provided CellStyle.
 *
 * If no border is provided, the first existing border is returned (assumed default).
 * Color values beginning with # are converted to RGB without the leading hash, as expected by Excel.
 *
 * @param payload - Input arguments
 * @param payload.border - Optional border style from the cell style
 * @param payload.existingBorders - Existing borders collection to fall back to
 * @returns XmlNode representing a <border> element
 */
export const borderToXml = (payload: {
	border?: CellStyle["border"];
	existingBorders: XmlNode["children"];
}): XmlNode => {
	const { border, existingBorders } = payload;

	if (!existingBorders?.length) {
		throw new Error("existingBorders is empty");
	}

	if (!border) {
		return existingBorders[0] as XmlNode;
	}

	const children: XmlNode["children"] = [];

	for (const side of ["left", "right", "top", "bottom"] as const) {
		const b = border?.[side];

		if (b) {
			const attrs = { style: b.style };
			const sideChildren = b.color
				? [{
					attrs: { rgb: b.color.replace(/^#/, "") },
					tag: "color",
				}]
				: [];
			children.push({
				attrs,
				children: sideChildren,
				tag: side,
			});
		} else {
			children.push({ tag: side });
		}
	}
	return {
		children,
		tag: "border",
	};
};
