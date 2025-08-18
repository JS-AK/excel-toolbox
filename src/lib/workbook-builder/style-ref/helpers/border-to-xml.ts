import { CellStyle, XmlNode } from "../../utils/index.js";

export const borderToXml = (payload: {
	border?: CellStyle["border"];
	existingBorders: XmlNode["children"];
}): XmlNode => {
	const { border, existingBorders } = payload;

	if (!existingBorders?.length) {
		throw new Error("existingBorders is empty");
	}

	if (!border) return existingBorders[0] as XmlNode;

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
