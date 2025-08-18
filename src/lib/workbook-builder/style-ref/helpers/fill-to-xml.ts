import { CellStyle, XmlNode } from "../../utils/index.js";

export const fillToXml = (payload: {
	existingFills: XmlNode["children"];
	fill?: CellStyle["fill"];
}): XmlNode => {
	const { existingFills, fill } = payload;

	if (!existingFills?.length) {
		throw new Error("existingFills is empty");
	}

	if (!fill) return existingFills[0] as XmlNode;

	const patternType = fill.patternType ?? "solid";
	const children: XmlNode["children"] = [];

	const attrs = { patternType };
	const fillChildren: XmlNode["children"] = [];

	if (fill.fgColor) {
		const colorVal = fill.fgColor.startsWith("#") ? fill.fgColor.slice(1) : fill.fgColor;
		fillChildren.push({
			attrs: { rgb: colorVal },
			tag: "fgColor",
		});
	}

	if (fill.bgColor) {
		const colorVal = fill.bgColor.startsWith("#") ? fill.bgColor.slice(1) : fill.bgColor;
		fillChildren.push({
			attrs: { rgb: colorVal },
			tag: "bgColor",
		});
	}
	
	children.push({
		attrs,
		children: fillChildren,
		tag: "patternFill",
	});

	return {
		children,
		tag: "fill",
	};
};
