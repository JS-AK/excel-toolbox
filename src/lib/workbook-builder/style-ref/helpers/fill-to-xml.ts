import type { CellStyle, XmlNode } from "../../types/index.js";

/**
 * Builds an XmlNode representing a <fill> element based on the provided CellStyle.
 *
 * If no fill is provided, the first existing fill is returned (assumed default).
 * Foreground/background colors may include a leading '#', which is removed to
 * match Excel's expected RGB format.
 *
 * @param payload - Input arguments
 * @param payload.existingFills - Existing fills collection to fall back to
 * @param payload.fill - Optional fill from the cell style
 * @returns XmlNode representing a <fill> element
 */
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
