import * as Default from "../../default/index.js";

import type { CellStyle, XmlNode } from "../../types/index.js";

/**
 * Builds an XmlNode representing a <font> element based on the provided CellStyle.
 *
 * If no font is provided, returns the default font from the template. Color
 * values may include a leading '#'. Six-digit RGB values are prefixed with
 * 'FF' (alpha) to form ARGB; eight-digit values are used as-is. Throws for
 * invalid color lengths.
 *
 * @param payload - Input arguments
 * @param payload.existingFonts - Existing fonts collection (not used directly; kept for parity)
 * @param payload.font - Optional font from the cell style
 * @returns XmlNode representing a <font> element
 */
export const fontToXml = (payload: {
	existingFonts: XmlNode["children"];
	font?: CellStyle["font"];
}): XmlNode => {
	const { existingFonts, font } = payload;

	if (!existingFonts?.length) {
		throw new Error("existingFonts is empty");
	}

	// Default values
	const defaultFont = Default.font();

	if (!font) {
		// If no font style provided — return default font (as in Excel template)
		return defaultFont;
	}

	const children: XmlNode[] = [];

	// Size is always present
	children.push({
		attrs: { val: String(font.size ?? defaultFont.children.at(0)?.attrs.val) },
		tag: "sz",
	});

	// Color (if provided) otherwise default
	if (font.color) {
		const colorVal = font.color.startsWith("#") ? font.color.slice(1) : font.color;
		if (colorVal.length === 6) {
			children.push({
				attrs: { rgb: "FF" + colorVal.toUpperCase() },
				tag: "color",
			});
		} else if (colorVal.length === 8) {
			children.push({
				attrs: { rgb: colorVal.toUpperCase() },
				tag: "color",
			});
		} else {
			throw new Error(`Invalid font color: ${font.color}`);
		}
	} else {
		children.push(defaultFont.children.at(1) as XmlNode);
	}

	// Font name (required)
	children.push({
		attrs: { val: font.name ?? defaultFont.children.at(2)?.attrs.val },
		tag: "name",
	});

	// Additional attributes
	if (font.bold) children.push({ tag: "b" });
	if (font.italic) children.push({ tag: "i" });
	if (font.underline) {
		const val = font.underline === true ? "single" : font.underline;
		children.push({
			attrs: { val },
			tag: "u",
		});
	}

	return {
		children,
		tag: "font",
	};
};
