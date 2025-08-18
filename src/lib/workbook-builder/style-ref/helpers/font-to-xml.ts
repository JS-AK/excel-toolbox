import { CellStyle, XmlNode } from "../../utils/index.js";

export const fontToXml = (payload: {
	existingFonts: XmlNode["children"];
	font?: CellStyle["font"];
}): XmlNode => {
	const { existingFonts, font } = payload;

	if (!existingFonts?.length) {
		throw new Error("existingFonts is empty");
	}

	if (!font) return existingFonts[0] as XmlNode;

	const children = [];
	if (font.size) children.push({
		attrs: { val: String(font.size) },
		tag: "sz",
	});
	if (font.color) {
		const colorVal = font.color.startsWith("#") ? font.color.slice(1) : font.color;
		if (colorVal.length === 6) {
			children.push({
				attrs: { rgb: "FF" + colorVal.toUpperCase() }, // добавляем FF - непрозрачность
				tag: "color",
			});
		} else if (colorVal.length === 8) {
			children.push({
				attrs: { rgb: colorVal.toUpperCase() },
				tag: "color",
			});
		} else {
			throw new Error(`Некорректный цвет: ${font.color}`);
		}
	}
	if (font.name) children.push({
		attrs: { val: font.name },
		tag: "name",
	});
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
