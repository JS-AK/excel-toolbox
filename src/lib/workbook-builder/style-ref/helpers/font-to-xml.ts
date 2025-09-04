import * as Default from "../../default/index.js";
import { CellStyle, XmlNode } from "../../utils/index.js";

export const fontToXml = (payload: {
	existingFonts: XmlNode["children"];
	font?: CellStyle["font"];
}): XmlNode => {
	const { existingFonts, font } = payload;

	if (!existingFonts?.length) {
		throw new Error("existingFonts is empty");
	}

	// значения по умолчанию
	const defaultFont = Default.font();

	if (!font) {
		// если стиля нет — возвращаем "нулевой" шрифт (как в шаблоне Excel)
		return defaultFont;
	}

	const children: XmlNode[] = [];

	// размер всегда должен быть
	children.push({
		attrs: { val: String(font.size ?? defaultFont.children.at(0)?.attrs.val) },
		tag: "sz",
	});

	// цвет (если есть) иначе дефолтный
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
			throw new Error(`Некорректный цвет: ${font.color}`);
		}
	} else {
		children.push(defaultFont.children.at(1) as XmlNode);
	}

	// имя шрифта (обязателен)
	children.push({
		attrs: { val: font.name ?? defaultFont.children.at(2)?.attrs.val },
		tag: "name",
	});

	// доп. атрибуты
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
