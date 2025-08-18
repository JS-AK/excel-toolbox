import { CellStyle } from "../utils/sheet.js";
import { WorkbookBuilder } from "../workbook-builder.js";

import * as Helpers from "./helpers/index.js";

export function addOrGet(
	this: WorkbookBuilder,
	payload: { style: CellStyle },
): number {
	const { style } = payload;

	// Конвертируем каждую часть
	const fontId = Helpers.addUnique(
		this.fonts,
		Helpers.fontToXml({ existingFonts: this.fonts, font: style.font }),
	);
	const fillId = Helpers.addUnique(
		this.fills,
		Helpers.fillToXml({ existingFills: this.fills, fill: style.fill }),
	);
	const borderId = Helpers.addUnique(
		this.borders,
		Helpers.borderToXml({ border: style.border, existingBorders: this.borders }),
	);
	const numFmtId = style.numberFormat
		? Helpers.addNumFmt({ formatCode: style.numberFormat, numFmts: this.numFmts })
		: 0;

	const xf = {
		alignment: style.alignment,
		borderId,
		fillId,
		fontId,
		numFmtId,
	};

	const xfKey = JSON.stringify(xf);

	if (this.styleMap.has(xfKey)) {
		return this.styleMap.get(xfKey)!;
	}

	const index = this.cellXfs.length;

	this.cellXfs.push(xf);

	this.styleMap.set(xfKey, index);

	return index;
};
