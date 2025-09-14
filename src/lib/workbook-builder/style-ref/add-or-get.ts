import type { CellStyle } from "../types/index.js";

import type { WorkbookBuilder } from "../workbook-builder.js";

import * as Helpers from "./helpers/index.js";

/**
 * Adds a new cell style to the workbook or returns the index of an existing identical style.
 * This function manages style deduplication by checking if an identical style already exists
 * before creating a new one.
 *
 * @param this - The WorkbookBuilder instance
 * @param payload - Object containing the style to add or get
 * @param payload.style - The cell style configuration
 *
 * @returns The index of the style in the cellXfs array
 */
export function addOrGet(
	this: WorkbookBuilder,
	payload: { style: CellStyle },
): number {
	const { style } = payload;

	// Convert each style component to XML and get their IDs
	// Helpers.addUnique() add unique item to array and return index of item

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
	} as const;

	const xfKey = JSON.stringify(xf);

	if (this.styleMap.has(xfKey)) {
		return this.styleMap.get(xfKey)!;
	}

	const index = this.cellXfs.length;

	this.cellXfs.push(xf);

	this.styleMap.set(xfKey, index);

	return index;
}
