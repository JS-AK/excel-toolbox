import { CellStyle } from "../utils/sheet.js";
import { WorkbookBuilder } from "../workbook-builder.js";

import * as Helpers from "./helpers/index.js";

export function remove(
	this: WorkbookBuilder,
	payload: { style: CellStyle },
): boolean {
	const { style } = payload;

	const styleIndex = style.index;
	if (styleIndex === undefined || styleIndex === null) {
		throw new Error("Invalid styleIndex");
	}

	let removedSomething = false;

	// Снимем части до splice — индексы ещё валидны
	const xf = this.cellXfs[styleIndex];

	if (xf) {
		this.cellXfs.splice(styleIndex, 1);

		// почин: переиндексация styleMap после splice
		Helpers.reindexStyleMapAfterRemoval({ removedIndex: styleIndex, styleMap: this.styleMap });

		// переиндексация ссылок в ячейках на всех листах
		for (const sheet of this.sheets.values()) {
			for (const row of sheet.rows.values()) {
				for (const cell of row.cells.values()) {
					if (cell.style?.index !== undefined && cell.style.index > styleIndex) {
						cell.style.index -= 1;
					}
				}
			}
		}

		removedSomething = true;
	}

	return removedSomething;
}
