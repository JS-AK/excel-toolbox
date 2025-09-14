import type { CellStyle } from "../types/index.js";

import type { WorkbookBuilder } from "../workbook-builder.js";

// import * as Helpers from "./helpers/index.js";

/**
 * Validates a cell style index and checks if the style exists in the workbook.
 * This function performs validation checks on the provided style index and
 * verifies that the style exists at the given index in the cellXfs array.
 *
 * @param this - The WorkbookBuilder instance
 * @param payload - Object containing the style to validate
 * @param payload.style - The cell style configuration with index
 *
 * @returns True if the style index is valid and the style exists
 *
 * @throws {Error} When styleIndex is invalid (not a number) or when style doesn't exist at the given index
 */
export function remove(
	this: WorkbookBuilder,
	payload: { style: CellStyle },
): true {
	const { style } = payload;

	const styleIndex = style.index;

	if (typeof styleIndex !== "number") {
		throw new Error("Invalid styleIndex: not a number");
	}

	if (styleIndex === 0) {
		throw new Error("Invalid styleIndex: 0 is the default style and cannot be removed");
	}

	const xf = this.cellXfs[styleIndex];

	if (!xf) {
		throw new Error(`Invalid styleIndex: style not found at index ${styleIndex}`);
	}

	return true;

	// let removedSomething = false;

	// // Get style parts before splice - indices are still valid
	// const xf = this.cellXfs[styleIndex];

	// if (xf) {
	// 	this.cellXfs.splice(styleIndex, 1);

	// 	// Fix: reindex styleMap after splice
	// 	Helpers.reindexStyleMapAfterRemoval({ removedIndex: styleIndex, styleMap: this.styleMap });

	// 	// Reindex style references in cells across all sheets
	// 	for (const sheet of this.sheets.values()) {
	// 		for (const row of sheet.rows.values()) {
	// 			for (const cell of row.cells.values()) {
	// 				if (cell.style?.index !== undefined && cell.style.index > styleIndex) {
	// 					cell.style.index -= 1;
	// 				}
	// 			}
	// 		}
	// 	}

	// 	removedSomething = true;
	// }

	// return removedSomething;
}
