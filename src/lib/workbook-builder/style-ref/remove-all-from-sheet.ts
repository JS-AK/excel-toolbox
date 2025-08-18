import { CellStyle } from "../utils/sheet.js";
import { WorkbookBuilder } from "../workbook-builder.js";

import { remove } from "./remove.js";

export function removeAllFromSheet(
	this: WorkbookBuilder,
	payload: {
		sheetName: string;
	},
) {
	const { sheetName } = payload;

	const sheet = this.sheets.get(sheetName);
	if (!sheet) return false;

	const stylesToRemove: CellStyle[] = [];

	let removedSomething = false;

	for (const row of sheet.rows.values()) {
		for (const cell of row.cells.values()) {
			if (cell.style?.index !== undefined) {

				stylesToRemove.push(cell.style);
			}
		}
	}

	for (const style of stylesToRemove) {
		const removed = remove.bind(this)({ style });

		if (removed) {
			removedSomething = true;
		}
	}

	return removedSomething;
}
