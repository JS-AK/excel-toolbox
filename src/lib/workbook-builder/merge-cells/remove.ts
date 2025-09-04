import { WorkbookBuilder } from "../workbook-builder.js";

import * as Helpers from "./helpers/index.js";
import { MergeCell } from "./types.js";

export function remove(
	this: WorkbookBuilder,
	payload: MergeCell & { sheetName: string },
): boolean {
	const { endCol, endRow, sheetName, startCol, startRow } = payload;

	if (!this.getSheet(sheetName)) {
		throw new Error("Sheet not found");
	}

	const merges = this.mergeCells.get(sheetName) ?? [];

	const i = merges.findIndex(m => Helpers.rangesEqual(m, { endCol, endRow, startCol, startRow }));

	if (i === -1) throw new Error("Invalid merge cell");

	merges.splice(i, 1);

	return true;
}
