import type { WorkbookBuilder } from "../workbook-builder.js";

import { MergeCell } from "../types/index.js";

import * as Helpers from "./helpers/index.js";

/**
 * Removes a merge cell range from the specified sheet.
 *
 * @param this - WorkbookBuilder instance
 * @param payload - Merge cell data with sheet name
 *
 * @returns True if the merge cell was successfully removed
 *
 * @throws Error if sheet is not found or merge cell does not exist
 */
export function remove(
	this: WorkbookBuilder,
	payload: MergeCell & { sheetName: string },
): boolean {
	const { endCol, endRow, sheetName, startCol, startRow } = payload;

	if (!this.getSheet(sheetName)) {
		throw new Error("Sheet not found: " + sheetName);
	}

	const merges = this.mergeCells.get(sheetName) ?? [];

	const i = merges.findIndex(m => Helpers.rangesEqual(m, { endCol, endRow, startCol, startRow }));

	if (i === -1) {
		throw new Error("Sheet: " + sheetName + " Invalid merge cell: " + JSON.stringify(payload));
	}

	merges.splice(i, 1);

	return true;
}
