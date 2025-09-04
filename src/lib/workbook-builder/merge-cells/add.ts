import type { WorkbookBuilder } from "../workbook-builder.js";

import * as Helpers from "./helpers/index.js";
import type { MergeCell } from "./types.js";

/**
 * Adds a merge cell range to the specified sheet.
 *
 * @param this - WorkbookBuilder instance
 * @param payload - Merge cell data with sheet name
 * 
 * @returns The added merge cell object
 *
 * @throws Error if sheet is not found or merge intersects with existing merged cells
 */
export function add(
	this: WorkbookBuilder,
	payload: MergeCell & { sheetName: string },
): MergeCell {
	const { endCol, endRow, sheetName, startCol, startRow } = payload;

	if (!this.getSheet(sheetName)) {
		throw new Error("Sheet not found");
	}

	const merges = this.mergeCells.get(sheetName) ?? [];

	// Check for intersection with existing merge cells
	for (const m of merges) {
		if (Helpers.rangesEqual(m, payload)) {
			return m; // Already exists
		}

		if (Helpers.rangesIntersect(m, payload)) {
			throw new Error("Merge intersects existing merged cell");
		}
	}

	const merge = {
		endCol,
		endRow,
		startCol,
		startRow,
	};

	merges.push(merge);

	this.mergeCells.set(sheetName, merges);

	return merge;
}
