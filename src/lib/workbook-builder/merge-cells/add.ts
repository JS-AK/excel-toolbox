import { WorkbookBuilder } from "../workbook-builder.js";

import * as Helpers from "./helpers/index.js";
import { MergeCell } from "./types.js";

export function add(
	this: WorkbookBuilder,
	payload: MergeCell & { sheetName: string },
): MergeCell {
	const { endCol, endRow, sheetName, startCol, startRow } = payload;

	if (!this.getSheet(sheetName)) {
		throw new Error("Sheet not found");
	}

	const merges = this.mergeCells.get(sheetName) ?? [];

	// Проверка пересечения
	for (const m of merges) {
		if (Helpers.rangesEqual(m, payload)) {
			return m; // уже есть
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
