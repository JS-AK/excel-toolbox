import type { WorkbookBuilder } from "../workbook-builder.js";

/**
 * Removes all shared string references for a specific sheet and cleans up unused strings.
 *
 * @param this - WorkbookBuilder instance
 * @param payload - Object containing the sheet name to remove references from
 */
export function removeAllFromSheet(
	this: WorkbookBuilder,
	payload: {
		sheetName: string;
	},
): void {
	const { sheetName } = payload;

	// 1. Collect strings that need to be removed
	const stringsToRemove: string[] = [];

	for (const [str, sheetsSet] of this.sharedStringRefs) {
		sheetsSet.delete(sheetName);
		if (sheetsSet.size === 0) {
			stringsToRemove.push(str);
		}
	}

	if (stringsToRemove.length === 0) {
		return;
	}

	// 2. Build map of old indices → new indices
	const oldToNew = new Map<number, number>();
	let newIdx = 0;

	for (let oldIdx = 0; oldIdx < this.sharedStrings.length; oldIdx++) {
		const str = this.sharedStrings[oldIdx];
		if (!str) {
			continue; // Skip if undefined
		}
		if (stringsToRemove.includes(str)) {
			// Remove string from refs
			this.sharedStringRefs.delete(str);
			continue; // Index is not accounted for
		}
		oldToNew.set(oldIdx, newIdx++);
	}

	// 3. Update sharedStrings array
	this.sharedStrings = this.sharedStrings.filter(s => !stringsToRemove.includes(s));

	// 4. Update indices in cells across all sheets
	for (const sheet of this.sheets.values()) {
		for (const row of sheet.rows.values()) {
			for (const cell of row.cells.values()) {
				if (cell.type === "s" && typeof cell.value === "number") {
					const newIdx = oldToNew.get(cell.value);
					if (newIdx !== undefined) {
						cell.value = newIdx;
					} else {
						// If cell.value was a removed string, set to 0 or null
						cell.value = 0;
					}
				}
			}
		}
	}
}
