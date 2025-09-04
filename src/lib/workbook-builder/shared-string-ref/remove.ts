import type { WorkbookBuilder } from "../workbook-builder.js";

/**
 * Removes a shared string reference for a specific sheet and cleans up if unused.
 *
 * @param this - WorkbookBuilder instance
 * @param payload - Object containing sheet name and string index to remove
 *
 * @returns True if the reference was successfully removed, false if string or reference not found
 */
export function remove(
	this: WorkbookBuilder,
	payload: {
		sheetName: string;
		strIdx: number;
	},
): boolean {
	const { sheetName, strIdx } = payload;

	const str = this.sharedStrings[strIdx];
	if (!str) {
		return false;
	}

	const refs = this.sharedStringRefs.get(str);
	if (!refs) {
		return false;
	}

	refs.delete(sheetName);

	if (refs.size === 0) {
		// Build map of old indices → new indices before removal
		const oldToNew = new Map<number, number>();
		for (let i = 0; i < this.sharedStrings.length; i++) {
			if (i < strIdx) {
				oldToNew.set(i, i);
			} else if (i > strIdx) {
				oldToNew.set(i, i - 1);
			}
			// i === strIdx — this string will be removed, no index
		}

		// Remove string from array and refs
		this.sharedStrings.splice(strIdx, 1);
		this.sharedStringRefs.delete(str);

		// Update indices across all sheets
		for (const sheet of this.sheets.values()) {
			for (const row of sheet.rows.values()) {
				for (const cell of row.cells.values()) {
					if (cell.type === "s" && typeof cell.value === "number") {
						const newIdx = oldToNew.get(cell.value);
						if (newIdx !== undefined) {
							cell.value = newIdx;
						} else {
							// Just in case, if cell.value was a removed string
							cell.value = 0; // or null, according to your application logic
						}
					}
				}
			}
		}
	}

	return true;
}
