import type { WorkbookBuilder } from "../workbook-builder.js";

/**
 * Adds a shared string to the workbook and returns its index.
 *
 * @param this - WorkbookBuilder instance
 * @param payload - Object containing sheet name and string value
 *
 * @returns The index of the shared string in the shared strings array
 */
export function add(
	this: WorkbookBuilder,
	payload: {
		sheetName: string;
		str: string;
	},
): number {
	const { sheetName, str } = payload;

	let idx = this.sharedStrings.indexOf(str);

	if (idx === -1) {
		idx = this.sharedStrings.length;
		this.sharedStrings.push(str);
		this.sharedStringRefs.set(str, new Set([sheetName]));
	} else {
		// Add sheet name to Set if not already present
		this.sharedStringRefs.get(str)?.add(sheetName);
	}

	return idx;
}
