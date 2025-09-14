import { MergeCell } from "../../types/index.js";

/**
 * Compares two merge cell ranges for equality.
 *
 * @param a - First merge cell range to compare
 * @param b - Second merge cell range to compare
 *
 * @returns True if both ranges have identical start and end coordinates
 */
export function rangesEqual(a: MergeCell, b: MergeCell): boolean {
	return a.startRow === b.startRow &&
		a.endRow === b.endRow &&
		a.startCol === b.startCol &&
		a.endCol === b.endCol;
}
