import type { MergeCell } from "../types.js";

/**
 * Checks if two merge cell ranges intersect with each other.
 *
 * @param a - First merge cell range to check
 * @param b - Second merge cell range to check
 *
 * @returns True if the ranges intersect (overlap), false otherwise
 */
export function rangesIntersect(a: MergeCell, b: MergeCell): boolean {
	return !(
		a.endRow < b.startRow ||
		a.startRow > b.endRow ||
		a.endCol < b.startCol ||
		a.startCol > b.endCol
	);
}
