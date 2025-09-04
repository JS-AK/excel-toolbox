import { MergeCell } from "../types.js";

export function rangesIntersect(a: MergeCell, b: MergeCell): boolean {
	return !(
		a.endRow < b.startRow ||
		a.startRow > b.endRow ||
		a.endCol < b.startCol ||
		a.startCol > b.endCol
	);
}
