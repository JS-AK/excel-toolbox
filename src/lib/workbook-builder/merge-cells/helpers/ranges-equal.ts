import { MergeCell } from "../types.js";

export function rangesEqual(a: MergeCell, b: MergeCell): boolean {
	return a.startRow === b.startRow &&
		a.endRow === b.endRow &&
		a.startCol === b.startCol &&
		a.endCol === b.endCol;
}
