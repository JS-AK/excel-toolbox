import { CellData } from "./cell-data.js";

/** Row representation. */
export interface RowData {
	cells: Map<string, CellData>; // key — for example, "A", "B"
}
