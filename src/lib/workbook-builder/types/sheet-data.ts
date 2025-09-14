import { CellData } from "./cell-data.js";
import { MergeCell } from "./merge-cell.js";
import { RowData } from "./row-data.js";

/** Sheet representation with convenience methods. */
export interface SheetData {
	name: string;
	rows: Map<number, RowData>;

	addMerge(mergeCell: MergeCell): MergeCell;
	removeMerge(mergeCell: MergeCell): boolean;
	setCell(rowIndex: number, column: string | number, cell: CellData): void;
	getCell(rowIndex: number, column: string | number): CellData | undefined;
	removeCell(rowIndex: number, column: string | number): boolean;
}
