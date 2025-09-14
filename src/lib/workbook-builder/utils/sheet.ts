import { columnIndexToLetter, columnLetterToIndex } from "../../template/utils/index.js";

import type { CellStyle, CellType, CellValue, MergeCell, RowData, SheetData } from "../types/index.js";

import { dateToExcelSerial } from "./date-to-excel-serial.js";

/** Maximum number of columns supported by Excel (XFD). */
const MAX_COLUMNS = 16384;
/** Maximum number of rows supported by Excel (1,048,576). */
const MAX_ROWS = 1_048_576;

/**
 * Factory for creating a new sheet with bound helpers.
 *
 * @param name - Sheet name
 * @param fn.addMerge - Function to add a merge range (bound to workbook)
 * @param fn.removeMerge - Function to remove a merge range (bound to workbook)
 * @param fn.addOrGetStyle - Function to add or get a style index
 * @param fn.addSharedString - Function to add a shared string and return its index
 * @returns SheetData instance with helpers
 */
export function createSheet(
	name: string,
	fn: {
		addMerge: (mergeCell: MergeCell & { sheetName: string }) => MergeCell;
		removeMerge: (mergeCell: MergeCell & { sheetName: string }) => boolean;
		addOrGetStyle: (style: CellStyle, sheetName: string) => number;
		addSharedString: (str: string, sheetName: string) => number;
	},
): SheetData {
	const {
		addMerge,
		addOrGetStyle,
		addSharedString,
		removeMerge,
	} = fn;
	const rows = new Map<number, RowData>();

	return {
		name,
		rows,

		addMerge(mergeCell: MergeCell): MergeCell {
			return addMerge({ ...mergeCell, sheetName: name });
		},

		removeMerge(mergeCell: MergeCell) {
			return removeMerge({ ...mergeCell, sheetName: name });
		},

		setCell(rowIndex, column, cell) {
			if (rowIndex <= 0) {
				throw new Error("Invalid rowIndex");
			}

			if (!Number.isInteger(rowIndex) || rowIndex <= 0) {
				throw new Error("Invalid rowIndex: must be a positive integer");
			}

			if (rowIndex > MAX_ROWS) {
				throw new Error(`Invalid rowIndex: exceeds Excel max rows (${MAX_ROWS})`);
			}

			if (!rows.has(rowIndex)) {
				rows.set(rowIndex, { cells: new Map() });
			}

			const letterColumn = typeof column === "number"
				? columnIndexToLetter(column)
				: column;

			if (!isValidColumn(letterColumn)) {
				throw new Error(`Invalid column string: "${letterColumn}"`);
			}

			// if is Date
			if (cell.value instanceof Date) {
				cell.value = dateToExcelSerial(cell.value);
			}

			if (cell.isFormula) {
				cell.type = undefined;
			} else {
				if (cell.type === "str") {
					throw new Error(`Cell type: "${cell.type}" valid only for formula cells`);
				}

				// If type is not provided — detect automatically
				cell.type = detectCellType(cell.value, cell.type);
			}

			// Handle shared string
			if (cell.type === "s") {
				const idx = addSharedString(String(cell.value ?? ""), name);

				cell = { ...cell, value: idx };
			}

			if (cell.style) {
				const styleIndex = addOrGetStyle(cell.style, name);

				cell.style.index = styleIndex;
			}

			rows.get(rowIndex)?.cells.set(letterColumn, cell);
		},

		getCell(rowIndex, column) {
			if (typeof column === "number") {
				if (column < 0 || column > MAX_COLUMNS) {
					throw new Error("Invalid column number");
				}

				return rows.get(rowIndex)?.cells.get(columnIndexToLetter(column));
			} else {
				if (!isValidColumn(column)) {
					throw new Error(`Invalid column string: "${column}"`);
				}

				return rows.get(rowIndex)?.cells.get(column);
			}
		},

		removeCell(rowIndex, column) {
			if (rowIndex <= 0) {
				throw new Error("Invalid rowIndex");
			}

			if (!Number.isInteger(rowIndex) || rowIndex <= 0) {
				throw new Error("Invalid rowIndex: must be a positive integer");
			}

			if (rowIndex > MAX_ROWS) {
				throw new Error(`Invalid rowIndex: exceeds Excel max rows (${MAX_ROWS})`);
			}

			const letterColumn = typeof column === "number"
				? columnIndexToLetter(column)
				: column;

			if (!isValidColumn(letterColumn)) {
				throw new Error(`Invalid column string: "${letterColumn}"`);
			}

			return rows.get(rowIndex)?.cells.delete(letterColumn) ?? false;
		},
	};
}

/** Validates an Excel column string (A-Z, AA, ..., XFD). */
function isValidColumn(column: string): boolean {
	if (!/^[A-Z]+$/.test(column)) return false;

	const idx = columnLetterToIndex(column);

	return idx > 0 && idx <= MAX_COLUMNS;
}

/**
 * Detects the appropriate cell type based on the value when not explicitly specified.
 *
 * @param value - Cell value
 * @param explicitType - Explicitly provided type, if any
 * @returns CellType inferred or the explicit type
 */
function detectCellType(value: CellValue, explicitType?: CellType): CellType {
	if (explicitType) {
		return explicitType;
	}

	if (value === null || value === undefined) {
		// For empty cells we default to numeric type with empty value
		return "n";
	}

	if (typeof value === "number") {
		return "n";
	}

	if (typeof value === "boolean") {
		return "b";
	}

	if (typeof value === "string") {
		// Default to inlineStr for plain strings
		return "inlineStr";
	}

	// Fallback to inlineStr
	return "inlineStr";
}
