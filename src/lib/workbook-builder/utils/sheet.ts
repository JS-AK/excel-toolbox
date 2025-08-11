import { columnIndexToLetter, columnLetterToIndex } from "../../template/utils/index.js";

const MAX_COLUMNS = 16384;

// Тип ячейки (Cell)
export interface CellData {
	value: string | number | boolean | null;
	type?: CellType; // s = shared string, n = number, b = boolean

	style?: {
		font?: {
			name?: string;           // Название шрифта, например "Calibri"
			size?: number;           // Размер шрифта, например 11, 14
			bold?: boolean;          // Жирный
			italic?: boolean;        // Курсив
			underline?: boolean | "single" | "double"; // Подчеркивание
			color?: string;          // Цвет текста в формате HEX, например "#FF0000"
		};
		fill?: {
			type?: "pattern";        // Пока поддерживаем patternFill
			patternType?: string;    // Например "solid", "gray125", "none"
			fgColor?: string;        // Цвет переднего плана (заливка) — HEX или ARGB
			bgColor?: string;        // Цвет фона (редко используется)
		};
		border?: {
			top?: BorderStyle;
			bottom?: BorderStyle;
			left?: BorderStyle;
			right?: BorderStyle;
		};
		alignment?: {
			horizontal?: "left" | "center" | "right" | "justify";
			vertical?: "top" | "center" | "bottom";
			wrapText?: boolean;
			indent?: number;
		};
		numberFormat?: string;     // Формат числа, например "0.00", "dd/mm/yyyy", "$#,##0.00"
	};
}

type CellType = "s" | "n" | "b" | "str" | "e";

type BorderStyle = {
	style: "thin" | "medium" | "thick" | "dashed" | "dotted" | "double" | "hair" | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot";
	color?: string;            // Цвет бордера (HEX или ARGB)
};

// Тип строки (Row)
export interface RowData {
	cells: Map<string, CellData>; // ключ — например, "A1", "B1"
}

// Тип листа (Sheet)
export interface SheetData {
	name: string;
	rows: Map<number, RowData>;

	// Методы для удобной работы
	setCell(rowIndex: number, column: string | number, cell: CellData): void;
	getCell(rowIndex: number, column: string | number): CellData | undefined;
	deleteCell(rowIndex: number, column: string | number): boolean;
}

// Фабрика для создания пустого листа
export function createSheet(
	name: string,
	addSharedString: (str: string, sheetName: string) => number,
	removeSharedStringRef: (strIdx: number, sheetName: string) => boolean,
): SheetData {
	const rows = new Map<number, RowData>();

	return {
		name,
		rows,

		setCell(rowIndex, column, cell) {
			if (rowIndex <= 0) {
				throw new Error("Invalid rowIndex");
			}

			if (!rows.has(rowIndex)) {
				rows.set(rowIndex, { cells: new Map() });
			}

			if (typeof column === "number") {
				if (column < 0 || column > MAX_COLUMNS) {
					throw new Error("Invalid column number");
				}

				const oldCell = rows.get(rowIndex)?.cells.get(columnIndexToLetter(column));

				// Обработка ситуации если до этого была в ячейке shared string
				if (oldCell) {
					if (oldCell?.type === "s" && typeof oldCell.value === "number") {
						removeSharedStringRef(oldCell.value, name);
					}
				}

				// Обработка shared string
				if (cell.type === "s") {
					const idx = addSharedString(String(cell.value ?? ""), name);

					cell = { type: cell.type, value: idx };
				}

				rows.get(rowIndex)?.cells.set(columnIndexToLetter(column), cell);
			} else {
				if (!isValidColumn(column)) {
					throw new Error(`Invalid column string: "${column}"`);
				}

				// Обработка shared string
				if (cell.type === "s") {
					const idx = addSharedString(String(cell.value ?? ""), name);

					cell = { type: cell.type, value: idx };
				}

				rows.get(rowIndex)?.cells.set(column, cell);
			}
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

		deleteCell(rowIndex, column) {
			if (typeof column === "number") {
				if (column < 0 || column > MAX_COLUMNS) {
					throw new Error("Invalid column number");
				}

				return rows.get(rowIndex)?.cells.delete(columnIndexToLetter(column)) ?? false;
			} else {
				if (!isValidColumn(column)) {
					throw new Error(`Invalid column string: "${column}"`);
				}

				return rows.get(rowIndex)?.cells.delete(column) ?? false;
			}
		},
	};
}

function isValidColumn(column: string): boolean {
	if (!/^[A-Z]+$/.test(column)) return false;

	const idx = columnLetterToIndex(column);

	return idx > 0 && idx <= MAX_COLUMNS;
}
