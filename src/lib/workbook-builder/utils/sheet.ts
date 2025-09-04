import { columnIndexToLetter, columnLetterToIndex } from "../../template/utils/index.js";
import { MergeCell } from "../merge-cells/types.js";

const MAX_COLUMNS = 16384;
const MAX_ROWS = 1_048_576;

// Тип ячейки (Cell)
export interface CellData {
	value: string | number | boolean | null;
	type?: CellType; // s = shared string, n = number, b = boolean

	style?: CellStyle;
}

export type CellStyle = {
	index?: number;
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

export type CellType = "s" | "inlineStr" | "n" | "b" | "str" | "e";

/*
	s	Shared string (ссылка на sharedStrings.xml)	<v> содержит индекс строки в sharedStrings
	inlineStr	Inline string	Вложенный элемент <is><t>текст</t></is>
	str	Formula string result (формула)	Не для обычных текстов, а для результата формулы
	b	Boolean	<v> — 0 или 1
	e	Ошибка	<v> — код ошибки
	n	Number (число)	Нет атрибута t или t="n"
*/

type CellValue = string | number | boolean | null | undefined;

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
	addMerge(mergeCell: MergeCell): MergeCell;
	removeMerge(mergeCell: MergeCell): boolean;
	setCell(rowIndex: number, column: string | number, cell: CellData): void;
	getCell(rowIndex: number, column: string | number): CellData | undefined;
	removeCell(rowIndex: number, column: string | number): boolean;
}

// Фабрика для создания пустого листа
export function createSheet(
	name: string,
	fn: {
		addMerge: (mergeCell: MergeCell & { sheetName: string }) => MergeCell;
		removeMerge: (mergeCell: MergeCell & { sheetName: string }) => boolean;
		addOrGetStyle: (style: CellStyle, sheetName: string) => number;
		addSharedString: (str: string, sheetName: string) => number;
		cleanupUnused: boolean; // новая опция
		removeSharedStringRef: (strIdx: number, sheetName: string) => boolean;
		removeStyleRef: (style: CellStyle, sheetName: string) => boolean;
	},
): SheetData {
	const {
		addMerge,
		addOrGetStyle,
		addSharedString,
		cleanupUnused,
		removeMerge,
		removeSharedStringRef,
		removeStyleRef,
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

			if (cleanupUnused) {
				const oldCell = rows.get(rowIndex)?.cells.get(letterColumn);

				// Обработка ситуации если до этого была в ячейке shared string
				if (oldCell) {
					if (oldCell?.type === "s" && typeof oldCell.value === "number") {
						removeSharedStringRef(oldCell.value, name);
					}

					if (oldCell?.style && typeof oldCell.style.index === "number") {
						removeStyleRef(oldCell.style, name);
					}
				}
			}

			// Если не указан type — определить сам
			cell.type = detectCellType(cell.value, cell.type);

			// Обработка shared string
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

			if (cleanupUnused) {
				const oldCell = rows.get(rowIndex)?.cells.get(letterColumn);

				// Обработка ситуации если до этого была в ячейке shared string
				if (oldCell) {
					if (oldCell?.type === "s" && typeof oldCell.value === "number") {
						removeSharedStringRef(oldCell.value, name);
					}

					if (oldCell?.style && typeof oldCell.style.index === "number") {
						removeStyleRef(oldCell.style, name);
					}
				}
			}

			return rows.get(rowIndex)?.cells.delete(letterColumn) ?? false;
		},
	};
}

function isValidColumn(column: string): boolean {
	if (!/^[A-Z]+$/.test(column)) return false;

	const idx = columnLetterToIndex(column);

	return idx > 0 && idx <= MAX_COLUMNS;
}

function detectCellType(value: CellValue, explicitType?: CellType): CellType {
	if (explicitType) {
		return explicitType;
	}

	if (value === null || value === undefined) {
		// Для пустых ячеек можно считать числовым типом с пустым значением
		return "n";
	}

	if (typeof value === "number") {
		return "n";
	}

	if (typeof value === "boolean") {
		return "b";
	}

	if (typeof value === "string") {
		// Проверка: если строка начинается с "=" — это формула, можно вернуть "str"
		// Но формулы лучше обрабатывать отдельно
		// По умолчанию — считаем inlineStr
		return "inlineStr";
	}

	// На всякий случай fallback
	return "inlineStr";
}
