interface ValidationResult {
	isValid: boolean;
	error?: {
		message: string;
		details?: string;
	};
}

export function validateWorksheetXml(xml: string): ValidationResult {
	const createError = (message: string, details?: string): ValidationResult => ({
		error: {
			details,
			message,
		},
		isValid: false,
	});

	// 1. Проверка базовой структуры XML
	if (!xml.startsWith("<?xml")) {
		return createError("XML должен начинаться с декларации <?xml>");
	}

	if (!xml.includes("<worksheet") || !xml.includes("</worksheet>")) {
		return createError("Не найден корневой элемент worksheet");
	}

	// 2. Проверка наличия обязательных элементов
	const requiredElements = [
		{ name: "sheetViews", tag: "<sheetViews>" },
		{ name: "sheetFormatPr", tag: "<sheetFormatPr" },
		{ name: "cols", tag: "<cols>" },
		{ name: "sheetData", tag: "<sheetData>" },
		{ name: "mergeCells", tag: "<mergeCells" },
	];

	for (const { name, tag } of requiredElements) {
		if (!xml.includes(tag)) {
			return createError(`Отсутствует обязательный элемент ${name}`);
		}
	}

	// 3. Извлечение и проверка sheetData
	const sheetDataStart = xml.indexOf("<sheetData>");
	const sheetDataEnd = xml.indexOf("</sheetData>");
	if (sheetDataStart === -1 || sheetDataEnd === -1) {
		return createError("Некорректная структура sheetData");
	}

	const sheetDataContent = xml.substring(sheetDataStart + 10, sheetDataEnd);
	const rows = sheetDataContent.split("</row>");

	if (rows.length < 2) {
		return createError("SheetData должен содержать хотя бы одну строку");
	}

	// Собираем информацию о всех строках и ячейках
	const allRows: number[] = [];
	const allCells: { row: number; col: string }[] = [];
	let prevRowNum = 0;

	for (const row of rows.slice(0, -1)) {
		if (!row.includes("<row ")) {
			return createError("Не найден тег row", `Фрагмент: ${row.substring(0, 50)}...`);
		}

		if (!row.includes("<c ")) {
			return createError("Строка не содержит ячеек", `Строка: ${row.substring(0, 50)}...`);
		}

		// Извлекаем номер строки
		const rowNumMatch = row.match(/<row\s+r="(\d+)"/);
		if (!rowNumMatch) {
			return createError("Не указан номер строки (атрибут r)", `Строка: ${row.substring(0, 50)}...`);
		}
		const rowNum = parseInt(rowNumMatch[1]!);

		// Проверка уникальности строк
		if (allRows.includes(rowNum)) {
			return createError("Найден дубликат номера строки", `Номер строки: ${rowNum}`);
		}
		allRows.push(rowNum);

		// Проверка порядка строк (должны идти по возрастанию)
		if (rowNum <= prevRowNum) {
			return createError(
				"Нарушен порядок следования строк",
				`Текущая строка: ${rowNum}, предыдущая: ${prevRowNum}`,
			);
		}
		prevRowNum = rowNum;

		// Извлекаем все ячейки в строке
		const cells = row.match(/<c\s+r="([A-Z]+)(\d+)"/g) || [];
		for (const cell of cells) {
			const match = cell.match(/<c\s+r="([A-Z]+)(\d+)"/);
			if (!match) {
				return createError("Некорректный формат ячейки", `Ячейка: ${cell}`);
			}

			const col = match[1]!;
			const cellRowNum = parseInt(match[2]!);

			// Проверяем соответствие номера строки
			if (cellRowNum !== rowNum) {
				return createError(
					"Несоответствие номера строки в ячейке",
					`Ожидалось: ${rowNum}, найдено: ${cellRowNum} в ячейке ${col}${cellRowNum}`,
				);
			}

			allCells.push({
				col,
				row: rowNum,
			});
		}
	}

	// 4. Проверка mergeCells
	const mergeCellsStart = xml.indexOf("<mergeCells");
	const mergeCellsEnd = xml.indexOf("</mergeCells>");
	if (mergeCellsStart === -1 || mergeCellsEnd === -1) {
		return createError("Некорректная структура mergeCells");
	}

	const mergeCellsContent = xml.substring(mergeCellsStart, mergeCellsEnd);
	const countMatch = mergeCellsContent.match(/count="(\d+)"/);
	if (!countMatch) {
		return createError("Не указано количество объединенных ячеек (атрибут count)");
	}

	const mergeCellTags = mergeCellsContent.match(/<mergeCell\s+ref="([A-Z]+\d+:[A-Z]+\d+)"\s*\/>/g);
	if (!mergeCellTags) {
		return createError("Не найдены объединенные ячейки");
	}

	// Проверка соответствия заявленного количества и фактического
	if (mergeCellTags.length !== parseInt(countMatch[1]!)) {
		return createError(
			"Несоответствие количества объединенных ячеек",
			`Ожидалось: ${countMatch[1]}, найдено: ${mergeCellTags.length}`,
		);
	}

	// Проверка на дублирующиеся mergeCell
	const mergeRefs = new Set<string>();
	const duplicates = new Set<string>();

	for (const mergeTag of mergeCellTags) {
		const refMatch = mergeTag.match(/ref="([A-Z]+\d+:[A-Z]+\d+)"/);
		if (!refMatch) {
			return createError("Некорректный формат объединения ячеек", `Тег: ${mergeTag}`);
		}

		const ref = refMatch[1];
		if (mergeRefs.has(ref!)) {
			duplicates.add(ref!);
		} else {
			mergeRefs.add(ref!);
		}
	}

	if (duplicates.size > 0) {
		return createError(
			"Найдены дублирующиеся объединения ячеек",
			`Дубликаты: ${Array.from(duplicates).join(", ")}`,
		);
	}

	// Проверка пересекающихся объединений
	const mergedRanges = Array.from(mergeRefs).map(ref => {
		const [start, end] = ref.split(":");
		return {
			endCol: end!.match(/[A-Z]+/)?.[0] || "",
			endRow: parseInt(end!.match(/\d+/)?.[0] || "0"),
			startCol: start!.match(/[A-Z]+/)?.[0] || "",
			startRow: parseInt(start!.match(/\d+/)?.[0] || "0"),
		};
	});

	for (let i = 0; i < mergedRanges.length; i++) {
		for (let j = i + 1; j < mergedRanges.length; j++) {
			const a = mergedRanges[i];
			const b = mergedRanges[j];

			if (rangesIntersect(a!, b!)) {
				return createError(
					"Найдены пересекающиеся объединения ячеек",
					`Пересекаются: ${getRangeString(a!)} и ${getRangeString(b!)}`,
				);
			}
		}
	}

	// 5. Проверка dimension и соответствия с реальными данными
	const dimensionMatch = xml.match(/<dimension\s+ref="([A-Z]+\d+:[A-Z]+\d+)"\s*\/>/);
	if (!dimensionMatch) {
		return createError("Не указана область данных (dimension)");
	}

	const [startCell, endCell] = dimensionMatch[1]!.split(":");
	const startCol = startCell!.match(/[A-Z]+/)?.[0];
	const startRow = parseInt(startCell!.match(/\d+/)?.[0] || "0");
	const endCol = endCell!.match(/[A-Z]+/)?.[0];
	const endRow = parseInt(endCell!.match(/\d+/)?.[0] || "0");

	if (!startCol || !endCol || isNaN(startRow) || isNaN(endRow)) {
		return createError("Некорректный формат dimension", `Dimension: ${dimensionMatch[1]}`);
	}

	const startColNum = colToNumber(startCol);
	const endColNum = colToNumber(endCol);

	// Проверяем все ячейки на вхождение в dimension
	for (const cell of allCells) {
		const colNum = colToNumber(cell.col);

		if (cell.row < startRow || cell.row > endRow) {
			return createError(
				"Ячейка находится вне указанной области (по строке)",
				`Ячейка: ${cell.col}${cell.row}, dimension: ${dimensionMatch[1]}`,
			);
		}

		if (colNum < startColNum || colNum > endColNum) {
			return createError(
				"Ячейка находится вне указанной области (по столбцу)",
				`Ячейка: ${cell.col}${cell.row}, dimension: ${dimensionMatch[1]}`,
			);
		}
	}

	// 6. Дополнительная проверка: все mergeCell ссылаются на существующие ячейки
	for (const mergeTag of mergeCellTags) {
		const refMatch = mergeTag.match(/ref="([A-Z]+\d+:[A-Z]+\d+)"/);
		if (!refMatch) {
			return createError("Некорректный формат объединения ячеек", `Тег: ${mergeTag}`);
		}

		const [cell1, cell2] = refMatch[1]!.split(":");
		const cell1Col = cell1!.match(/[A-Z]+/)?.[0];
		const cell1Row = parseInt(cell1!.match(/\d+/)?.[0] || "0");
		const cell2Col = cell2!.match(/[A-Z]+/)?.[0];
		const cell2Row = parseInt(cell2!.match(/\d+/)?.[0] || "0");

		if (!cell1Col || !cell2Col || isNaN(cell1Row) || isNaN(cell2Row)) {
			return createError("Некорректные координаты объединения ячеек", `Объединение: ${refMatch[1]}`);
		}

		// Проверяем что объединяемые ячейки существуют
		const cell1Exists = allCells.some(c => c.row === cell1Row && c.col === cell1Col);
		const cell2Exists = allCells.some(c => c.row === cell2Row && c.col === cell2Col);

		if (!cell1Exists || !cell2Exists) {
			return createError(
				"Объединение ссылается на несуществующие ячейки",
				`Объединение: ${refMatch[1]}, отсутствует: ${!cell1Exists ? `${cell1Col}${cell1Row}` : `${cell2Col}${cell2Row}`
				}`,
			);
		}
	}

	return { isValid: true };
}

// Вспомогательные функции для проверки пересечений
function rangesIntersect(a: { startCol: string; startRow: number; endCol: string; endRow: number },
	b: { startCol: string; startRow: number; endCol: string; endRow: number }): boolean {
	const aStartColNum = colToNumber(a.startCol);
	const aEndColNum = colToNumber(a.endCol);
	const bStartColNum = colToNumber(b.startCol);
	const bEndColNum = colToNumber(b.endCol);

	// Проверяем пересечение по строкам
	const rowsIntersect = !(a.endRow < b.startRow || a.startRow > b.endRow);

	// Проверяем пересечение по колонкам
	const colsIntersect = !(aEndColNum < bStartColNum || aStartColNum > bEndColNum);

	return rowsIntersect && colsIntersect;
}

function getRangeString(range: { startCol: string; startRow: number; endCol: string; endRow: number }): string {
	return `${range.startCol}${range.startRow}:${range.endCol}${range.endRow}`;
}

// Функция для преобразования букв колонки в число
function colToNumber(col: string): number {
	let num = 0;
	for (let i = 0; i < col.length; i++) {
		num = num * 26 + (col.charCodeAt(i) - 64);
	}
	return num;
};
