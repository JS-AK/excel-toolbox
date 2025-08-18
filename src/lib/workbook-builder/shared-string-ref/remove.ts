import { WorkbookBuilder } from "../workbook-builder.js";

export function remove(
	this: WorkbookBuilder,
	payload: {
		sheetName: string;
		strIdx: number;
	},
): boolean {
	const { sheetName, strIdx } = payload;

	const str = this.sharedStrings[strIdx];
	if (!str) return false;

	const refs = this.sharedStringRefs.get(str);
	if (!refs) return false;

	refs.delete(sheetName);

	if (refs.size === 0) {
		// Строим карту старых индексов → новых до удаления
		const oldToNew = new Map<number, number>();
		for (let i = 0; i < this.sharedStrings.length; i++) {
			if (i < strIdx) oldToNew.set(i, i);
			else if (i > strIdx) oldToNew.set(i, i - 1);
			// i === strIdx — эта строка будет удалена, индекса нет
		}

		// Удаляем строку из массива и рефов
		this.sharedStrings.splice(strIdx, 1);
		this.sharedStringRefs.delete(str);

		// Обновляем индексы на всех листах
		for (const sheet of this.sheets.values()) {
			for (const row of sheet.rows.values()) {
				for (const cell of row.cells.values()) {
					if (cell.type === "s" && typeof cell.value === "number") {
						const newIdx = oldToNew.get(cell.value);
						if (newIdx !== undefined) {
							cell.value = newIdx;
						} else {
							// На всякий случай, если cell.value был удалённой строкой
							cell.value = 0; // или null, по логике твоего приложения
						}
					}
				}
			}
		}
	}

	return true;
}
