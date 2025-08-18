import { WorkbookBuilder } from "../workbook-builder.js";

export function removeAllFromSheet(
	this: WorkbookBuilder,
	payload: {
		sheetName: string;
	},
) {
	const { sheetName } = payload;

	// 1. Собираем строки, которые нужно удалить
	const stringsToRemove: string[] = [];

	for (const [str, sheetsSet] of this.sharedStringRefs) {
		sheetsSet.delete(sheetName);
		if (sheetsSet.size === 0) {
			stringsToRemove.push(str);
		}
	}

	if (stringsToRemove.length === 0) return;

	// 2. Строим карту старых индексов → новых
	const oldToNew = new Map<number, number>();
	let newIdx = 0;

	for (let oldIdx = 0; oldIdx < this.sharedStrings.length; oldIdx++) {
		const str = this.sharedStrings[oldIdx];
		if (!str) continue; // пропускаем, если undefined
		if (stringsToRemove.includes(str)) {
			// Удаляем строку из рефов
			this.sharedStringRefs.delete(str);
			continue; // индекс не учитывается
		}
		oldToNew.set(oldIdx, newIdx++);
	}

	// 3. Обновляем массив sharedStrings
	this.sharedStrings = this.sharedStrings.filter(s => !stringsToRemove.includes(s));

	// 4. Обновляем индексы в ячейках на всех листах
	for (const sheet of this.sheets.values()) {
		for (const row of sheet.rows.values()) {
			for (const cell of row.cells.values()) {
				if (cell.type === "s" && typeof cell.value === "number") {
					const newIdx = oldToNew.get(cell.value);
					if (newIdx !== undefined) {
						cell.value = newIdx;
					} else {
						// Если cell.value была удалённой строкой, можно поставить 0 или null
						cell.value = 0;
					}
				}
			}
		}
	}
}
