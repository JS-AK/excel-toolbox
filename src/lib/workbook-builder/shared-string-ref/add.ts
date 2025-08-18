import { WorkbookBuilder } from "../workbook-builder.js";

export function add(
	this: WorkbookBuilder,
	payload: {
		sheetName: string;
		str: string;
	},
): number {
	const { sheetName, str } = payload;

	let idx = this.sharedStrings.indexOf(str);

	if (idx === -1) {
		idx = this.sharedStrings.length;
		this.sharedStrings.push(str);
		this.sharedStringRefs.set(str, new Set([sheetName]));
	} else {
		// Добавляем имя листа в Set, если ещё нет
		this.sharedStringRefs.get(str)?.add(sheetName);
	}

	return idx;
}
