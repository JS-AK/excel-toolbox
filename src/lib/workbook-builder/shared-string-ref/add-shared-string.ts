export function addSharedString(
	payload: {
		str: string;
		sheetName: string;
	},
	options: {
		sharedStrings: string[];
		sharedStringRefs: Map<string, Set<string>>;
	},
): number {
	const { sheetName, str } = payload;
	const { sharedStringRefs, sharedStrings } = options;

	let idx = sharedStrings.indexOf(str);

	if (idx === -1) {
		idx = sharedStrings.length;
		sharedStrings.push(str);
		sharedStringRefs.set(str, new Set([sheetName]));
	} else {
		// Добавляем имя листа в Set, если ещё нет
		sharedStringRefs.get(str)?.add(sheetName);
	}

	return idx;
}
