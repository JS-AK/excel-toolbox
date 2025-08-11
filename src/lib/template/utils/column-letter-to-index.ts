export function columnLetterToIndex(col: string): number {
	if (!col) return -1;

	let index = 0;

	for (let i = 0; i < col.length; i++) {
		const charCode = col.charCodeAt(i);
		if (charCode < 65 || charCode > 90) return -1; // не A-Z
		index = index * 26 + (charCode - 64); // 'A' -> 1
	}

	return index;
}
