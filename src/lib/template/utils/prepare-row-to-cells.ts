import { escapeXml } from "../../utils/index.js";

import { columnIndexToLetter } from "./column-index-to-letter.js";

export function prepareRowToCells(row: unknown[], rowNumber: number) {
	return row.map((value, colIndex) => {
		const colLetter = columnIndexToLetter(colIndex);
		const cellRef = `${colLetter}${rowNumber}`;
		const cellValue = escapeXml(String(value ?? ""));

		return {
			cellRef,
			cellValue,
			cellXml: `<c r="${cellRef}" t="inlineStr"><is><t>${cellValue}</t></is></c>`,
		};
	});
}
