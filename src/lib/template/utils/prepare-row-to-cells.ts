import { columnIndexToLetter } from "./column-index-to-letter.js";
import { escapeXml } from "./escape-xml.js";

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
