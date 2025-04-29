import { describe, expect, it } from "vitest";

import { extractRowsFromSheetSync } from "./extract-rows-from-sheet-sync.js";

// Упрощённый шаблон XML-страницы Excel
const sampleSheet = `
	<?xml version="1.0" encoding="UTF-8"?>
	<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		<sheetData>
			<row r="1"><c r="A1" t="s"><v>0</v></c></row>
			<row r="2"><c r="A2" t="s"><v>1</v></c></row>
			<row r="5"><c r="A5" t="s"><v>2</v></c></row>
		</sheetData>
		<mergeCells count="2">
			<mergeCell ref="A1:B1"/>
			<mergeCell ref="A2:A3"/>
		</mergeCells>
	</worksheet>`;

const noRowsSheet = `
	<?xml version="1.0" encoding="UTF-8"?>
	<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		<sheetData></sheetData>
	</worksheet>`;

const noSheetData = `
	<?xml version="1.0" encoding="UTF-8"?>
	<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
	</worksheet>`;

describe("extractRowsFromSheet", () => {
	it("extracts rows and mergeCells correctly", () => {
		const result = extractRowsFromSheetSync(sampleSheet);

		expect(result.rows.length).toBe(3);
		expect(result.lastRowNumber).toBe(5);
		expect(result.rows[0]).toContain("<row r=\"1\">");
		expect(result.rows[2]).toContain("<row r=\"5\">");
		expect(result.mergeCells).toEqual([
			{ ref: "A1:B1" },
			{ ref: "A2:A3" },
		]);
		expect(result.xml).toContain("<worksheet");
	});

	it("returns empty row list and mergeCells if sheetData is empty", () => {
		const result = extractRowsFromSheetSync(noRowsSheet);
		expect(result.rows).toEqual([]);
		expect(result.lastRowNumber).toBe(0);
		expect(result.mergeCells).toEqual([]);
	});

	it("throws an error if sheetData is not found", () => {
		expect(() => extractRowsFromSheetSync(noSheetData)).toThrow("sheetData not found in worksheet XML");
	});

	it("accepts Buffer input", () => {
		const buffer = Buffer.from(sampleSheet, "utf-8");
		const result = extractRowsFromSheetSync(buffer);

		expect(result.rows.length).toBe(3);
		expect(result.lastRowNumber).toBe(5);
	});
});
