import { describe, expect, it } from "vitest";

import { buildMergedSheet } from "./build-merged-sheet.js";

describe("buildMergedSheet", () => {
	it("should merge rows into sheet XML", () => {
		const originalXml = `<?xml version="1.0" encoding="UTF-8"?>
			<worksheet>
				<sheetData>
					<row r="1"><c r="A1"><v>1</v></c></row>
				</sheetData>
			</worksheet>`;

		const mergedRows = [
			"<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>",
			"<row r=\"2\"><c r=\"A2\"><v>2</v></c></row>",
		];

		const result = buildMergedSheet(originalXml, mergedRows);
		const resultStr = result.toString();

		expect(resultStr).toContain("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>");
		expect(resultStr).toContain("<row r=\"2\"><c r=\"A2\"><v>2</v></c></row>");
		expect(resultStr).toContain("<sheetData>");
		expect(resultStr).toContain("</sheetData>");
		expect(resultStr).not.toContain("<mergeCells");
	});

	it("should add merge cells when provided", () => {
		const originalXml = `<?xml version="1.0" encoding="UTF-8"?>
			<worksheet>
				<sheetData>
					<row r="1"><c r="A1"><v>1</v></c></row>
				</sheetData>
			</worksheet>`;

		const mergedRows = [
			"<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>",
		];

		const mergeCells = [
			{ ref: "A1:B1" },
			{ ref: "C1:D1" },
		];

		const result = buildMergedSheet(originalXml, mergedRows, mergeCells);
		const resultStr = result.toString();

		expect(resultStr).toContain("<mergeCells count=\"2\">");
		expect(resultStr).toContain("<mergeCell ref=\"A1:B1\"/>");
		expect(resultStr).toContain("<mergeCell ref=\"C1:D1\"/>");
	});

	it("should replace existing merge cells", () => {
		const originalXml = `<?xml version="1.0" encoding="UTF-8"?>
			<worksheet>
				<sheetData>
					<row r="1"><c r="A1"><v>1</v></c></row>
				</sheetData>
				<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
			</worksheet>`;

		const mergedRows = [
			"<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>",
		];

		const mergeCells = [
			{ ref: "C1:D1" },
		];

		const result = buildMergedSheet(originalXml, mergedRows, mergeCells);
		const resultStr = result.toString();

		expect(resultStr).toContain("<mergeCells count=\"1\">");
		expect(resultStr).toContain("<mergeCell ref=\"C1:D1\"/>");
		expect(resultStr).not.toContain("<mergeCell ref=\"A1:B1\"/>");
	});
});
